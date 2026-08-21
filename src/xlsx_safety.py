"""Resource-safety preflight for untrusted XLSX input packages."""

from __future__ import annotations

import ntpath
import posixpath
import zlib
from pathlib import Path, PurePosixPath
from urllib.parse import urlsplit
from xml.etree.ElementTree import Element, ParseError
from zipfile import ZIP_DEFLATED, ZIP_STORED, BadZipFile, ZipFile, ZipInfo

import openpyxl.xml
from defusedxml.ElementTree import iterparse
from defusedxml.common import DefusedXmlException
from openpyxl.xml.constants import (
    ARC_CONTENT_TYPES,
    PKG_REL_NS,
    REL_NS,
    SHEET_MAIN_NS,
    XLSX,
)
from openpyxl.utils.cell import coordinate_to_tuple


MEBIBYTE = 1024 * 1024
MAX_XLSX_FILE_SIZE = 10 * MEBIBYTE
MAX_ZIP_MEMBERS = 1_000
MAX_TOTAL_UNCOMPRESSED_SIZE = 100 * MEBIBYTE
MAX_INDIVIDUAL_MEMBER_SIZE = 50 * MEBIBYTE
MAX_COMPRESSION_RATIO = 100
MIN_INDIVIDUAL_RATIO_SIZE = 1 * MEBIBYTE
MIN_AGGREGATE_RATIO_SIZE = 10 * MEBIBYTE
MAX_WORKSHEETS = 50
MAX_ROWS_PER_WORKSHEET = 100_000
MAX_COLUMNS_PER_WORKSHEET = 256
MAX_TOTAL_LOGICAL_CELLS = 1_000_000
CRC_READ_CHUNK_SIZE = 64 * 1024

_ENCRYPTED_FLAG = 0x1
_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_CONTENT_TYPES_ROOT_TAG = f"{{{_CONTENT_TYPES_NS}}}Types"
_CONTENT_TYPES_OVERRIDE_TAG = (
    f"{{{_CONTENT_TYPES_NS}}}Override"
)
_RELATIONSHIPS_ROOT_TAG = f"{{{PKG_REL_NS}}}Relationships"
_RELATIONSHIP_TAG = f"{{{PKG_REL_NS}}}Relationship"
_WORKSHEET_RELATIONSHIP_TYPE = f"{REL_NS}/worksheet"
_CHARTSHEET_RELATIONSHIP_TYPE = f"{REL_NS}/chartsheet"
_WORKBOOK_ROOT_TAG = f"{{{SHEET_MAIN_NS}}}workbook"
_WORKBOOK_SHEET_TAG = f"{{{SHEET_MAIN_NS}}}sheet"
_RELATIONSHIP_ID_ATTRIBUTE = f"{{{REL_NS}}}id"
_SUPPORTED_COMPRESSION_METHODS = frozenset({ZIP_STORED, ZIP_DEFLATED})


class UnsafeXlsxPackageError(Exception):
    """Raised when an XLSX package violates an input-safety invariant."""


def _validate_xml_protection() -> None:
    if openpyxl.xml.DEFUSEDXML is not True:
        raise UnsafeXlsxPackageError(
            "required defusedxml protection is not active in openpyxl"
        )


def _member_source_name(member: ZipInfo) -> str:
    return getattr(member, "orig_filename", member.filename)


def _validate_member_name(member: ZipInfo) -> None:
    name = _member_source_name(member)
    path = PurePosixPath(name)
    if "\x00" in name:
        raise UnsafeXlsxPackageError("ZIP member name contains a NUL character")
    if path.is_absolute() or ntpath.isabs(name):
        raise UnsafeXlsxPackageError(f"ZIP member name is absolute: {name!r}")
    if "\\" in name:
        raise UnsafeXlsxPackageError(
            f"ZIP member name contains a backslash: {name!r}"
        )
    if ".." in path.parts:
        raise UnsafeXlsxPackageError(
            f"ZIP member name contains parent traversal: {name!r}"
        )


def _ratio_exceeds_limit(uncompressed: int, compressed: int) -> bool:
    if uncompressed == 0:
        return False
    if compressed == 0:
        return True
    return uncompressed > compressed * MAX_COMPRESSION_RATIO


def _aggregate_ratio_exceeds_limit(
    total_uncompressed: int,
    total_compressed: int,
) -> bool:
    return (
        total_uncompressed > MIN_AGGREGATE_RATIO_SIZE
        and _ratio_exceeds_limit(total_uncompressed, total_compressed)
    )


def _member_ratio_exceeds_limit(
    member: ZipInfo,
    total_uncompressed: int,
) -> bool:
    return (
        member.file_size > MIN_INDIVIDUAL_RATIO_SIZE
        or total_uncompressed > MIN_AGGREGATE_RATIO_SIZE
    ) and _ratio_exceeds_limit(member.file_size, member.compress_size)


def _workbook_relationships_path(workbook_part: str) -> str:
    folder, filename = posixpath.split(workbook_part)
    relationships = f"_rels/{filename}.rels"
    return posixpath.join(folder, relationships) if folder else relationships


def _find_xlsx_workbook_part(archive: ZipFile, members: set[str]) -> str:
    if ARC_CONTENT_TYPES not in members:
        raise UnsafeXlsxPackageError(
            f"XLSX package is missing required part {ARC_CONTENT_TYPES!r}"
        )

    workbook_part: str | None = None
    root_seen = False
    try:
        with archive.open(ARC_CONTENT_TYPES, "r") as content_types:
            for event, element in iterparse(content_types, events=("start", "end")):
                if not root_seen:
                    root_seen = True
                    if event != "start" or element.tag != _CONTENT_TYPES_ROOT_TAG:
                        raise UnsafeXlsxPackageError(
                            "XLSX content-types manifest has an invalid root element"
                        )
                if event != "end":
                    continue
                if (
                    element.tag == _CONTENT_TYPES_OVERRIDE_TAG
                    and element.attrib.get("ContentType") == XLSX
                ):
                    part_name = element.attrib.get("PartName")
                    if not part_name or not part_name.startswith("/"):
                        raise UnsafeXlsxPackageError(
                            "XLSX content-types manifest contains an invalid "
                            "workbook part name"
                        )
                    if workbook_part is not None:
                        raise UnsafeXlsxPackageError(
                            "XLSX content-types manifest must identify exactly one "
                            "workbook part"
                        )
                    workbook_part = part_name[1:]
                element.clear()
    except (DefusedXmlException, ParseError) as error:
        raise UnsafeXlsxPackageError(
            "XLSX content-types manifest contains prohibited or malformed XML"
        ) from error

    if workbook_part is None:
        raise UnsafeXlsxPackageError(
            "XLSX content-types manifest must identify exactly one workbook part"
        )

    if workbook_part not in members:
        raise UnsafeXlsxPackageError(
            f"XLSX package is missing declared workbook part {workbook_part!r}"
        )

    relationships_part = _workbook_relationships_path(workbook_part)
    if relationships_part not in members:
        raise UnsafeXlsxPackageError(
            "XLSX package is missing required workbook relationships part "
            f"{relationships_part!r}"
        )
    return workbook_part


def _resolve_relationship_target(workbook_part: str, target: str) -> str:
    if not target or "\x00" in target or "\\" in target:
        raise UnsafeXlsxPackageError(
            f"workbook worksheet relationship has an invalid target: {target!r}"
        )

    try:
        parsed = urlsplit(target)
    except ValueError as error:
        raise UnsafeXlsxPackageError(
            f"workbook worksheet relationship has an invalid target: {target!r}"
        ) from error
    if parsed.scheme or parsed.netloc or parsed.query or parsed.fragment:
        raise UnsafeXlsxPackageError(
            f"workbook worksheet relationship has an external or invalid target: "
            f"{target!r}"
        )

    if target.startswith("//"):
        raise UnsafeXlsxPackageError(
            f"workbook worksheet relationship has an external or invalid target: "
            f"{target!r}"
        )
    if target.startswith("/"):
        candidate = target[1:]
    else:
        candidate = posixpath.join(posixpath.dirname(workbook_part), target)

    normalized = posixpath.normpath(candidate)
    if (
        normalized in {"", ".", ".."}
        or normalized.startswith("../")
        or normalized.startswith("/")
    ):
        raise UnsafeXlsxPackageError(
            f"workbook worksheet relationship escapes the package: {target!r}"
        )
    return normalized


def _find_workbook_sheet_relationship_ids(
    archive: ZipFile,
    workbook_part: str,
) -> list[str]:
    relationship_ids: list[str] = []
    seen_relationship_ids: set[str] = set()
    root_seen = False
    try:
        with archive.open(workbook_part, "r") as workbook:
            for event, element in iterparse(workbook, events=("start", "end")):
                if not root_seen:
                    root_seen = True
                    if event != "start" or element.tag != _WORKBOOK_ROOT_TAG:
                        raise UnsafeXlsxPackageError(
                            "XLSX workbook part has an invalid root element"
                        )
                if event != "end":
                    continue
                if element.tag == _WORKBOOK_SHEET_TAG:
                    relationship_id = element.attrib.get(_RELATIONSHIP_ID_ATTRIBUTE)
                    if not relationship_id:
                        raise UnsafeXlsxPackageError(
                            "XLSX workbook sheet is missing its relationship ID"
                        )
                    if relationship_id in seen_relationship_ids:
                        raise UnsafeXlsxPackageError(
                            "XLSX workbook contains duplicate sheet relationship IDs: "
                            f"{relationship_id!r}"
                        )
                    seen_relationship_ids.add(relationship_id)
                    relationship_ids.append(relationship_id)
                    if len(relationship_ids) > MAX_ZIP_MEMBERS:
                        raise UnsafeXlsxPackageError(
                            "XLSX workbook contains more sheet declarations than "
                            "the package member limit"
                        )
                element.clear()
    except (DefusedXmlException, ParseError) as error:
        raise UnsafeXlsxPackageError(
            "XLSX workbook part contains prohibited or malformed XML"
        ) from error
    return relationship_ids


def _find_worksheet_parts(
    archive: ZipFile,
    workbook_part: str,
    members: dict[str, ZipInfo],
) -> list[ZipInfo]:
    relationships_part = _workbook_relationships_path(workbook_part)
    sheet_relationship_ids = _find_workbook_sheet_relationship_ids(
        archive,
        workbook_part,
    )
    relationships_by_id: dict[str, Element] = {}
    wanted_relationship_ids = set(sheet_relationship_ids)
    worksheet_parts: list[ZipInfo] = []
    seen_targets: set[str] = set()
    root_seen = False

    try:
        with archive.open(relationships_part, "r") as relationships:
            for event, element in iterparse(
                relationships,
                events=("start", "end"),
            ):
                if not root_seen:
                    root_seen = True
                    if event != "start" or element.tag != _RELATIONSHIPS_ROOT_TAG:
                        raise UnsafeXlsxPackageError(
                            "XLSX workbook relationships has an invalid root element"
                        )
                if event != "end":
                    continue
                if element.tag == _RELATIONSHIP_TAG:
                    relationship_id = element.attrib.get("Id")
                    if not relationship_id:
                        raise UnsafeXlsxPackageError(
                            "XLSX workbook relationship is missing its ID"
                        )
                    if relationship_id not in wanted_relationship_ids:
                        element.clear()
                        continue
                    if relationship_id in relationships_by_id:
                        raise UnsafeXlsxPackageError(
                            "XLSX workbook contains duplicate relationship IDs: "
                            f"{relationship_id!r}"
                        )
                    relationships_by_id[relationship_id] = Element(
                        element.tag,
                        element.attrib,
                    )
                element.clear()
    except (DefusedXmlException, ParseError) as error:
        raise UnsafeXlsxPackageError(
            "XLSX workbook relationships contains prohibited or malformed XML"
        ) from error

    for relationship_id in sheet_relationship_ids:
        try:
            relationship = relationships_by_id[relationship_id]
        except KeyError as error:
            raise UnsafeXlsxPackageError(
                "XLSX workbook sheet references a missing relationship: "
                f"{relationship_id!r}"
            ) from error
        relationship_type = relationship.attrib.get("Type")
        if relationship_type not in {
            _WORKSHEET_RELATIONSHIP_TYPE,
            _CHARTSHEET_RELATIONSHIP_TYPE,
        }:
            raise UnsafeXlsxPackageError(
                "XLSX workbook sheet relationship has an invalid type: "
                f"{relationship_type!r}"
            )
        target = relationship.attrib.get("Target", "")
        target_mode = relationship.attrib.get("TargetMode")
        if target_mode == "External":
            raise UnsafeXlsxPackageError(
                "XLSX workbook contains an external worksheet relationship"
            )
        if target_mode not in {None, "Internal"}:
            raise UnsafeXlsxPackageError(
                "XLSX workbook worksheet relationship has an invalid "
                f"TargetMode: {target_mode!r}"
            )
        resolved_target = _resolve_relationship_target(workbook_part, target)
        if resolved_target in seen_targets:
            raise UnsafeXlsxPackageError(
                "XLSX workbook contains duplicate worksheet targets: "
                f"{resolved_target!r}"
            )
        seen_targets.add(resolved_target)
        try:
            target_member = members[resolved_target]
        except KeyError as error:
            raise UnsafeXlsxPackageError(
                "XLSX workbook sheet relationship targets a missing "
                f"part: {resolved_target!r}"
            ) from error
        if relationship_type == _WORKSHEET_RELATIONSHIP_TYPE:
            worksheet_parts.append(target_member)
            if len(worksheet_parts) > MAX_WORKSHEETS:
                raise UnsafeXlsxPackageError(
                    f"XLSX package exceeds the worksheet limit of {MAX_WORKSHEETS}"
                )
    return worksheet_parts


def _validate_zip_integrity(archive: ZipFile, members: list[ZipInfo]) -> None:
    for member in members:
        if member.is_dir():
            continue
        with archive.open(member, "r") as source:
            while source.read(CRC_READ_CHUNK_SIZE):
                pass


def _inspect_worksheet_xml(archive: ZipFile, member: ZipInfo) -> int:
    max_row = 0
    max_column = 0
    sheet_data: Element | None = None
    current_row = 0
    current_column = 0
    row_is_active = False
    sheet_data_tag = f"{{{SHEET_MAIN_NS}}}sheetData"
    row_tag = f"{{{SHEET_MAIN_NS}}}row"
    cell_tag = f"{{{SHEET_MAIN_NS}}}c"
    try:
        with archive.open(member, "r") as worksheet_xml:
            for event, element in iterparse(
                worksheet_xml,
                events=("start", "end"),
            ):
                if event == "start":
                    if element.tag == sheet_data_tag and sheet_data is None:
                        sheet_data = element
                    elif element.tag == row_tag and sheet_data is not None:
                        row_reference = element.attrib.get("r")
                        if row_reference is None:
                            current_row += 1
                        else:
                            try:
                                current_row = int(row_reference)
                            except ValueError as error:
                                raise UnsafeXlsxPackageError(
                                    f"worksheet member {member.filename!r} contains "
                                    f"an invalid row reference: {row_reference!r}"
                                ) from error
                            if current_row < 1:
                                raise UnsafeXlsxPackageError(
                                    f"worksheet member {member.filename!r} contains "
                                    f"an invalid row reference: {row_reference!r}"
                                )
                        current_column = 0
                        row_is_active = True
                        max_row = max(max_row, current_row)
                    continue

                if element.tag == cell_tag and sheet_data is not None and row_is_active:
                    coordinate = element.attrib.get("r")
                    if coordinate is None:
                        current_column += 1
                        row_index = current_row
                        column_index = current_column
                    else:
                        if not coordinate:
                            raise UnsafeXlsxPackageError(
                                f"worksheet member {member.filename!r} contains "
                                "an empty cell reference"
                            )
                        try:
                            row_index, column_index = coordinate_to_tuple(coordinate)
                        except (TypeError, ValueError) as error:
                            raise UnsafeXlsxPackageError(
                                f"worksheet member {member.filename!r} contains "
                                f"an invalid cell reference: {coordinate!r}"
                            ) from error
                        if row_index < 1 or column_index < 1:
                            raise UnsafeXlsxPackageError(
                                f"worksheet member {member.filename!r} contains "
                                f"an invalid cell reference: {coordinate!r}"
                            )
                        current_column = column_index
                    max_row = max(max_row, row_index)
                    max_column = max(max_column, column_index)

                element.clear()
                if element.tag == row_tag and sheet_data is not None:
                    row_is_active = False
                    sheet_data.clear()
                elif element is sheet_data:
                    sheet_data = None
    except (DefusedXmlException, ParseError) as error:
        raise UnsafeXlsxPackageError(
            f"worksheet member {member.filename!r} contains prohibited or malformed XML"
        ) from error

    if max_row > MAX_ROWS_PER_WORKSHEET:
        raise UnsafeXlsxPackageError(
            f"worksheet member {member.filename!r} exceeds the row limit "
            f"of {MAX_ROWS_PER_WORKSHEET}"
        )
    if max_column > MAX_COLUMNS_PER_WORKSHEET:
        raise UnsafeXlsxPackageError(
            f"worksheet member {member.filename!r} exceeds the column limit "
            f"of {MAX_COLUMNS_PER_WORKSHEET}"
        )
    return max_row * max_column


def inspect_xlsx_package(file_path: Path) -> None:
    """Reject unsafe XLSX package structures before openpyxl loads them."""
    _validate_xml_protection()

    try:
        file_size = file_path.stat().st_size
    except OSError as error:
        raise UnsafeXlsxPackageError(
            f"XLSX file could not be inspected: {error}"
        ) from error
    if file_size > MAX_XLSX_FILE_SIZE:
        raise UnsafeXlsxPackageError(
            f"XLSX file exceeds the size limit of {MAX_XLSX_FILE_SIZE} bytes"
        )

    try:
        with ZipFile(file_path, "r") as archive:
            members = archive.infolist()
            if not members:
                raise UnsafeXlsxPackageError("XLSX package contains no members")
            if len(members) > MAX_ZIP_MEMBERS:
                raise UnsafeXlsxPackageError(
                    f"XLSX package exceeds the member limit of {MAX_ZIP_MEMBERS}"
                )

            seen_names: set[str] = set()
            total_compressed = 0
            total_uncompressed = 0
            ratio_candidates: list[ZipInfo] = []
            members_by_name: dict[str, ZipInfo] = {}

            for member in members:
                _validate_member_name(member)
                name = _member_source_name(member)
                if name in seen_names:
                    raise UnsafeXlsxPackageError(
                        f"XLSX package contains a duplicate member name: {name!r}"
                    )
                seen_names.add(name)
                members_by_name[name] = member

                if member.flag_bits & _ENCRYPTED_FLAG:
                    raise UnsafeXlsxPackageError(
                        f"XLSX package contains an encrypted member: {name!r}"
                    )
                if member.compress_type not in _SUPPORTED_COMPRESSION_METHODS:
                    raise UnsafeXlsxPackageError(
                        f"ZIP member {name!r} uses unsupported compression method "
                        f"{member.compress_type}"
                    )
                if member.file_size > MAX_INDIVIDUAL_MEMBER_SIZE:
                    raise UnsafeXlsxPackageError(
                        f"ZIP member {name!r} exceeds the individual size limit "
                        f"of {MAX_INDIVIDUAL_MEMBER_SIZE} bytes"
                    )
                if member.file_size > 0 and member.compress_size == 0:
                    raise UnsafeXlsxPackageError(
                        f"ZIP member {name!r} has zero compressed size for "
                        "non-empty content"
                    )

                total_compressed += member.compress_size
                total_uncompressed += member.file_size
                ratio_candidates.append(member)

            if total_compressed > MAX_XLSX_FILE_SIZE:
                raise UnsafeXlsxPackageError(
                    "XLSX package compressed member data exceeds the file-size limit"
                )
            if total_uncompressed > MAX_TOTAL_UNCOMPRESSED_SIZE:
                raise UnsafeXlsxPackageError(
                    "XLSX package exceeds the total uncompressed-size limit "
                    f"of {MAX_TOTAL_UNCOMPRESSED_SIZE} bytes"
                )

            if _aggregate_ratio_exceeds_limit(
                total_uncompressed,
                total_compressed,
            ):
                raise UnsafeXlsxPackageError(
                    "XLSX package exceeds the aggregate compression ratio limit "
                    f"of {MAX_COMPRESSION_RATIO}:1"
                )

            for member in ratio_candidates:
                if _member_ratio_exceeds_limit(member, total_uncompressed):
                    raise UnsafeXlsxPackageError(
                        f"ZIP member {member.filename!r} exceeds the compression "
                        f"ratio limit of {MAX_COMPRESSION_RATIO}:1"
                    )

            workbook_part = _find_xlsx_workbook_part(archive, seen_names)
            worksheet_members = _find_worksheet_parts(
                archive,
                workbook_part,
                members_by_name,
            )
            _validate_zip_integrity(archive, members)

            if len(worksheet_members) > MAX_WORKSHEETS:
                raise UnsafeXlsxPackageError(
                    f"XLSX package exceeds the worksheet limit of {MAX_WORKSHEETS}"
                )

            total_logical_cells = 0
            for member in worksheet_members:
                total_logical_cells += _inspect_worksheet_xml(archive, member)
                if total_logical_cells > MAX_TOTAL_LOGICAL_CELLS:
                    raise UnsafeXlsxPackageError(
                        "XLSX package exceeds the total logical-cell limit "
                        f"of {MAX_TOTAL_LOGICAL_CELLS}"
                    )
    except BadZipFile as error:
        raise UnsafeXlsxPackageError(
            "XLSX package is not a valid ZIP archive"
        ) from error
    except OSError as error:
        raise UnsafeXlsxPackageError(
            f"XLSX package could not be opened or read: {error}"
        ) from error
    except (EOFError, NotImplementedError, zlib.error) as error:
        raise UnsafeXlsxPackageError(
            f"XLSX package contains unreadable compressed data: {error}"
        ) from error


def validate_loaded_workbook_limits(workbook: object) -> None:
    """Recheck limits that depend on openpyxl's final workbook model."""
    worksheets = workbook.worksheets
    if len(worksheets) > MAX_WORKSHEETS:
        raise UnsafeXlsxPackageError(
            f"XLSX workbook exceeds the worksheet limit of {MAX_WORKSHEETS}"
        )

    total_logical_cells = 0
    for worksheet in worksheets:
        max_row = worksheet.max_row
        max_column = worksheet.max_column
        if max_row > MAX_ROWS_PER_WORKSHEET:
            raise UnsafeXlsxPackageError(
                f"worksheet {worksheet.title!r} exceeds the row limit "
                f"of {MAX_ROWS_PER_WORKSHEET}"
            )
        if max_column > MAX_COLUMNS_PER_WORKSHEET:
            raise UnsafeXlsxPackageError(
                f"worksheet {worksheet.title!r} exceeds the column limit "
                f"of {MAX_COLUMNS_PER_WORKSHEET}"
            )
        total_logical_cells += max_row * max_column
        if total_logical_cells > MAX_TOTAL_LOGICAL_CELLS:
            raise UnsafeXlsxPackageError(
                "XLSX workbook exceeds the total logical-cell limit "
                f"of {MAX_TOTAL_LOGICAL_CELLS}"
            )

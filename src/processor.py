"""Input discovery and structural loading for sales data sources."""

from __future__ import annotations

import csv
from collections.abc import Sequence
from pathlib import Path
from typing import TypeAlias


REQUIRED_COLUMNS = frozenset(
    {"order_id", "customer_name", "order_date", "amount", "status"}
)
RESERVED_TRACEABILITY_COLUMNS = frozenset(
    {"source_file", "source_sheet", "source_row"}
)
SUPPORTED_CSV_SUFFIX = ".csv"

CsvValue: TypeAlias = str | None
LoadedRecord: TypeAlias = dict[str, CsvValue | int]


class StructuralInputError(Exception):
    """Base exception for input failures that prevent a successful run."""


class InputDirectoryNotFoundError(StructuralInputError):
    """Raised when the configured input directory does not exist."""


class NoSupportedCsvFilesError(StructuralInputError):
    """Raised when the input directory contains no supported CSV files."""


class InputDirectoryAccessError(StructuralInputError):
    """Raised when an operational filesystem error prevents input discovery."""

    def __init__(self, input_dir: Path, detail: str) -> None:
        self.input_dir = input_dir
        self.detail = detail
        super().__init__(f"Could not access input directory '{input_dir}': {detail}")


class CsvStructuralError(StructuralInputError):
    """Raised when a CSV source does not satisfy the structural contract."""

    def __init__(self, file_path: Path, detail: str) -> None:
        self.file_path = file_path
        self.detail = detail
        super().__init__(f"CSV structural error in '{file_path}': {detail}")


def discover_csv_files(input_dir: Path) -> list[Path]:
    """Return supported CSV files directly inside *input_dir* in stable order."""
    try:
        if not input_dir.is_dir():
            raise InputDirectoryNotFoundError(
                f"Input directory does not exist: {input_dir}"
            )

        csv_files = sorted(
            (
                path
                for path in input_dir.iterdir()
                if not path.is_symlink()
                and path.is_file()
                and path.suffix.lower() == SUPPORTED_CSV_SUFFIX
            ),
            key=lambda path: (path.name.lower(), path.name),
        )
    except OSError as error:
        raise InputDirectoryAccessError(input_dir, str(error)) from error

    if not csv_files:
        raise NoSupportedCsvFilesError(
            f"No supported CSV files found in: {input_dir}"
        )

    return csv_files


def _detect_non_comma_delimiter(sample: str, file_path: Path) -> None:
    """Reject a detectable alternative delimiter without using it for loading."""
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        return

    if dialect.delimiter != ",":
        raise CsvStructuralError(
            file_path,
            f"expected comma delimiter, detected {dialect.delimiter!r}",
        )


def _validate_header(header: list[str], file_path: Path) -> None:
    """Validate the exact, case-sensitive CSV header contract."""
    if not header or not any(header):
        raise CsvStructuralError(file_path, "missing header")

    duplicate_columns = sorted(
        {column for column in header if header.count(column) > 1}
    )
    if duplicate_columns:
        raise CsvStructuralError(
            file_path,
            f"duplicate header columns: {', '.join(duplicate_columns)}",
        )

    reserved_columns = sorted(RESERVED_TRACEABILITY_COLUMNS.intersection(header))
    if reserved_columns:
        raise CsvStructuralError(
            file_path,
            f"reserved header columns: {', '.join(reserved_columns)}",
        )

    missing_columns = sorted(REQUIRED_COLUMNS.difference(header))
    if missing_columns:
        raise CsvStructuralError(
            file_path,
            f"missing required columns: {', '.join(missing_columns)}",
        )


def load_csv_file(file_path: Path) -> list[LoadedRecord]:
    """Validate and load one CSV file with source traceability metadata."""
    try:
        with file_path.open(encoding="utf-8-sig", errors="strict", newline="") as file:
            sample = file.read(8192)
            if not sample:
                raise CsvStructuralError(file_path, "missing header")

            _detect_non_comma_delimiter(sample, file_path)
            file.seek(0)
            reader = csv.reader(file, delimiter=",", strict=True)

            try:
                header = next(reader)
            except StopIteration as error:
                raise CsvStructuralError(file_path, "missing header") from error

            _validate_header(header, file_path)
            records: list[LoadedRecord] = []

            while True:
                source_row = reader.line_num + 1
                try:
                    row = next(reader)
                except StopIteration:
                    break

                if len(row) > len(header):
                    raise CsvStructuralError(
                        file_path,
                        f"row {source_row} contains more values than the header",
                    )

                values: list[CsvValue] = [*row, *([None] * (len(header) - len(row)))]
                record: LoadedRecord = dict(zip(header, values, strict=True))
                record["source_file"] = file_path.name
                record["source_sheet"] = None
                record["source_row"] = source_row
                records.append(record)

            if not records:
                raise CsvStructuralError(file_path, "file contains no data records")

            return records
    except UnicodeDecodeError as error:
        raise CsvStructuralError(file_path, "file is not valid UTF-8") from error
    except csv.Error as error:
        raise CsvStructuralError(file_path, f"malformed CSV: {error}") from error
    except OSError as error:
        raise CsvStructuralError(file_path, f"file could not be read: {error}") from error


def load_csv_files(file_paths: Sequence[Path]) -> list[LoadedRecord]:
    """Load all CSV sources, failing the complete operation on any source error."""
    records: list[LoadedRecord] = []
    for file_path in file_paths:
        records.extend(load_csv_file(file_path))
    return records

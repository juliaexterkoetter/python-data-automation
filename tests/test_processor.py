from __future__ import annotations

from pathlib import Path

import pytest

from src.processor import (
    CsvStructuralError,
    InputDirectoryAccessError,
    InputDirectoryNotFoundError,
    NoSupportedCsvFilesError,
    discover_csv_files,
    load_csv_file,
    load_csv_files,
    process_records,
)


VALID_HEADER = "order_id,customer_name,email,order_date,amount,status"
VALID_ROW = "00123,Julia,julia@example.com,2026-08-20,149.90,paid"


def write_csv(path: Path, content: str, *, encoding: str = "utf-8") -> Path:
    path.write_text(content, encoding=encoding)
    return path


def test_discover_csv_files_fails_when_input_directory_does_not_exist(
    tmp_path: Path,
) -> None:
    with pytest.raises(InputDirectoryNotFoundError, match="does not exist"):
        discover_csv_files(tmp_path / "missing")


def test_discover_csv_files_fails_when_no_supported_csv_exists(
    tmp_path: Path,
) -> None:
    write_csv(tmp_path / "orders.txt", "ignored")

    with pytest.raises(NoSupportedCsvFilesError, match="No supported CSV"):
        discover_csv_files(tmp_path)


def test_discover_csv_files_is_case_insensitive_non_recursive_and_sorted(
    tmp_path: Path,
) -> None:
    write_csv(tmp_path / "b.CSV", "content")
    write_csv(tmp_path / "A.cSv", "content")
    write_csv(tmp_path / "ignored.xlsx", "content")
    nested_dir = tmp_path / "nested"
    nested_dir.mkdir()
    write_csv(nested_dir / "nested.csv", "content")
    csv_named_directory = tmp_path / "directory.csv"
    csv_named_directory.mkdir()

    discovered = discover_csv_files(tmp_path)

    assert [path.name for path in discovered] == ["A.cSv", "b.CSV"]


def test_discover_csv_files_ignores_symlinks_without_reading_the_target(
    tmp_path: Path,
) -> None:
    external_file = write_csv(
        tmp_path.parent / "external-invalid.csv",
        "this target must not be read",
    )
    (tmp_path / "linked.csv").symlink_to(external_file)
    regular_file = write_csv(
        tmp_path / "orders.csv",
        f"{VALID_HEADER}\n{VALID_ROW}\n",
    )

    discovered = discover_csv_files(tmp_path)
    records = load_csv_files(discovered)

    assert discovered == [regular_file]
    assert records[0]["source_file"] == "orders.csv"


def test_discover_csv_files_fails_when_only_csv_symlinks_exist(
    tmp_path: Path,
) -> None:
    external_file = write_csv(
        tmp_path.parent / "external.csv",
        f"{VALID_HEADER}\n{VALID_ROW}\n",
    )
    (tmp_path / "linked.csv").symlink_to(external_file)

    with pytest.raises(NoSupportedCsvFilesError, match="No supported CSV"):
        discover_csv_files(tmp_path)


def test_discover_csv_files_wraps_filesystem_errors_and_preserves_cause(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    permission_error = PermissionError("access denied")

    def raise_permission_error(_path: Path):
        raise permission_error

    monkeypatch.setattr(Path, "iterdir", raise_permission_error)

    with pytest.raises(InputDirectoryAccessError) as captured:
        discover_csv_files(tmp_path)

    assert captured.value.input_dir == tmp_path
    assert "access denied" in str(captured.value)
    assert captured.value.__cause__ is permission_error


def test_load_csv_file_accepts_utf8_bom_and_preserves_order_id(
    tmp_path: Path,
) -> None:
    file_path = write_csv(
        tmp_path / "orders.csv",
        f"{VALID_HEADER}\n{VALID_ROW}\n",
        encoding="utf-8-sig",
    )

    records = load_csv_file(file_path)

    assert records[0]["order_id"] == "00123"
    assert "\ufefforder_id" not in records[0]


def test_load_csv_file_rejects_non_utf8_bytes(tmp_path: Path) -> None:
    file_path = tmp_path / "orders.csv"
    file_path.write_bytes(
        f"{VALID_HEADER}\n00123,Andr".encode("utf-8")
        + b"\xe9,customer@example.com,2026-08-20,10.00,paid\n"
    )

    with pytest.raises(CsvStructuralError, match="not valid UTF-8"):
        load_csv_file(file_path)


@pytest.mark.parametrize("content", ["", "\n"])
def test_load_csv_file_rejects_missing_header(
    tmp_path: Path, content: str
) -> None:
    file_path = write_csv(tmp_path / "orders.csv", content)

    with pytest.raises(CsvStructuralError, match="missing header"):
        load_csv_file(file_path)


def test_load_csv_file_rejects_header_without_data_records(tmp_path: Path) -> None:
    file_path = write_csv(tmp_path / "orders.csv", f"{VALID_HEADER}\n")

    with pytest.raises(CsvStructuralError, match="file contains no data records"):
        load_csv_file(file_path)


@pytest.mark.parametrize("delimiter", [";", "\t", "|"])
def test_load_csv_file_rejects_non_comma_delimiters(
    tmp_path: Path,
    delimiter: str,
) -> None:
    header = delimiter.join(VALID_HEADER.split(","))
    row = delimiter.join(VALID_ROW.split(","))
    file_path = write_csv(tmp_path / "orders.csv", f"{header}\n{row}\n")

    with pytest.raises(CsvStructuralError, match="expected comma delimiter"):
        load_csv_file(file_path)


@pytest.mark.parametrize(
    "replacement",
    ["Order_ID", " order_id", "order_id ", "ORDER_ID"],
)
def test_load_csv_file_validates_header_names_exactly(
    tmp_path: Path,
    replacement: str,
) -> None:
    header = VALID_HEADER.replace("order_id", replacement, 1)
    file_path = write_csv(tmp_path / "orders.csv", f"{header}\n{VALID_ROW}\n")

    with pytest.raises(CsvStructuralError, match="missing required columns: order_id"):
        load_csv_file(file_path)


@pytest.mark.parametrize(
    "missing_column",
    ["order_id", "customer_name", "order_date", "amount", "status"],
)
def test_load_csv_file_rejects_each_missing_required_column(
    tmp_path: Path,
    missing_column: str,
) -> None:
    columns = VALID_HEADER.split(",")
    values = VALID_ROW.split(",")
    index = columns.index(missing_column)
    columns.pop(index)
    values.pop(index)
    file_path = write_csv(
        tmp_path / "orders.csv",
        f"{','.join(columns)}\n{','.join(values)}\n",
    )

    with pytest.raises(CsvStructuralError, match=missing_column):
        load_csv_file(file_path)


@pytest.mark.parametrize("reserved", ["source_file", "source_sheet", "source_row"])
def test_load_csv_file_rejects_reserved_traceability_columns(
    tmp_path: Path,
    reserved: str,
) -> None:
    file_path = write_csv(
        tmp_path / "orders.csv",
        f"{VALID_HEADER},{reserved}\n{VALID_ROW},client value\n",
    )

    with pytest.raises(CsvStructuralError, match=f"reserved header columns: {reserved}"):
        load_csv_file(file_path)


def test_load_csv_file_rejects_duplicate_header_columns(tmp_path: Path) -> None:
    file_path = write_csv(
        tmp_path / "orders.csv",
        f"{VALID_HEADER},status\n{VALID_ROW},paid\n",
    )

    with pytest.raises(CsvStructuralError, match="duplicate header columns: status"):
        load_csv_file(file_path)


def test_load_csv_file_preserves_extra_columns_and_adds_traceability(
    tmp_path: Path,
) -> None:
    file_path = write_csv(
        tmp_path / "orders.csv",
        f"{VALID_HEADER},sales_channel\n{VALID_ROW},online\n"
        "00124,Ada,,2026-08-21,0,pending,store\n",
    )

    records = load_csv_file(file_path)

    assert records == [
        {
            "order_id": "00123",
            "customer_name": "Julia",
            "email": "julia@example.com",
            "order_date": "2026-08-20",
            "amount": "149.90",
            "status": "paid",
            "sales_channel": "online",
            "source_file": "orders.csv",
            "source_sheet": None,
            "source_row": 2,
        },
        {
            "order_id": "00124",
            "customer_name": "Ada",
            "email": "",
            "order_date": "2026-08-21",
            "amount": "0",
            "status": "pending",
            "sales_channel": "store",
            "source_file": "orders.csv",
            "source_sheet": None,
            "source_row": 3,
        },
    ]


def test_load_csv_file_uses_physical_row_for_multiline_records(
    tmp_path: Path,
) -> None:
    file_path = write_csv(
        tmp_path / "orders.csv",
        f'{VALID_HEADER}\n00123,"Julia\nSilva",,2026-08-20,10.00,paid\n'
        "00124,Ada,,2026-08-21,20.00,pending\n",
    )

    records = load_csv_file(file_path)

    assert records[0]["source_row"] == 2
    assert records[1]["source_row"] == 4


def test_load_csv_file_preserves_commas_and_escaped_quotes(tmp_path: Path) -> None:
    file_path = write_csv(
        tmp_path / "orders.csv",
        f'{VALID_HEADER}\n00123,"Doe, Julia ""JJ""",,2026-08-20,10.00,paid\n',
    )

    records = load_csv_file(file_path)

    assert records[0]["customer_name"] == 'Doe, Julia "JJ"'


def test_load_csv_file_preserves_blank_rows(tmp_path: Path) -> None:
    file_path = write_csv(
        tmp_path / "orders.csv",
        f"{VALID_HEADER}\n\n{VALID_ROW}\n",
    )

    records = load_csv_file(file_path)

    assert records[0] == {
        "order_id": None,
        "customer_name": None,
        "email": None,
        "order_date": None,
        "amount": None,
        "status": None,
        "source_file": "orders.csv",
        "source_sheet": None,
        "source_row": 2,
    }
    assert records[1]["source_row"] == 3


def test_load_csv_file_preserves_short_rows_with_missing_values(
    tmp_path: Path,
) -> None:
    file_path = write_csv(
        tmp_path / "orders.csv",
        f"{VALID_HEADER}\n00123,Julia\n",
    )

    records = load_csv_file(file_path)

    assert records[0]["order_id"] == "00123"
    assert records[0]["customer_name"] == "Julia"
    assert records[0]["email"] is None
    assert records[0]["order_date"] is None
    assert records[0]["amount"] is None
    assert records[0]["status"] is None


def test_load_csv_file_rejects_structurally_invalid_quotes(tmp_path: Path) -> None:
    file_path = write_csv(
        tmp_path / "orders.csv",
        f'{VALID_HEADER}\n00123,"Julia"unexpected,,2026-08-20,10.00,paid\n',
    )

    with pytest.raises(CsvStructuralError, match="malformed CSV"):
        load_csv_file(file_path)


def test_load_csv_file_rejects_rows_wider_than_header(tmp_path: Path) -> None:
    file_path = write_csv(
        tmp_path / "orders.csv",
        f"{VALID_HEADER}\n{VALID_ROW},unexpected\n",
    )

    with pytest.raises(CsvStructuralError, match="row 2 contains more values"):
        load_csv_file(file_path)


def test_load_csv_files_aborts_when_any_source_is_invalid(tmp_path: Path) -> None:
    valid_file = write_csv(
        tmp_path / "valid.csv", f"{VALID_HEADER}\n{VALID_ROW}\n"
    )
    invalid_file = write_csv(
        tmp_path / "invalid.csv", "order_id,status\n00123,paid\n"
    )

    with pytest.raises(CsvStructuralError, match="invalid.csv"):
        load_csv_files([valid_file, invalid_file])


def loaded_record(order_id: str | None, **overrides: object) -> dict[str, object]:
    record: dict[str, object] = {
        "order_id": order_id,
        "customer_name": " Julia ",
        "email": " JULIA@EXAMPLE.COM ",
        "order_date": " 2026-08-20 ",
        "amount": " 10.005 ",
        "status": " PAID ",
        "source_file": "orders.csv",
        "source_sheet": None,
        "source_row": 2,
        "sales_channel": "online",
    }
    record.update(overrides)
    return record


def test_process_records_normalizes_and_classifies_valid_unique_record() -> None:
    result = process_records([loaded_record(" 00123 ")])

    assert len(result.records) == 1
    assert result.valid_records == result.records
    assert result.invalid_records == ()
    assert result.duplicate_records == ()
    assert result.records[0].record["order_id"] == "00123"
    assert result.records[0].record["customer_name"] == "Julia"
    assert result.records[0].record["email"] == "julia@example.com"
    assert str(result.records[0].record["amount"]) == "10.01"


def test_process_records_classifies_every_repeated_nonempty_id() -> None:
    result = process_records(
        [
            loaded_record(" 001 ", source_row=2),
            loaded_record("001", source_row=3),
            loaded_record("002", source_row=4),
        ]
    )

    assert [record.record["source_row"] for record in result.duplicate_records] == [
        2,
        3,
    ]
    assert [record.record["order_id"] for record in result.valid_records] == ["002"]


def test_process_records_classifies_all_three_occurrences_as_duplicates() -> None:
    result = process_records(
        [
            loaded_record("001", source_row=2),
            loaded_record("001", source_row=3),
            loaded_record("001", source_row=4),
        ]
    )

    assert result.duplicate_records == result.records
    assert result.valid_records == ()


def test_three_duplicates_retain_independent_validation_state() -> None:
    result = process_records(
        [
            loaded_record("001", email="invalid", source_row=2),
            loaded_record("001", status="unknown", source_row=3),
            loaded_record("001", amount="invalid", source_row=4),
        ]
    )

    assert result.duplicate_records == result.records
    assert result.invalid_records == result.records
    assert result.valid_records == ()


def test_process_records_does_not_treat_missing_ids_as_duplicates() -> None:
    result = process_records(
        [
            loaded_record("", source_row=2),
            loaded_record("   ", source_row=3),
            loaded_record("001", source_row=4),
            loaded_record("001", source_row=5),
        ]
    )

    assert [record.is_duplicate for record in result.records] == [False, False, True, True]
    assert [record.is_invalid for record in result.records] == [True, True, False, False]


def test_process_records_allows_invalid_and_duplicate_overlap() -> None:
    result = process_records(
        [
            loaded_record("001", email="invalid", source_row=2),
            loaded_record("001", source_row=3),
        ]
    )

    first_record = result.records[0]
    assert first_record in result.invalid_records
    assert first_record in result.duplicate_records
    assert result.records[1] not in result.invalid_records
    assert result.records[1] in result.duplicate_records
    assert result.valid_records == ()


def test_process_records_compares_ids_case_sensitively_and_preserves_zeros() -> None:
    result = process_records(
        [loaded_record("001"), loaded_record("1"), loaded_record("ABC"), loaded_record("abc")]
    )

    assert result.duplicate_records == ()
    assert [record.record["order_id"] for record in result.valid_records] == [
        "001",
        "1",
        "ABC",
        "abc",
    ]


def test_process_records_preserves_traceability_extra_columns_and_accounting() -> None:
    input_records = [
        loaded_record("001", source_file="a.csv", source_row=2),
        loaded_record("002", source_file="b.csv", source_row=7, status="unknown"),
    ]

    result = process_records(input_records)

    assert len(result.records) == len(input_records)
    assert result.records[0].record["source_file"] == "a.csv"
    assert result.records[1].record["source_file"] == "b.csv"
    assert result.records[1].record["source_row"] == 7
    assert all(record.record["sales_channel"] == "online" for record in result.records)
    assert len(result.valid_records) == 1
    assert len(result.invalid_records) == 1


def test_process_records_does_not_mutate_raw_records() -> None:
    raw_record = loaded_record(
        " 001 ",
        customer_name=" Julia ",
        email=" JULIA@EXAMPLE.COM ",
        amount=" 10.005 ",
        status=" PAID ",
    )
    original_record = dict(raw_record)

    process_records([raw_record])

    assert raw_record == original_record


def test_classification_projections_reference_processed_record_instances() -> None:
    result = process_records(
        [
            loaded_record("001", source_row=2),
            loaded_record("002", status="unknown", source_row=3),
            loaded_record("003", email="invalid", source_row=4),
            loaded_record("003", source_row=5),
        ]
    )

    assert result.valid_records[0] is result.records[0]
    assert result.invalid_records[0] is result.records[1]
    assert result.invalid_records[1] is result.records[2]
    assert result.duplicate_records[0] is result.records[2]
    assert result.duplicate_records[1] is result.records[3]


def test_process_records_detects_duplicates_across_loaded_csv_files(
    tmp_path: Path,
) -> None:
    first_file = write_csv(
        tmp_path / "a.csv",
        f"{VALID_HEADER}\n00123,Julia,,2026-08-20,10.00,paid\n",
    )
    second_file = write_csv(
        tmp_path / "b.csv",
        f"{VALID_HEADER}\n00123,Ada,,2026-08-21,20.00,pending\n",
    )

    result = process_records(load_csv_files([first_file, second_file]))

    assert len(result.records) == 2
    assert len(result.duplicate_records) == 2
    assert result.valid_records == ()
    assert [record.record["source_file"] for record in result.duplicate_records] == [
        "a.csv",
        "b.csv",
    ]

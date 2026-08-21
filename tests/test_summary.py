from __future__ import annotations

from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook

from src.processor import (
    discover_supported_files,
    load_supported_files,
    process_records,
)
from src.summary import ProcessingSummary, calculate_summary


def raw_record(order_id: str, **overrides: object) -> dict[str, object]:
    record: dict[str, object] = {
        "order_id": order_id,
        "customer_name": "Julia",
        "email": None,
        "order_date": "2026-08-20",
        "amount": "10.00",
        "status": "paid",
        "source_file": "orders.csv",
        "source_sheet": None,
        "source_row": 2,
        "sales_channel": "online",
    }
    record.update(overrides)
    return record


def test_empty_internal_result_has_zero_summary() -> None:
    summary = calculate_summary(process_records([]))

    assert summary == ProcessingSummary(0, 0, 0, 0, Decimal("0.00"))
    assert isinstance(summary.total_paid_amount, Decimal)


def test_counts_valid_invalid_and_overlapping_duplicate_records() -> None:
    result = process_records(
        [
            raw_record("unique", source_row=2),
            raw_record("invalid", email="invalid", source_row=3),
            raw_record("repeated", source_row=4),
            raw_record("repeated", status="unknown", source_row=5),
        ]
    )

    summary = calculate_summary(result)

    assert summary.total_records == 4
    assert summary.valid_records == 1
    assert summary.invalid_records == 2
    assert summary.duplicate_records == 2
    assert summary.total_paid_amount == Decimal("10.00")


def test_all_occurrences_of_multiple_duplicates_are_counted() -> None:
    result = process_records(
        [raw_record("001", source_row=row) for row in (2, 3, 4)]
    )

    summary = calculate_summary(result)

    assert summary == ProcessingSummary(3, 0, 0, 3, Decimal("0.00"))


def test_paid_total_includes_only_valid_unique_paid_records() -> None:
    result = process_records(
        [
            raw_record("paid", amount="12.345", source_row=2),
            raw_record("invalid", amount="20.00", email="invalid", source_row=3),
            raw_record("duplicate", amount="30.00", source_row=4),
            raw_record("duplicate", amount="30.00", source_row=5),
            raw_record("pending", amount="40.00", status="pending", source_row=6),
            raw_record("cancelled", amount="50.00", status="cancelled", source_row=7),
            raw_record("refunded", amount="60.00", status="refunded", source_row=8),
            raw_record("zero", amount="0", source_row=9),
        ]
    )

    summary = calculate_summary(result)

    assert summary.total_records == 8
    assert summary.valid_records == 5
    assert summary.invalid_records == 1
    assert summary.duplicate_records == 2
    assert summary.total_paid_amount == Decimal("12.35")


def test_no_eligible_record_returns_decimal_zero_with_two_places() -> None:
    result = process_records(
        [
            raw_record("pending", status="pending"),
            raw_record("invalid", email="invalid"),
        ]
    )

    total = calculate_summary(result).total_paid_amount

    assert total == Decimal("0.00")
    assert total.as_tuple().exponent == -2


def test_sums_multiple_large_decimal_values_without_precision_loss() -> None:
    large_amount = "1234567890123456789012345678901234567890.12"
    result = process_records(
        [
            raw_record("large-1", amount=large_amount, source_row=2),
            raw_record("large-2", amount=large_amount, source_row=3),
            raw_record("small", amount="0.01", source_row=4),
        ]
    )

    total = calculate_summary(result).total_paid_amount

    assert total == Decimal("2469135780246913578024691357802469135780.25")
    assert isinstance(total, Decimal)


def test_preserves_one_large_normalized_decimal_exactly() -> None:
    amount = "9876543210987654321098765432109876543210.98"
    result = process_records([raw_record("large", amount=amount)])

    total = calculate_summary(result).total_paid_amount

    assert total == Decimal(amount)
    assert total.as_tuple().exponent == -2


def test_summary_does_not_mutate_processing_result_or_record_data() -> None:
    result = process_records([raw_record("001")])
    records_before = result.records
    valid_records_before = result.valid_records
    invalid_records_before = result.invalid_records
    duplicate_records_before = result.duplicate_records
    record_before = dict(result.records[0].record)

    calculate_summary(result)

    assert result.records is records_before
    assert result.valid_records is valid_records_before
    assert result.invalid_records is invalid_records_before
    assert result.duplicate_records is duplicate_records_before
    assert result.records[0].record == record_before
    assert result.valid_records[0] is result.records[0]


def test_calculates_summary_for_combined_csv_and_xlsx_sources(
    tmp_path: Path,
) -> None:
    (tmp_path / "orders.csv").write_text(
        "order_id,customer_name,email,order_date,amount,status,channel\n"
        "csv-1,Julia,,2026-08-20,10.00,paid,online\n",
        encoding="utf-8",
    )
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Orders"
    worksheet.append(
        [
            "order_id",
            "customer_name",
            "email",
            "order_date",
            "amount",
            "status",
            "channel",
        ]
    )
    worksheet.append(
        ["xlsx-1", "Ada", None, date(2026, 8, 21), 20, "paid", "store"]
    )
    workbook.save(tmp_path / "orders.xlsx")
    workbook.close()

    result = process_records(
        load_supported_files(discover_supported_files(tmp_path))
    )
    summary = calculate_summary(result)

    assert summary == ProcessingSummary(2, 2, 0, 0, Decimal("30.00"))
    assert [record.record["source_file"] for record in result.records] == [
        "orders.csv",
        "orders.xlsx",
    ]
    assert [record.record["source_sheet"] for record in result.records] == [
        None,
        "Orders",
    ]
    assert [record.record["channel"] for record in result.records] == [
        "online",
        "store",
    ]

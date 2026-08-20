from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal

import pytest

from src.validator import normalize_and_validate_record


def valid_record(**overrides: object) -> dict[str, object]:
    record: dict[str, object] = {
        "order_id": "00123",
        "customer_name": "Julia Silva",
        "email": "julia@example.com",
        "order_date": "2026-08-20",
        "amount": "149.90",
        "status": "paid",
        "source_file": "orders.csv",
        "source_sheet": None,
        "source_row": 2,
        "sales_channel": "online",
    }
    record.update(overrides)
    return record


def error_codes(**overrides: object) -> set[tuple[str, str]]:
    result = normalize_and_validate_record(valid_record(**overrides))
    return {(error.field, error.code) for error in result.errors}


def test_normalizes_approved_text_fields_and_preserves_other_values() -> None:
    result = normalize_and_validate_record(
        valid_record(
            order_id=" 00Ab ",
            customer_name="  Ada LOVELACE  ",
            email="  ADA@Example.COM ",
            status=" PAID ",
        )
    )

    assert result.errors == ()
    assert result.record["order_id"] == "00Ab"
    assert result.record["customer_name"] == "Ada LOVELACE"
    assert result.record["email"] == "ada@example.com"
    assert result.record["status"] == "paid"
    assert result.record["source_file"] == "orders.csv"
    assert result.record["source_sheet"] is None
    assert result.record["source_row"] == 2
    assert result.record["sales_channel"] == "online"


@pytest.mark.parametrize(
    "field",
    ["order_id", "customer_name", "order_date", "amount", "status"],
)
@pytest.mark.parametrize("missing_value", [None, "", "   "])
def test_required_values_are_missing_after_normalization(
    field: str,
    missing_value: object,
) -> None:
    assert (field, "required") in error_codes(**{field: missing_value})


@pytest.mark.parametrize("status", ["paid", "pending", "cancelled", "refunded"])
def test_accepts_every_valid_normalized_status(status: str) -> None:
    result = normalize_and_validate_record(valid_record(status=f" {status.upper()} "))

    assert result.errors == ()
    assert result.record["status"] == status


def test_rejects_unknown_status() -> None:
    assert ("status", "invalid_status") in error_codes(status="processing")


@pytest.mark.parametrize("email", [None, "", "   "])
def test_accepts_optional_missing_email(email: object) -> None:
    result = normalize_and_validate_record(valid_record(email=email))

    assert result.errors == ()


@pytest.mark.parametrize(
    "email",
    [
        "julia",
        "julia@",
        "@example.com",
        "julia@@example.com",
        "julia example.com",
        "julia@example",
        "julia@example..com",
    ],
)
def test_rejects_malformed_present_email(email: str) -> None:
    assert ("email", "invalid_email") in error_codes(email=email)


def test_normalizes_valid_email_before_validation() -> None:
    result = normalize_and_validate_record(valid_record(email=" JULIA@EXAMPLE.COM "))

    assert result.errors == ()
    assert result.record["email"] == "julia@example.com"


def test_parses_trimmed_canonical_date_to_date() -> None:
    result = normalize_and_validate_record(valid_record(order_date=" 2026-08-20 "))

    assert result.errors == ()
    assert result.record["order_date"] == date(2026, 8, 20)
    assert type(result.record["order_date"]) is date


def test_accepts_date_and_normalizes_datetime_to_date() -> None:
    date_result = normalize_and_validate_record(
        valid_record(order_date=date(2026, 8, 20))
    )
    datetime_result = normalize_and_validate_record(
        valid_record(order_date=datetime(2026, 8, 20, 15, 30))
    )

    assert date_result.record["order_date"] == date(2026, 8, 20)
    assert datetime_result.record["order_date"] == date(2026, 8, 20)
    assert date_result.errors == datetime_result.errors == ()


def test_accepts_valid_leap_year_date() -> None:
    result = normalize_and_validate_record(valid_record(order_date="2024-02-29"))

    assert result.errors == ()
    assert result.record["order_date"] == date(2024, 2, 29)


def test_rejects_invalid_leap_year_date() -> None:
    assert ("order_date", "invalid_date") in error_codes(
        order_date="2025-02-29"
    )


@pytest.mark.parametrize(
    "order_date",
    ["01/02/2026", "2026/08/20", "2026-8-20", "2026-02-30", "2026-08-20T10:00:00"],
)
def test_rejects_unsupported_or_invalid_date_text(order_date: str) -> None:
    assert ("order_date", "invalid_date") in error_codes(order_date=order_date)


@pytest.mark.parametrize(
    ("amount", "expected"),
    [
        ("1.005", Decimal("1.01")),
        ("12.999", Decimal("13.00")),
        ("149.9", Decimal("149.90")),
        (" 10.00 ", Decimal("10.00")),
        ("0", Decimal("0.00")),
        (Decimal("2.345"), Decimal("2.35")),
    ],
)
def test_parses_and_normalizes_valid_amounts(
    amount: object,
    expected: Decimal,
) -> None:
    result = normalize_and_validate_record(valid_record(amount=amount))

    assert result.errors == ()
    assert result.record["amount"] == expected
    assert isinstance(result.record["amount"], Decimal)


def test_normalizes_very_large_decimal_without_float_conversion() -> None:
    amount = "123456789012345678901234567890.555"

    result = normalize_and_validate_record(valid_record(amount=amount))

    assert result.errors == ()
    assert result.record["amount"] == Decimal("123456789012345678901234567890.56")


@pytest.mark.parametrize(
    "amount",
    ["$10.00", "10,00", "1e2", "NaN", "Infinity", "+10.00", "10."],
)
def test_rejects_noncanonical_or_nonfinite_amounts(amount: str) -> None:
    assert ("amount", "invalid_amount") in error_codes(amount=amount)


@pytest.mark.parametrize("amount", [Decimal("NaN"), Decimal("Infinity")])
def test_rejects_nonfinite_decimal_amounts(amount: Decimal) -> None:
    assert ("amount", "invalid_amount") in error_codes(amount=amount)


@pytest.mark.parametrize(
    ("amount", "expected"),
    [
        ("-0.001", Decimal("-0.00")),
        ("-0.004", Decimal("-0.00")),
        ("-0.005", Decimal("-0.01")),
        ("-1.00", Decimal("-1.00")),
    ],
)
def test_rejects_negative_amount_before_rounding(
    amount: str,
    expected: Decimal,
) -> None:
    result = normalize_and_validate_record(valid_record(amount=amount))

    assert ("amount", "negative_amount") in {
        (error.field, error.code) for error in result.errors
    }
    assert result.record["amount"] == expected


@pytest.mark.parametrize(
    ("amount", "expected"),
    [("0", Decimal("0.00")), ("0.001", Decimal("0.00"))],
)
def test_accepts_nonnegative_amounts_that_round_to_zero(
    amount: str,
    expected: Decimal,
) -> None:
    result = normalize_and_validate_record(valid_record(amount=amount))

    assert result.errors == ()
    assert result.record["amount"] == expected


def test_accumulates_multiple_errors_on_one_record() -> None:
    result = normalize_and_validate_record(
        valid_record(
            order_id=" ",
            customer_name=None,
            email="invalid",
            order_date="01/02/2026",
            amount="not-money",
            status="unknown",
        )
    )

    assert [(error.field, error.code) for error in result.errors] == [
        ("order_id", "required"),
        ("customer_name", "required"),
        ("email", "invalid_email"),
        ("order_date", "invalid_date"),
        ("amount", "invalid_amount"),
        ("status", "invalid_status"),
    ]

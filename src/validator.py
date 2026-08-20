"""Record-level normalization and validation rules."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
from typing import Mapping


VALID_STATUSES = frozenset({"paid", "pending", "cancelled", "refunded"})
DECIMAL_QUANTUM = Decimal("0.01")
CANONICAL_AMOUNT_PATTERN = re.compile(r"^-?\d+(?:\.\d+)?$")
CANONICAL_DATE_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}$")


@dataclass(frozen=True)
class ValidationError:
    """One deterministic row-level validation failure."""

    field: str
    code: str
    message: str


@dataclass(frozen=True)
class RecordValidationResult:
    """A normalized record and every validation error found in it."""

    record: dict[str, object]
    errors: tuple[ValidationError, ...]


def _is_missing(value: object) -> bool:
    return value is None or value == ""


def _normalize_text(
    value: object,
    *,
    lowercase: bool = False,
) -> object:
    if not isinstance(value, str):
        return value

    normalized = value.strip()
    return normalized.lower() if lowercase else normalized


def _required_error(field: str) -> ValidationError:
    return ValidationError(field, "required", f"{field} is required")


def _normalize_order_date(value: object) -> tuple[object, ValidationError | None]:
    if isinstance(value, datetime):
        return value.date(), None
    if isinstance(value, date):
        return value, None
    if value is None:
        return None, _required_error("order_date")
    if not isinstance(value, str):
        return value, ValidationError(
            "order_date",
            "invalid_date",
            "order_date must use YYYY-MM-DD",
        )

    normalized = value.strip()
    if not normalized:
        return normalized, _required_error("order_date")
    if not CANONICAL_DATE_PATTERN.fullmatch(normalized):
        return normalized, ValidationError(
            "order_date",
            "invalid_date",
            "order_date must use YYYY-MM-DD",
        )

    try:
        return date.fromisoformat(normalized), None
    except ValueError:
        return normalized, ValidationError(
            "order_date",
            "invalid_date",
            "order_date must be a valid calendar date in YYYY-MM-DD format",
        )


def _quantize_amount(value: Decimal) -> Decimal:
    precision = max(28, len(value.as_tuple().digits) + 2)
    with localcontext() as context:
        context.prec = precision
        return value.quantize(DECIMAL_QUANTUM, rounding=ROUND_HALF_UP)


def _normalize_amount(value: object) -> tuple[object, ValidationError | None]:
    if value is None:
        return None, _required_error("amount")

    if isinstance(value, str):
        normalized: object = value.strip()
        if not normalized:
            return normalized, _required_error("amount")
        if not CANONICAL_AMOUNT_PATTERN.fullmatch(normalized):
            return normalized, ValidationError(
                "amount",
                "invalid_amount",
                "amount must use canonical decimal notation",
            )
        decimal_input = normalized
    elif isinstance(value, Decimal):
        normalized = value
        decimal_input = value
    else:
        return value, ValidationError(
            "amount",
            "invalid_amount",
            "amount must use canonical decimal notation",
        )

    try:
        parsed = Decimal(decimal_input)
        if not parsed.is_finite():
            raise InvalidOperation
        is_negative = parsed < 0
        quantized = _quantize_amount(parsed)
    except InvalidOperation:
        return normalized, ValidationError(
            "amount",
            "invalid_amount",
            "amount must be a finite decimal value",
        )

    if is_negative:
        return quantized, ValidationError(
            "amount",
            "negative_amount",
            "amount must not be negative",
        )

    return quantized, None


def _validate_email(value: object) -> ValidationError | None:
    if _is_missing(value):
        return None
    if not isinstance(value, str):
        return ValidationError("email", "invalid_email", "email format is invalid")
    if any(character.isspace() for character in value) or value.count("@") != 1:
        return ValidationError("email", "invalid_email", "email format is invalid")

    local_part, domain = value.split("@")
    domain_labels = domain.split(".")
    if not local_part or len(domain_labels) < 2 or any(not label for label in domain_labels):
        return ValidationError("email", "invalid_email", "email format is invalid")
    return None


def normalize_and_validate_record(
    raw_record: Mapping[str, object],
) -> RecordValidationResult:
    """Normalize approved fields and accumulate all row-level errors."""
    record = dict(raw_record)
    errors: list[ValidationError] = []

    record["order_id"] = _normalize_text(record.get("order_id"))
    if _is_missing(record["order_id"]):
        errors.append(_required_error("order_id"))
    elif not isinstance(record["order_id"], str):
        errors.append(
            ValidationError("order_id", "invalid_type", "order_id must be text")
        )

    record["customer_name"] = _normalize_text(record.get("customer_name"))
    if _is_missing(record["customer_name"]):
        errors.append(_required_error("customer_name"))
    elif not isinstance(record["customer_name"], str):
        errors.append(
            ValidationError(
                "customer_name",
                "invalid_type",
                "customer_name must be text",
            )
        )

    record["email"] = _normalize_text(record.get("email"), lowercase=True)
    email_error = _validate_email(record["email"])
    if email_error is not None:
        errors.append(email_error)

    record["order_date"], date_error = _normalize_order_date(
        record.get("order_date")
    )
    if date_error is not None:
        errors.append(date_error)

    record["amount"], amount_error = _normalize_amount(record.get("amount"))
    if amount_error is not None:
        errors.append(amount_error)

    record["status"] = _normalize_text(record.get("status"), lowercase=True)
    if _is_missing(record["status"]):
        errors.append(_required_error("status"))
    elif not isinstance(record["status"], str) or record["status"] not in VALID_STATUSES:
        errors.append(
            ValidationError(
                "status",
                "invalid_status",
                "status must be one of: paid, pending, cancelled, refunded",
            )
        )

    return RecordValidationResult(record=record, errors=tuple(errors))

"""Aggregate processing results into the approved sales summary."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP, localcontext

from src.processor import ProcessingResult


MONEY_QUANTUM = Decimal("0.01")


@dataclass(frozen=True)
class ProcessingSummary:
    """The five approved aggregate metrics for one processing result."""

    total_records: int
    valid_records: int
    invalid_records: int
    duplicate_records: int
    total_paid_amount: Decimal


def _sum_paid_amounts(amounts: tuple[Decimal, ...]) -> Decimal:
    """Sum normalized amounts without losing precision for large values."""
    if not amounts:
        return Decimal("0.00")

    largest_operand_precision = max(
        len(amount.as_tuple().digits) + max(amount.as_tuple().exponent, 0)
        for amount in amounts
    )
    carry_digits = len(str(len(amounts)))

    with localcontext() as context:
        context.prec = max(28, largest_operand_precision + carry_digits + 2)
        total = sum(amounts, start=Decimal("0.00"))
        return total.quantize(MONEY_QUANTUM, rounding=ROUND_HALF_UP)


def calculate_summary(result: ProcessingResult) -> ProcessingSummary:
    """Calculate counts and paid total without modifying processed records."""
    paid_amounts: list[Decimal] = []
    for processed_record in result.valid_records:
        if processed_record.record.get("status") != "paid":
            continue

        amount = processed_record.record.get("amount")
        if not isinstance(amount, Decimal):
            raise TypeError("valid paid record amount must be Decimal")
        paid_amounts.append(amount)

    return ProcessingSummary(
        total_records=len(result.records),
        valid_records=len(result.valid_records),
        invalid_records=len(result.invalid_records),
        duplicate_records=len(result.duplicate_records),
        total_paid_amount=_sum_paid_amounts(tuple(paid_amounts)),
    )

# Business Rules

`requirements.md` is the source of truth for functional requirements. This document defines the approved rules that govern detailed processing behavior. Their formal decision history is recorded in `DECISIONS.md`.

## Input Schema

Records use the fields `order_id`, `customer_name`, `email`, `order_date`, `amount`, and `status`.

`order_id`, `customer_name`, `order_date`, `amount`, and `status` are required. Values that are empty after normalization are missing. `email` is optional.

Extra input columns are allowed and must be preserved. Missing required columns are structural source errors, not row-level validation errors.

## Order Identity and Duplicates

`order_id` is a textual, case-sensitive identifier. Leading and trailing whitespace is removed, while capitalization and leading zeros are preserved.

If an `order_id` occurs more than once across the complete loaded dataset, every occurrence is a duplicate. The application must not arbitrarily select one occurrence as valid.

Duplicate and invalid classifications are not mutually exclusive. A record may appear in both the `Duplicates` and `Invalid Records` worksheets.

Duplicate count is an independent summary metric. Because invalid and duplicate classifications may overlap, the application must not imply that `total_records = valid_records + invalid_records + duplicate_records`.

## Field Normalization and Validation

### Customer Name

Remove leading and trailing whitespace. Preserve all other capitalization.

### Email

Email is optional. When present, remove leading and trailing whitespace, convert it to lowercase, and apply basic pragmatic format validation that rejects obviously malformed addresses without attempting complete RFC validation. A present malformed email makes the record invalid.

### Status

Remove leading and trailing whitespace and convert the value to lowercase before validation. Valid statuses are `paid`, `pending`, `cancelled`, and `refunded`; all other values are invalid.

### Order Date

The canonical textual format is `YYYY-MM-DD`. Native Excel date and datetime cell values are also accepted. Ambiguous or unsupported textual formats are invalid and must not be guessed.

### Amount

Version 1 uses USD and canonical decimal notation without currency symbols or locale-dependent formatting. Monetary calculations use decimal arithmetic. When rounding is required, values use two decimal places and `ROUND_HALF_UP`.

Zero is valid. Negative amounts are invalid. Refunds are represented by the `refunded` status rather than negative amounts. Invalid amounts must not crash processing and must make the affected record invalid.

## Source Contracts

### CSV

Version 1 CSV files use UTF-8, contain a header, and use a comma delimiter. File extensions are matched case-insensitively. The application must not silently guess another encoding or delimiter.

### XLSX

Process each worksheet that contains the required schema. Preserve the worksheet name in source metadata. A worksheet that does not satisfy the structural contract must not silently produce zero records.

## Traceability and Auditability

Attach `source_file`, `source_sheet`, and `source_row` to every loaded record as early as possible. `source_sheet` may be empty or null for CSV records. Keep this metadata associated with the record throughout processing.

A processing timestamp is run-level metadata and must not be duplicated on every record. Normalized values are used for processing and normal output; do not copy every original field solely to preserve formatting when source metadata can locate the original record.

No loaded record may disappear without an explicit classification or documented processing error.

## Failure Policy

A structural or operational failure in any input source makes the run unsuccessful. Log the failure clearly, exit with a non-zero status, and do not publish an apparently complete final report after skipping the source. Partial-processing mode is outside Version 1 scope.

Row-level validation failures are retained as invalid records and do not by themselves constitute structural source failures.

## Reporting

The report contains the worksheets required by `requirements.md`. Valid records are both valid and unique. The total paid amount includes only records that are valid, unique, and normalized to `status == "paid"`.

An existing `data/output/sales_report.xlsx` may be replaced. Generate the new workbook in a temporary file and atomically replace the final path only after successful generation.

Input-controlled text beginning with a spreadsheet formula-triggering character, including `=`, `+`, `-`, or `@`, must be exported safely as text. The mitigation must prevent formula execution while preserving the underlying data as much as reasonably possible.

## Pending Business Decisions

None. New ambiguities discovered during implementation must be documented and submitted for approval rather than resolved by assumption.

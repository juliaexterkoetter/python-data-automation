# Testing Strategy

Tests have not been implemented yet. Commands below are placeholders until the Python environment and test configuration are established.

## Core Invariants

- No loaded record may disappear without an explicit classification or documented processing error.
- Invalid and duplicate classifications may overlap and must be reconciled independently.
- A structurally or operationally failed source must prevent successful final-report publication.
- A failed export must not destroy an existing valid report.

## Unit Tests

### Normalization and Validation

- Missing required values, including values that become empty after trimming.
- `order_id` trimming, case preservation, textual representation, and preservation of leading zeros such as `00123`.
- Customer-name trimming without capitalization changes.
- Status trimming and lowercase conversion before validation.
- Every valid status and unknown statuses.
- Optional absent email, normalized valid email, and present malformed email.
- Canonical `YYYY-MM-DD` dates.
- Rejection of ambiguous text such as `01/02/2026` and other unsupported date formats.
- Native Excel date and datetime values.
- Accepted canonical USD decimal inputs and rejected symbols or locale-dependent formats.
- Decimal arithmetic without binary floating-point artifacts.
- Two-decimal `ROUND_HALF_UP` behavior where rounding is required.
- Valid zero amounts and invalid negative amounts.
- Multiple validation errors on one record.

### Processing and Summary

- Consolidation across CSV files, workbooks, and multiple qualifying worksheets.
- Detection of duplicates within one source and across sources.
- Classification of every occurrence of a duplicated `order_id`.
- Simultaneous invalid and duplicate classification and presence in both result datasets.
- Independent duplicate count without additive category assumptions.
- Valid-record eligibility requiring both valid and unique status.
- Paid-total eligibility requiring valid, unique, normalized `paid` records.
- Exclusion of invalid, duplicate, pending, cancelled, and refunded records from paid total.
- Preservation of extra columns and source metadata.
- Empty datasets and record accounting against the core invariants.

### Export Safety

- Neutralization of input-controlled text beginning with `=`, `+`, `-`, or `@`.
- Preservation of the underlying textual value as far as the selected safe representation permits.
- Independent summary semantics communicated in the workbook.

## Integration Tests

Use temporary directories and small deterministic files to cover:

- UTF-8 comma-delimited CSV input with a header.
- Case-insensitive `.csv` and `.xlsx` extension discovery.
- Rejection of unsupported CSV encoding, delimiter, missing header, and malformed input without guessing.
- XLSX input containing native date cells.
- Processing every worksheet with the required schema and preserving worksheet names.
- A nonconforming worksheet causing a visible structural failure rather than silently yielding zero records.
- Missing required columns as structural errors.
- Extra columns accepted, preserved through processing, and written to output.
- Combined CSV and multi-worksheet XLSX input.
- Corrupted, protected, empty, zero-byte, or unreadable sources causing non-zero execution status.
- Failure of any source preventing publication of an apparently complete report.
- `source_file`, `source_sheet`, and `source_row` attached correctly and retained in output datasets.
- Run-level timestamp not duplicated onto every record.
- Workbook creation and read-back verification.
- Required worksheet names, columns, classifications, and summary content.
- Existing report replacement after successful generation.
- Simulated generation or replacement failure leaving the existing valid report unchanged.
- Formula-injection payloads remaining non-executable after workbook read-back.
- A complete successful end-to-end execution.

## Regression Tests

Every corrected defect should receive a minimal test that reproduces the original failure. Production-like examples must be anonymized and must not contain credentials or sensitive business data.

## Fixtures

Fixtures should be small, readable, and focused on one behavior where possible. Prefer programmatically created temporary CSV and XLSX files for integration tests. Shared fixtures should include valid records, each invalid condition, repeated IDs, overlapping invalid/duplicate records, multiple worksheets, extra columns, traceability metadata, and formula-like text.

## Edge Cases

Test handling of:

- No input directory or no supported files.
- UTF-8 byte order marks and malformed UTF-8.
- Duplicate or blank headers.
- Protected, empty, or structurally invalid workbooks.
- IDs that differ only by case and IDs with leading or trailing whitespace.
- Excel dates, datetimes, and invalid text dates.
- Zero, negative, non-finite, over-precision, and very large monetary values.
- Unicode whitespace and invisible characters.
- Collisions between extra-column names and internal metadata names.
- Existing, locked, or unavailable output paths.
- Spreadsheet formula payloads in required and extra textual columns.
- Data volumes approaching workbook limits when an operational volume is defined.

## Placeholder Commands

```bash
python -m pytest
python -m pytest tests/test_validator.py
python -m pytest tests/test_processor.py
python -m pytest --cov=src --cov-report=term-missing
```

These commands must be confirmed after dependencies and test configuration are added.

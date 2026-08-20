# Testing Strategy

Tests are implemented incrementally with each feature. CSV and XLSX discovery and structural loading, record normalization, validation, duplicate detection, and classification are currently covered; tests for later pipeline stages remain planned.

## Implemented CSV Coverage

- Missing input directory and absence of supported CSV files.
- Case-insensitive, non-recursive discovery of regular `.csv` files.
- UTF-8 and UTF-8 BOM input, plus rejection of invalid UTF-8 bytes.
- Required comma delimiter, header presence, exact case-sensitive header names, duplicate headers, and required columns.
- Reserved traceability-column collisions and preservation of extra columns.
- Textual `order_id` values with leading zeros preserved.
- `source_file`, null `source_sheet`, and physical 1-based `source_row`, including multiline records.
- Complete-operation failure when any CSV source is structurally invalid.
- Logging and zero/non-zero exit-status coordination.

## Implemented Record-Processing Coverage

- Field-specific trimming, case handling, required values, and preservation of extra columns and traceability metadata.
- Pragmatic optional-email validation and all accepted status values.
- Canonical textual dates, native date and datetime values, internal `datetime.date` normalization, and rejection of unsupported formats.
- Canonical decimal parsing, `Decimal` representation, `ROUND_HALF_UP`, excess precision, zero, negatives, non-finite values, and very large amounts.
- Multiple validation errors accumulated on one record without aborting the run.
- Global, case-sensitive duplicate detection within and across CSV files, excluding missing normalized IDs.
- Classification of every repeated-ID occurrence, invalid-and-duplicate overlap, and valid-record eligibility.
- Preservation of record order, source metadata, extra columns, and complete record accounting.

## Implemented XLSX Coverage

- Case-insensitive, non-recursive discovery of regular `.xlsx` files without following symlinks.
- Deterministic mixed CSV and XLSX discovery and loading.
- Physical row-1 exact headers, duplicate headers, populated empty-header columns, required columns, reserved columns, and extra columns.
- Auxiliary, partial-schema, complete-schema, header-only, empty, and multiple worksheets according to the approved classification policy.
- Explicit auxiliary-sheet logging and workbook failure when no usable worksheet exists.
- Physical `source_row`, exact `source_sheet`, source filename, and preservation of empty intermediate rows.
- Native dates, datetimes, integers, floats, booleans, numeric identifiers, and shared record validation.
- Formula rejection in headers and record-level rejection in required and extra data cells without cached values.
- Corrupted workbook failures with preserved causes and complete-operation failure across mixed sources.
- Duplicate detection across CSV and XLSX records.

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
- Basic email boundaries: `julia@example.com` is valid, while `julia`, `julia@`, `@example.com`, `julia@@example.com`, `julia example.com`, and `julia@example` are invalid.
- Canonical `YYYY-MM-DD` dates.
- Rejection of ambiguous text such as `01/02/2026` and other unsupported date formats.
- Native Excel date and datetime values.
- Accepted canonical USD decimal inputs and rejected symbols or locale-dependent formats.
- Decimal arithmetic without binary floating-point artifacts.
- Excess input precision accepted and normalized with `Decimal` and `ROUND_HALF_UP`: `1.005` to `1.01`, `12.999` to `13.00`, and `149.9` to `149.90`.
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
- A worksheet with no required columns skipped as auxiliary with an explicit log entry.
- A worksheet with some but not all required columns causing structural failure.
- A workbook with auxiliary worksheets but no usable data worksheet causing structural failure.
- Missing required columns as structural errors.
- Reserved traceability-column collisions causing structural errors without overwriting, renaming, or discarding input data.
- Extra columns accepted, preserved through processing, and written to output.
- Combined CSV and multi-worksheet XLSX input.
- Corrupted, protected, empty, zero-byte, or unreadable sources causing non-zero execution status.
- Failure of any source preventing publication of an apparently complete report.
- `source_file`, `source_sheet`, and `source_row` attached correctly and retained in output datasets.
- Physical 1-based `source_row` values including the header, with the first record below a row-1 header mapped to row 2.
- Run-level timestamp not duplicated onto every record.
- Workbook creation and read-back verification.
- Required worksheet names, columns, classifications, and summary content.
- Existing report replacement after successful generation.
- Simulated generation or replacement failure leaving the existing valid report unchanged.
- Formula-injection payloads remaining non-executable after workbook read-back.
- A complete successful end-to-end execution.
- Missing input directory, no supported files, and no usable data source each causing non-zero status without creating a new final report.

## Regression Tests

Every corrected defect should receive a minimal test that reproduces the original failure. Production-like examples must be anonymized and must not contain credentials or sensitive business data.

## Fixtures

Fixtures should be small, readable, and focused on one behavior where possible. Prefer programmatically created temporary CSV and XLSX files for integration tests. Shared fixtures should include valid records, each invalid condition, repeated IDs, overlapping invalid/duplicate records, multiple worksheets, extra columns, traceability metadata, and formula-like text.

## Edge Cases

Test handling of:

- Missing input directory, no supported files, and no usable source.
- UTF-8 byte order marks and malformed UTF-8.
- Duplicate or blank headers.
- Protected, empty, or structurally invalid workbooks.
- IDs that differ only by case and IDs with leading or trailing whitespace.
- Excel dates, datetimes, and invalid text dates.
- Zero, negative, non-finite, over-precision, and very large monetary values, with over-precision normalized rather than rejected.
- Unicode whitespace and invisible characters.
- Collisions between extra-column names and reserved traceability names, which must fail structurally.
- Existing, locked, or unavailable output paths.
- Spreadsheet formula payloads in required and extra textual columns.
- Data volumes approaching workbook limits when an operational volume is defined.

## Test Commands

```bash
.venv/bin/python -m pytest
.venv/bin/python -m pytest tests/test_validator.py tests/test_processor.py tests/test_xlsx_processor.py tests/test_main.py
```

The complete-suite command is confirmed for the current development environment. Additional commands will be documented when later test modules and coverage tooling are added.

# Testing Strategy

Tests cover the implemented Version 1 pipeline: CSV and XLSX discovery and structural loading, record normalization, validation, duplicate detection, classification, summary calculation, workbook generation, and atomic report publication.

## Implemented CSV Coverage

- Missing input directory and absence of supported CSV files.
- Case-insensitive, non-recursive discovery of regular `.csv` files.
- UTF-8 and UTF-8 BOM input, plus rejection of invalid UTF-8 bytes.
- Required comma delimiter, header presence, non-empty exact case-sensitive header names, duplicate headers, and required columns.
- Reserved traceability-column collisions and preservation of extra columns.
- Textual `order_id` values with leading zeros preserved.
- `source_file`, null `source_sheet`, and physical 1-based `source_row`, including multiline records.
- Complete-operation failure when any CSV source is structurally invalid.
- Explicit 128 KiB field-size boundary, controlled failure above it, and restoration of the surrounding parser setting.
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
- Formula rejection in headers and record-level rejection in required and extra data cells without cached values, including deterministic traces for ordinary, array, and data-table formulas across repeated workbook loads.
- Corrupted workbook failures with preserved causes and complete-operation failure across mixed sources.
- Duplicate detection across CSV and XLSX records.
- Active defusedxml integration and controlled rejection of a minimal prohibited XML entity declaration without dangerous expansion.
- ZIP preflight without extraction, including empty and non-XLSX packages, content-types and relationship roots, conventional and custom workbook and worksheet paths, safe internal relationship resolution, invalid or external targets, unsupported compression, CRC integrity for used and ignored members, member names, duplicates, encryption, member counts, compressed and uncompressed sizes, zero compressed sizes, and exact material individual and aggregate compression-ratio boundaries.
- Exact and exceeded boundaries for worksheet count, explicit and implicit physical row and column references, total logical cells, and sparse extreme coordinates, with worksheet counting scoped to SpreadsheetML `sheetData` elements.
- Revalidation against the loaded workbook model and external-link caches disabled with `keep_links=False`.
- Complete-operation failure before partial records can escape when package preflight or workbook limits fail.

## Implemented Summary Coverage

- Total, valid, invalid, and duplicate counts derived from the existing classification projections.
- Independent invalid and duplicate counts, including records present in both classifications.
- All occurrences of repeated identifiers included in the duplicate count.
- Paid-total eligibility limited to valid, unique records normalized to `paid`.
- Exclusion of invalid, duplicate, pending, cancelled, and refunded records from the paid total.
- Decimal zero when no record is eligible and exact two-place aggregation of normalized values.
- Multiple very large monetary values summed without silent context-precision loss or float conversion.
- Combined CSV and XLSX inputs with source metadata and extra columns left unchanged.
- Summary calculation without mutation of the processing result or its records.
- Structural input failure stopping execution before summary calculation.

## Implemented Report Coverage

- Exact four-worksheet structure and approved summary labels, values, order, and overlap note.
- Export of the supplied processing projections and summary without recalculation.
- Deterministic global extra-column union, collision-safe validation-error columns, source metadata, and record order.
- Exact two-place monetary text, textual identifiers, native dates, and strict supported-type handling.
- Formula-injection protection for every approved trigger, initial ASCII whitespace, retained input formulas, extra values and names, and source metadata.
- Independent logical round-trip verification with `data_only=False`, including complete values and row order, relevant formats, non-formula cell types, and preserved carriage returns.
- Pre-save renderer fault injection proving that omitted records or headers, changed headers, reversed rows, and failed formula neutralization cannot alter the independent logical expectation or be published.
- XLSX row, column, UTF-16 cell-length, XML-character, finite-float, and unsupported-type validation without silent data loss.
- Output-directory and temporary-file creation, first publication, replacement, complete logical workbook validation, temporary cleanup, and preservation of an existing report across injected construction, conversion, save, close, permission, logical corruption, and replacement failures.
- Explicit rejection of negative signed float zero when its sign cannot survive XLSX numeric serialization.
- Deterministic logical workbook comparison across repeated exports.
- Complete CSV and XLSX processing through summary calculation and reopened final report.
- Main-pipeline success only after publication and controlled non-zero status for structural or export failures.

## Core Invariants

- No loaded record may disappear without an explicit classification or documented processing error.
- Invalid and duplicate classifications may overlap and must be reconciled independently.
- A structurally or operationally failed source must prevent successful final-report publication.
- A failed export must not destroy an existing valid report.

## Portfolio Renderer Coverage

- Exact deterministic validation of all four programmatic portfolio assets against a temporary regeneration from the source workbook.
- Rejection of white or otherwise incorrect valid PNGs, single-pixel changes, swapped filenames, wrong dimensions, corruption, and missing assets.
- Byte-identical repeated rendering and preservation of the source workbook hash.
- Portfolio tests use the dedicated Pillow dependency included through `requirements-portfolio.txt`; this does not add Pillow to the Version 1 runtime requirements.

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
- XLSX packages at and beyond approved file, member, expansion, worksheet, row, column, and logical-cell limits, using small synthetic fixtures and lowered test limits rather than dangerous payloads.

## Test Commands

```bash
.venv/bin/python -m pip install -r requirements-dev.txt
.venv/bin/python -m pytest
.venv/bin/python -m pytest tests/test_validator.py tests/test_processor.py tests/test_xlsx_processor.py tests/test_summary.py tests/test_exporter.py tests/test_main.py
.venv/bin/python -m compileall src tests
.venv/bin/python -m pip check
git diff --check
```

The complete-suite command is confirmed for the current development environment. Additional commands will be documented when later test modules and coverage tooling are added.

Release dependency auditing uses PyPA `pip-audit` from a separate tool environment:

```bash
python -m pip_audit -r requirements.txt
python -m pip_audit -r requirements-dev.txt
```

The validated tool version, latest result, and audit limitations are recorded in `RELEASE_READINESS.md`.

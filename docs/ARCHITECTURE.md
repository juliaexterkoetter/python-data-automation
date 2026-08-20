# Architecture

The initial architecture is approved but not implemented. It uses the existing modules and avoids unnecessary layers or global mutable state.

## Processing Pipeline

```text
discover CSV and XLSX sources
    -> validate each source structure
    -> load records and attach source metadata
    -> combine records across files and worksheets
    -> normalize approved fields
    -> validate records
    -> classify invalid records and all duplicate occurrences independently
    -> calculate summary information
    -> generate a protected workbook in a temporary file
    -> atomically replace the final report
```

A structural or operational failure stops successful publication and results in a non-zero exit status. Version 1 does not publish partial results.

## Module Responsibilities

### `src/main.py`

- Provide the application entry point.
- Configure logging and resolve input and output paths.
- Coordinate discovery, processing, validation, and export.
- Stop publication after structural or operational failure.
- Convert fatal failures into clear English messages and non-zero exit status codes.
- Avoid containing detailed business rules.

### `src/processor.py`

- Discover supported files case-insensitively.
- Enforce the UTF-8 comma-delimited CSV contract.
- Inspect and load every qualifying XLSX worksheet.
- Treat required-column failures as structural source errors.
- Preserve extra columns.
- Attach `source_file`, `source_sheet`, and `source_row` while loading.
- Preserve `order_id` as text, including leading zeros.
- Combine records across all successfully loaded sources.
- Apply approved normalization transformations.
- Identify every occurrence of repeated IDs across the combined dataset.
- Calculate independent counts and the eligible paid total using decimal arithmetic.

### `src/validator.py`

- Apply required-value, email, amount, date, and status rules.
- Accept canonical `YYYY-MM-DD` text and native Excel date or datetime values without guessing ambiguous text formats.
- Use decimal arithmetic, two decimal places, and `ROUND_HALF_UP` where rounding is required.
- Return structured validation results rather than hiding failures.
- Preserve all relevant validation errors for a record instead of stopping at the first error.
- Keep row validation independent from structural source checks and report formatting.

### `src/exporter.py`

- Generate the required workbook and worksheets.
- Preserve required source metadata and extra input columns.
- Represent duplicate count as an independent metric and make overlapping classifications clear.
- Export input-controlled textual values without allowing spreadsheet formula execution.
- Preserve a stable and documented column order and apply practical formatting.
- Write to a temporary file and atomically replace the final report only after successful generation.
- Surface generation, permission, and replacement failures as operational errors.

## Data Model and Flow

Loaded records move through the pipeline with a stable internal identity and their source metadata. Each stage returns explicit results rather than mutating shared global data. Normalized values, validation errors, duplicate membership, and summary eligibility remain attributable to the same loaded record.

Invalid and duplicate are independent attributes. A record may belong to both report datasets. The valid-record dataset contains only records that have no validation errors and whose `order_id` is not duplicated.

Extra input columns remain associated with their records through normal output. Normalized values are the normal processing and report values; full copies of original fields are not required because source metadata identifies the original record.

A small `dataclass` may represent the processing result and group valid records, invalid records, duplicates, summary values, source errors, and run metadata. The processing timestamp belongs to run metadata rather than individual rows.

## Traceability

`source_file`, `source_sheet`, and `source_row` are attached during loading and retained throughout processing. `source_sheet` is empty or null for CSV input. Worksheet names are preserved for XLSX records.

Record accounting must enforce the invariant that no loaded record disappears without classification or a documented error.

## Error Boundaries

Row-level validation errors, such as a malformed email, unsupported textual date, negative amount, or missing required value, are accumulated on the record and reported in `Invalid Records`.

Structural and operational failures include missing required columns, a worksheet that cannot satisfy the source contract, unreadable or malformed files, and failed report publication. They are logged clearly, prevent a successful run, and produce a non-zero exit status. They must not be converted into ordinary invalid rows or hidden behind an apparently complete report.

## Expected Dependencies

- `pandas` for tabular loading, transformation, consolidation, and classification, with explicit safeguards for textual identifiers and decimal values.
- `openpyxl` for native Excel values, multi-worksheet XLSX processing, and workbook output.
- `pytest` for automated tests.
- `pytest-cov` for optional coverage reporting.

Use the standard library for paths, logging, `Decimal`, dataclasses, temporary files, and atomic replacement where appropriate. Dependencies must not be installed or added until implementation planning confirms them.

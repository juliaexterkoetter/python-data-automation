# Architecture

The initial architecture is approved and partially implemented. CSV and XLSX discovery and structural loading, record normalization, row-level validation, duplicate detection, record classification, and summary calculation are implemented; report generation remains planned. The architecture uses focused modules and avoids unnecessary layers or global mutable state.

## Processing Pipeline

```text
discover CSV and XLSX sources
    -> fail if no supported source exists
    -> classify XLSX worksheets and validate data-source structures
    -> load records and attach source metadata
    -> combine records across files and worksheets
    -> normalize approved fields
    -> validate records
    -> classify invalid records and all duplicate occurrences independently
    -> calculate summary information
    -> generate a protected workbook in a temporary file
    -> atomically replace the final report
```

A structural or operational failure stops successful publication and results in a non-zero exit status. Version 1 does not publish partial or empty successful reports.

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
- Discover only regular files directly inside the input directory, without recursion.
- Fail when the input directory is missing, no supported file exists, or no usable data source is found.
- Enforce the comma-delimited CSV contract using strict `utf-8-sig` decoding, accepting UTF-8 with or without BOM and rejecting other encodings.
- Validate CSV header names exactly and case-sensitively without silent normalization.
- Classify XLSX worksheets by required-column presence, log skipped auxiliary worksheets, fail on partial schemas, and load complete schemas.
- Read XLSX headers only from physical row 1 and validate them exactly without silent normalization.
- Open XLSX workbooks with `read_only=False` and `data_only=False`, close them after success or failure, and surface predictable workbook errors structurally.
- Preserve empty XLSX rows between records and reject populated columns with empty headers.
- Retain formulas without evaluating them so row validation can reject formulas in required or extra data cells.
- Fail a workbook that contains no usable data worksheet.
- Treat required-column failures as structural source errors.
- Reject input columns that collide with reserved traceability names.
- Preserve extra columns.
- Attach `source_file`, `source_sheet`, and physical 1-based `source_row` while loading.
- Preserve `order_id` as text, including leading zeros.
- Combine records across all successfully loaded sources.
- Coordinate record normalization and validation without duplicating field rules.
- Identify every occurrence of repeated IDs across the combined dataset.
- Exclude missing normalized IDs from duplicate detection.
- Produce explicit valid, invalid, and duplicate classifications while preserving overlap and record order.

### `src/validator.py`

- Own the approved field-normalization and row-validation rules.
- Trim approved textual fields and preserve the required capitalization and identifier semantics.
- Apply required-value, explicitly bounded basic email, amount, date, and status rules.
- Accept canonical `YYYY-MM-DD` text and native Excel date or datetime values without guessing ambiguous text formats.
- Normalize valid dates internally to `datetime.date`.
- Parse finite canonical decimal amounts and normalize them to two places using `Decimal` and `ROUND_HALF_UP`.
- Convert native XLSX integers exactly and floats through `Decimal(str(value))`, while rejecting booleans.
- Return structured validation results rather than hiding failures.
- Preserve all relevant validation errors for a record instead of stopping at the first error.
- Keep row validation independent from structural source checks and report formatting.

### `src/summary.py`

- Aggregate an existing `ProcessingResult` without repeating normalization, validation, or duplicate detection.
- Report total, valid, invalid, and duplicate record counts with intentionally overlapping invalid and duplicate classifications.
- Calculate total paid amount from valid, unique records normalized to `status == "paid"`.
- Use `Decimal` exclusively for monetary arithmetic and preserve exact two-place results for very large aggregates.
- Return an immutable summary without modifying processed records or their classification projections.

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

Small dataclasses represent validation errors, processed records, the processing result, and its independently calculated summary. Valid, invalid, and duplicate projections reference the same processed records so overlapping classifications remain consistent. The summary contains only the approved counts and paid total, while source errors and run metadata remain separate concerns. The processing timestamp belongs to run metadata rather than individual rows.

## Traceability

`source_file`, `source_sheet`, and `source_row` are attached during loading and retained throughout processing. `source_sheet` is empty or null for CSV input. Worksheet names are preserved for XLSX records. `source_row` is the physical 1-based source row, including the header row, so a first data record below a row-1 header has value 2.

The traceability names are reserved. A collision with an input column is a structural source error and must not be resolved by silently overwriting, renaming, or discarding client data.

Record accounting must enforce the invariant that no loaded record disappears without classification or a documented error.

## Error Boundaries

Row-level validation errors, such as a malformed email, unsupported textual date, negative amount, or missing required value, are accumulated on the record and reported in `Invalid Records`.

An XLSX worksheet with none of the required columns is auxiliary: it is skipped with an explicit non-data log entry. A worksheet with some but not all required columns is a malformed data source. A complete schema is processed. A workbook without any complete-schema worksheet is structurally invalid.

Structural and operational failures include missing input directories, no supported or usable sources, partial required schemas, reserved-column collisions, unreadable or malformed files, and failed report publication. They are logged clearly, prevent a successful run, and produce a non-zero exit status. They must not be converted into ordinary invalid rows or hidden behind an apparently complete or empty report.

## Dependencies

- `openpyxl` for native Excel values and multi-worksheet XLSX processing, and later workbook output.
- `pytest` as a development dependency for automated tests.

The implemented loading and processing pipeline does not require `pandas`. Use the standard library for paths, logging, `Decimal`, dataclasses, temporary files, and atomic replacement where appropriate.

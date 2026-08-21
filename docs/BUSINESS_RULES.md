# Business Rules

`requirements.md` is the source of truth for functional requirements. This document defines the approved rules that govern detailed processing behavior. Their formal decision history is recorded in `DECISIONS.md`.

## Input Schema

Records use the fields `order_id`, `customer_name`, `email`, `order_date`, `amount`, and `status`.

`order_id`, `customer_name`, `order_date`, `amount`, and `status` are required. Values that are empty after normalization are missing. `email` is optional.

Extra input columns are allowed and must be preserved. Missing required columns are structural source errors, not row-level validation errors.

Every CSV header name must be non-empty. An empty internal or trailing header is a structural source error and must not be renamed, removed, or assigned a placeholder.

## Order Identity and Duplicates

`order_id` is a textual, case-sensitive identifier. Leading and trailing whitespace is removed, while capitalization and leading zeros are preserved.

If an `order_id` occurs more than once across the complete loaded dataset, every occurrence is a duplicate. The application must not arbitrarily select one occurrence as valid.

Duplicate detection considers only normalized, non-empty `order_id` values. Records with a missing normalized ID are invalid but are not duplicates solely because other records also have a missing ID.

Duplicate and invalid classifications are not mutually exclusive. A record may appear in both the `Duplicates` and `Invalid Records` worksheets.

Duplicate count is an independent summary metric. Because invalid and duplicate classifications may overlap, the application must not imply that `total_records = valid_records + invalid_records + duplicate_records`.

## Field Normalization and Validation

### Customer Name

Remove leading and trailing whitespace. Preserve all other capitalization.

### Email

Email is optional. When present, remove leading and trailing whitespace and convert it to lowercase before validation. Basic validation requires exactly one `@` separating a non-empty local part from a non-empty domain, no whitespace anywhere in the normalized address, and a plausible dotted domain. The domain must contain at least one dot with non-empty content on both sides.

Examples such as `julia`, `julia@`, `@example.com`, `julia example.com`, and `julia@example` are invalid. `julia@example.com` is valid. This is intentionally pragmatic validation, not RFC-complete validation. A present malformed email makes the record invalid.

### Status

Remove leading and trailing whitespace and convert the value to lowercase before validation. Valid statuses are `paid`, `pending`, `cancelled`, and `refunded`; all other values are invalid.

### Order Date

Remove leading and trailing whitespace from textual values before parsing. Whitespace-only values are missing. The canonical textual format is `YYYY-MM-DD`. Native Excel date and datetime cell values are also accepted. Ambiguous or unsupported textual formats are invalid and must not be guessed.

Represent every valid normalized order date internally as `datetime.date`. Normalize accepted datetime values to their date component. External formatting belongs to the export stage.

### Amount

Version 1 uses USD and canonical decimal notation without currency symbols or locale-dependent formatting. Remove leading and trailing whitespace from textual values before parsing; whitespace-only values are missing. Finite decimal inputs may contain more than two decimal places. Parse them directly as `Decimal` and normalize every valid amount to two decimal places using `ROUND_HALF_UP`. Excess decimal precision alone does not make a record invalid. For example, `1.005` becomes `1.01`, `12.999` becomes `13.00`, and `149.9` becomes `149.90`.

Zero is valid. Negative amounts are invalid. Refunds are represented by the `refunded` status rather than negative amounts. Invalid amounts must not crash processing and must make the affected record invalid.

## Source Contracts

### CSV

Version 1 CSV files use UTF-8, contain a header, and use a comma delimiter. UTF-8 with BOM is accepted and read with `utf-8-sig` so the BOM does not become part of the first header name. Other encodings remain invalid. The application must not silently guess another encoding or delimiter.

Each CSV field is limited to 128 KiB (131,072 characters). A field above that explicit limit is a structural source error; the application does not truncate it or increase the limit silently.

File extensions are matched case-insensitively. Discovery is not recursive: only regular files directly inside `data/input/` are considered, and files in subdirectories are ignored.

Header names are validated exactly and case-sensitively. The application must not trim, lowercase, rename, correct, or otherwise normalize them silently.

### XLSX

Classify each worksheet from its header columns:

- A worksheet containing none of the required columns is auxiliary. Skip it and log that it was ignored as non-data.
- A worksheet containing at least one, but not all, required columns is a malformed data source and causes a structural failure.
- A worksheet containing the complete required schema is a data worksheet and must be processed.

Preserve each processed worksheet name in source metadata. A workbook containing no usable data worksheet causes a structural failure.

The header must be on physical row 1 and is compared exactly and case-sensitively without trimming or correction. Duplicate non-empty names are structural errors. An empty header over a populated column is also a structural error; only trailing empty header cells beyond the effectively used data may be ignored.

A complete-schema worksheet is usable only when it contains at least one data record. Preserve empty physical rows between records so their row numbers and subsequent row-level validation remain traceable.

XLSX discovery is non-recursive, case-insensitive by extension, limited to regular files directly inside the input directory, and does not follow symlinks.

`data/input/` is a trusted local operator directory. The Version 1 workflow is not designed to resist a concurrent local attacker who can replace directory entries while a run is in progress. Operators must prevent untrusted writes during processing.

Native numeric `order_id` values are invalid and are not converted to text. Native integer amounts are converted exactly to `Decimal`; native float amounts use `Decimal(str(value))`, never `Decimal(float_value)`. Boolean values are invalid in every business field.

Before openpyxl loads an XLSX input, inspect its ZIP package without extracting members. Reject empty or non-XLSX ZIPs, invalid or missing content-types, workbook, or workbook-relationships parts, unsafe, external, duplicate, or missing referenced worksheet targets, unsupported compression methods, CRC failures in any member, suspicious or duplicate member names, encrypted members, prohibited XML, and packages that exceed any approved resource limit. Resolve referenced worksheet parts through the workbook's OOXML relationships without assuming conventional package paths. Stored and deflated ZIP members are the only supported compression methods.

- 10 MiB XLSX file size;
- 1,000 ZIP members;
- 100 MiB total uncompressed size;
- 50 MiB for an individual member;
- 100:1 individual or aggregate material compression ratio;
- 50 worksheets;
- 100,000 rows per worksheet;
- 256 columns per worksheet;
- 1,000,000 total logical cells across worksheets.

Compression-ratio enforcement is material when an individual uncompressed member exceeds 1 MiB or the package total uncompressed size exceeds 10 MiB. Logical-cell accounting uses the effective worksheet area and must detect sparse extreme coordinates rather than trusting only declared dimensions.

Open XLSX input with `read_only=False`, `data_only=False`, and `keep_links=False`. Do not preserve external-workbook caches or access external resources. Formulas are never executed or interpreted. A formula in a header is a structural error; a formula in any required or extra data cell makes that record invalid while retaining the expression for traceability.

Protected XML parsing through defusedxml must be active before XLSX processing. If that runtime protection is unavailable or disabled, fail structurally instead of processing the source with weaker silent defaults.

## Traceability and Auditability

Attach `source_file`, `source_sheet`, and `source_row` to every loaded record as early as possible. `source_sheet` may be empty or null for CSV records. Keep this metadata associated with the record throughout processing.

`source_row` is the physical 1-based row number in the original source, including the header row. When the header is on row 1, the first data record has `source_row = 2`.

The names `source_file`, `source_sheet`, and `source_row` are reserved for application-generated traceability metadata. If an input source already contains any reserved name, the source has a structural error. The application must not overwrite, rename, or discard the client's column silently.

A processing timestamp is run-level metadata and must not be duplicated on every record. Normalized values are used for processing and normal output; do not copy every original field solely to preserve formatting when source metadata can locate the original record.

No loaded record may disappear without an explicit classification or documented processing error.

## Failure Policy

A structural or operational failure in any input source makes the run unsuccessful. Log the failure clearly, exit with a non-zero status, and do not publish an apparently complete final report after skipping the source. Partial-processing mode is outside Version 1 scope.

The run also fails with a non-zero exit status and does not generate a new final report when the input directory does not exist, no supported input files exist, or no usable data source can be found. Version 1 does not allow a successful empty report.

Row-level validation failures are retained as invalid records and do not by themselves constitute structural source failures.

## Reporting

The report contains exactly `Summary`, `Valid Records`, `Invalid Records`, and `Duplicates`, in that order. Valid records are both valid and unique. The total paid amount includes only records that are valid, unique, and normalized to `status == "paid"`. Invalid and duplicate records may overlap and appear in both corresponding worksheets.

Record worksheets use the approved business-field order followed by the case-sensitive sorted global union of extra columns, source metadata, and validation errors where applicable. No input column may be discarded or overwritten. Multiple validation errors are serialized in their existing order as `field [code]: message`, separated by ` | `.

Export monetary values as exact canonical two-place text without float conversion. Export dates as native spreadsheet dates with a stable `yyyy-mm-dd` format. Preserve identifiers as text.

An existing `data/output/sales_report.xlsx` may be replaced. Generate the new workbook in a temporary file and atomically replace the final path only after successful generation.

Atomic replacement prevents an expected intermediate destination state when supported by the filesystem. It does not guarantee durability across power loss, kernel crash, or physical storage failure.

Input-controlled text whose first significant character after initial ASCII whitespace is `=`, `+`, `-`, or `@` must be prefixed with an apostrophe and exported as text. This includes retained `FormulaValue` expressions, source metadata, extra columns, and extra-column names. The mitigation must prevent formula execution while preserving the underlying data as much as reasonably possible.

Output must not silently truncate, remove, coerce, replace, or ignore data. XLSX limits, incompatible XML content, and unexpected value types cause an explicit export failure before publication. Logical workbook content is deterministic for the same processed records and summary; byte-identical XLSX archives are not required.

Negative signed float zero is not representable under the approved exact-output guarantees and causes an explicit export failure. Positive float zero remains supported where native float values are otherwise allowed.

## Pending Business Decisions

None. New ambiguities discovered during implementation must be documented and submitted for approval rather than resolved by assumption.

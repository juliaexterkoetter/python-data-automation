# Decision Log

Statuses used in this log are `proposed`, `accepted`, and `superseded`. Proposed entries require explicit approval before they govern implementation.

## DEC-001 — File-Based Processing Scope

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** The requirements define batch processing of sales files from a fixed input directory into an Excel report.
- **Decision:** Build a maintainable Python file-processing application for CSV and XLSX inputs, producing `data/output/sales_report.xlsx`.
- **Consequences:** The initial architecture does not require a web interface, database, or distributed processing system.

## DEC-002 — No Silent Data Discard

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Input data may be invalid or duplicated, and the requirements prohibit silently discarding records.
- **Decision:** Invalid and duplicate data must remain reviewable, with clear classification or documented processing errors.
- **Consequences:** Processing must maintain record accounting and traceability through the pipeline.

## DEC-003 — Module Responsibility Boundaries

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** The repository provides `main.py`, `processor.py`, `validator.py`, and `exporter.py`, and the proposed responsibilities have been reviewed.
- **Decision:** Use those modules as the initial boundaries described in `ARCHITECTURE.md`.
- **Consequences:** The design remains simple and avoids unnecessary layers; material boundary changes require justification and approval.

## DEC-004 — Automated Verification

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Automated tests are an explicit reliability and scope requirement.
- **Decision:** Cover behavior changes with relevant automated tests, including the minimum cases in `requirements.md`.
- **Consequences:** Implementation work is not complete until relevant tests have been run and reviewed.

## DEC-005 — Duplicate Classification and Summary Accounting

- **Status:** superseded
- **Date:** 2026-08-20
- **Context:** The original requirements did not fully define duplicate classification or overlapping summary categories.
- **Decision:** Replace this aggregate proposal with accepted decisions D01–D03.
- **Consequences:** Duplicate and summary behavior is now governed by the more specific decisions.

## DEC-006 — Parsing and Normalization Policies

- **Status:** superseded
- **Date:** 2026-08-20
- **Context:** Date, money, status, identifier, text, and email behavior required approval.
- **Decision:** Replace this aggregate proposal with accepted decisions D04–D16.
- **Consequences:** Parsing and normalization behavior is now governed by the more specific decisions.

## DEC-007 — Input and Failure Policies

- **Status:** superseded
- **Date:** 2026-08-20
- **Context:** Source contracts, workbook selection, malformed-file behavior, and extra columns required approval.
- **Decision:** Replace this aggregate proposal with accepted decisions D17–D21.
- **Consequences:** Input and failure behavior is now governed by the more specific decisions.

## DEC-008 — Output, Traceability, and Paid Total Policies

- **Status:** superseded
- **Date:** 2026-08-20
- **Context:** Traceability, overwrite safety, formula injection, and paid-total eligibility required approval.
- **Decision:** Replace this aggregate proposal with accepted decisions D10 and D22–D24.
- **Consequences:** Output, traceability, and paid-total behavior is now governed by the more specific decisions.

## D01 — Duplicate Classification

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Selecting one occurrence of a repeated order as valid could preserve the wrong record.
- **Decision:** When an `order_id` occurs more than once across the loaded dataset, classify every occurrence as a duplicate.
- **Consequences:** No occurrence of a repeated ID is eligible for the valid-record dataset.

## D02 — Duplicate and Invalid Overlap

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Duplicate identity and row validity are independent properties.
- **Decision:** Allow a record to be both duplicate and invalid and to appear in both corresponding worksheets.
- **Consequences:** Classification datasets are intentionally non-exclusive.

## D03 — Summary Duplicate Semantics

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Overlapping classifications make additive category totals misleading.
- **Decision:** Report duplicate count as an independent metric and make the overlap semantics clear.
- **Consequences:** The summary must not imply that total records equal valid plus invalid plus duplicates.

## D04 — `order_id` Representation

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Numeric inference can remove leading zeros and change identifier meaning.
- **Decision:** Treat `order_id` as text and preserve leading zeros.
- **Consequences:** Loading and export must prevent silent numeric conversion of identifiers.

## D05 — Date Format

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** A canonical date contract is needed across text and Excel sources.
- **Decision:** Accept textual `order_date` values in `YYYY-MM-DD` format and native Excel date or datetime cells.
- **Consequences:** Validation and normalization have an explicit accepted date representation.

## D06 — Ambiguous Dates

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Locale-dependent date guessing can produce incorrect order dates.
- **Decision:** Reject ambiguous or unsupported textual date formats without guessing their interpretation.
- **Consequences:** Values such as `01/02/2026` are invalid.

## D07 — Monetary Format and Currency

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Locale and currency ambiguity can change monetary values.
- **Decision:** Version 1 uses USD and canonical decimal notation without currency symbols or locale-dependent representations.
- **Consequences:** Other currencies and localized monetary strings are outside the Version 1 input contract.

## D08 — Monetary Precision

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Binary floating-point arithmetic can introduce monetary errors.
- **Decision:** Use decimal arithmetic, two decimal places, and `ROUND_HALF_UP` when rounding is required.
- **Consequences:** Parsing, validation, aggregation, and tests must avoid binary floating-point calculations.

## D09 — Zero and Negative Amounts

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** The permitted range of order amounts was undefined.
- **Decision:** Accept zero, reject negative amounts, and represent refunds with the `refunded` status.
- **Consequences:** Negative values create row-level validation errors.

## D10 — Paid Total

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Summary eligibility must exclude records that should not contribute to paid revenue.
- **Decision:** Include only valid, unique records normalized to `status == "paid"` in total paid amount.
- **Consequences:** Invalid, duplicated, pending, cancelled, and refunded records contribute nothing to this total.

## D11 — Status Normalization

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Whitespace and capitalization should not create false unknown statuses.
- **Decision:** Trim status values and convert them to lowercase before validation.
- **Consequences:** Variants such as `" PAID "` normalize to `paid`.

## D12 — Customer Name Normalization

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** External whitespace is inconsistent, while capitalization may be meaningful.
- **Decision:** Trim customer names without changing their remaining capitalization.
- **Consequences:** The application must not automatically title-case or lowercase names.

## D13 — `order_id` Normalization

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Identifier normalization must remove accidental whitespace without changing identity.
- **Decision:** Treat IDs as text, trim external whitespace, and preserve capitalization and leading zeros.
- **Consequences:** Duplicate matching remains case-sensitive.

## D14 — Email Behavior

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Email is optional, but present values should be normalized and obviously malformed addresses rejected.
- **Decision:** Trim and lowercase present email values and apply basic pragmatic, non-RFC-complete format validation.
- **Consequences:** Missing email is valid; present malformed email makes the record invalid.

## D15 — Empty Strings

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Whitespace-only values should not satisfy required-field rules.
- **Decision:** Treat values that are empty after normalization as missing when evaluating required fields.
- **Consequences:** Required whitespace-only values create row-level validation errors.

## D16 — Normalized Values and Auditability

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Auditability is needed without unnecessary copies of every input field.
- **Decision:** Use normalized values in processing and normal output while retaining enough source metadata to locate the original record.
- **Consequences:** Original-format mirror columns are not required without a concrete future need.

## D17 — CSV Input Contract

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Silent encoding or delimiter guessing can corrupt data interpretation.
- **Decision:** Version 1 accepts header-bearing, comma-delimited UTF-8 CSV files and matches extensions case-insensitively.
- **Consequences:** Other encodings or delimiters are not guessed and cause source failure.

## D18 — XLSX Worksheet Selection

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Workbooks may contain multiple worksheets, and silently ignoring structurally relevant content risks data loss.
- **Decision:** Process worksheets containing the required schema and preserve each worksheet name as source metadata.
- **Consequences:** A worksheet that cannot satisfy the structural contract must not silently yield zero records.

## D19 — Missing Required Columns

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** A missing column is a source-structure problem rather than a defect in one row.
- **Decision:** Treat missing required columns as structural source errors.
- **Consequences:** These failures cannot be represented merely as invalid records.

## D20 — Extra Columns

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Additional source data may be useful for downstream review and traceability.
- **Decision:** Allow and preserve extra input columns.
- **Consequences:** Extra columns alone never cause source rejection.

## D21 — Partial Processing Policy

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** A report that silently omits failed sources can appear complete while containing incorrect totals.
- **Decision:** Any structural or operational source failure makes the run unsuccessful, requires clear logging and a non-zero exit status, and prevents publication of an apparently complete report.
- **Consequences:** Explicit partial-processing mode is outside Version 1 scope.

## D22 — Source Traceability

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Every result must be attributable to its source without duplicating run metadata per row.
- **Decision:** Attach `source_file`, `source_sheet`, and `source_row` during loading and retain them; keep processing timestamp at run level.
- **Consequences:** CSV records may have an empty or null `source_sheet`, and all record transformations preserve the metadata.

## D23 — Safe Report Replacement

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Failed report generation must not destroy the last valid output.
- **Decision:** Permit overwrite through temporary-file generation followed by atomic replacement only after success.
- **Consequences:** Export failures leave the existing valid report intact whenever supported by the filesystem contract.

## D24 — Excel Formula Injection

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Input-controlled text can be interpreted as executable spreadsheet formulas.
- **Decision:** Safely export text beginning with formula-triggering characters such as `=`, `+`, `-`, or `@` while preserving its underlying value as much as reasonably possible.
- **Consequences:** Formula-injection mitigation is required across required and extra textual columns.

## D25 — XLSX Worksheet Classification

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Worksheet selection did not distinguish auxiliary content from a malformed data worksheet.
- **Decision:** Ignore and log worksheets containing no required columns, fail structurally on partial required schemas, process complete schemas, and fail workbooks with no usable data worksheet.
- **Consequences:** Auxiliary sheets are handled visibly while incomplete data sources and unusable workbooks cannot produce misleading output.

## D26 — Monetary Input Precision

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** The behavior of finite decimal inputs with more than two decimal places was undefined.
- **Decision:** Accept excess precision, parse with `Decimal`, and normalize to two places using `ROUND_HALF_UP`.
- **Consequences:** Excess precision alone is valid; values such as `1.005`, `12.999`, and `149.9` normalize to `1.01`, `13.00`, and `149.90`.

## D27 — No Usable Input

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** A successful empty report could falsely appear to represent a complete run.
- **Decision:** Fail with non-zero status and publish no new final report when the input directory is missing, no supported files exist, or no usable data source is found.
- **Consequences:** Version 1 cannot complete successfully with empty input.

## D28 — Basic Email Validation

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Pragmatic email validation needed deterministic minimum boundaries.
- **Decision:** For present emails, require one `@` between non-empty local and domain portions, no whitespace, and a plausible dotted domain; do not attempt RFC-complete validation.
- **Consequences:** `julia@example.com` is valid, while `julia`, `julia@`, `@example.com`, `julia example.com`, and `julia@example` are invalid.

## D29 — `source_row` Semantics

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Traceability requires a row value that maps directly to the physical source.
- **Decision:** Use the physical 1-based row number including the header row; a first data record below a row-1 header has `source_row = 2`.
- **Consequences:** Users can locate records directly in their source files.

## D30 — Reserved Traceability Columns

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Input columns could collide with application-generated traceability metadata.
- **Decision:** Reserve `source_file`, `source_sheet`, and `source_row`; any input collision is a structural source error.
- **Consequences:** Client columns are never silently overwritten, renamed, or discarded.

## D31 — UTF-8 BOM

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** A UTF-8 BOM could otherwise become part of the first CSV header name.
- **Decision:** Accept UTF-8 CSV files with or without BOM by reading them with `utf-8-sig`; reject other encodings.
- **Consequences:** A BOM is removed safely without silently broadening the encoding contract.

## D32 — CSV Header Names

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** Silent header correction could hide structural defects or change client data meaning.
- **Decision:** Validate CSV header names exactly and case-sensitively without trimming, lowercasing, renaming, or correcting them.
- **Consequences:** Only exact required names satisfy the schema; variants such as `Order_ID` and `ORDER_ID` do not.

## D33 — Input Directory Recursion

- **Status:** accepted
- **Date:** 2026-08-20
- **Context:** The Version 1 discovery boundary needed an explicit rule for nested directories.
- **Decision:** Discover only regular CSV files directly inside `data/input/`; do not recurse into subdirectories.
- **Consequences:** Nested files are ignored and discovery remains deterministic and narrowly scoped.

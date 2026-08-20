# Project Tasks

## Blocked — Business Decisions

None.

## Planned

1. Plan the next implementation increment against the approved rules and architecture.
2. Add runtime dependencies only when required by an approved increment.
3. Implement classified multi-worksheet XLSX loading with exact source metadata and preserved extra columns.
4. Implement summary calculation and paid-total eligibility.
5. Implement Excel export with formula-injection protection and atomic replacement.
6. Extend logging, explicit errors, and non-zero failure status with each pipeline stage.
7. Create unit tests alongside each behavior.
8. Create integration and failure-path tests for later pipeline stages.
9. Create representative sample input data.
10. Validate complete end-to-end execution and record accounting.
11. Complete the README after the application is functional.

## Completed

- Defined the initial functional requirements.
- Reframed project scope for a real, maintainable application.
- Analyzed requirement ambiguities, reliability risks, and testing needs.
- Established the permanent documentation structure.
- Approved and documented business decisions D01–D36.
- Approved the initial module responsibility boundaries.
- Resolved pre-merge ambiguities for XLSX worksheet classification, monetary precision, empty input, email validation, source-row semantics, and reserved traceability columns.
- Implemented non-recursive, case-insensitive CSV discovery with explicit missing-input failures.
- Implemented UTF-8 and UTF-8 BOM structural CSV loading with exact headers, comma delimiters, required and reserved-column checks, textual identifiers, and source metadata.
- Added explicit structural exceptions, operational logging, exit-status coordination, and automated tests for the CSV increment.
- Implemented field normalization and row-level validation for required values, email, status, canonical dates, and decimal amounts.
- Implemented global duplicate detection and non-exclusive valid, invalid, and duplicate classifications while preserving traceability and extra columns.

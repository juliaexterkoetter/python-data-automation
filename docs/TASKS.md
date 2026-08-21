# Project Tasks

## Blocked — Business Decisions

None.

## Planned

1. Perform the manual Microsoft Excel smoke test.
2. Optionally capture the separate manual Excel smoke-test evidence from the real workbook.
3. Publish the prepared repository, LinkedIn, and Upwork presentation material as appropriate.

## Completed

- Defined the initial functional requirements.
- Reframed project scope for a real, maintainable application.
- Analyzed requirement ambiguities, reliability risks, and testing needs.
- Established the permanent documentation structure.
- Approved and documented business decisions D01–D66.
- Approved the initial module responsibility boundaries.
- Resolved pre-merge ambiguities for XLSX worksheet classification, monetary precision, empty input, email validation, source-row semantics, and reserved traceability columns.
- Implemented non-recursive, case-insensitive CSV discovery with explicit missing-input failures.
- Implemented UTF-8 and UTF-8 BOM structural CSV loading with exact headers, comma delimiters, required and reserved-column checks, textual identifiers, and source metadata.
- Added explicit structural exceptions, operational logging, exit-status coordination, and automated tests for the CSV increment.
- Implemented field normalization and row-level validation for required values, email, status, canonical dates, and decimal amounts.
- Implemented global duplicate detection and non-exclusive valid, invalid, and duplicate classifications while preserving traceability and extra columns.
- Implemented deterministic CSV and XLSX discovery with shared filesystem safety rules.
- Implemented classified multi-worksheet XLSX loading, native value handling, formula rejection, and exact source traceability.
- Added mixed-format integration and XLSX structural, native-type, formula, and regression tests.
- Implemented immutable summary calculation with independent classification counts and valid, unique paid-total eligibility.
- Added exact Decimal aggregation for zero, normalized values, and multiple very large amounts without float conversion.
- Added summary unit and mixed-source integration coverage, including overlap, non-eligible records, non-mutation, and structural-failure boundaries.
- Approved and documented output decisions D48–D58 for workbook structure, columns, errors, exact money, formula safety, atomic publication, no silent loss, determinism, strict types, empty CSV headers, and negative signed float zero.
- Implemented the deterministic four-worksheet Excel report without recalculating processing results or summary values.
- Implemented exact monetary text, collision-safe error columns, safe formula-like text, stable date output, and complete traceability and extra-column preservation.
- Implemented same-directory temporary generation, workbook validation, atomic replacement, cleanup, explicit export errors, and non-zero main-pipeline failure status.
- Added workbook round-trip, projection, formula-safety, technical-limit, determinism, end-to-end, and fault-injection tests.
- Hardened temporary validation against logically incomplete but structurally valid workbooks and rejected untraceable empty headers and lossy negative signed float zero.
- Replaced the workbook-derived expectation with an immutable logical model built directly from processing projections and summary, plus independent saved-workbook invariants.
- Added protected XML parsing, non-extracting XLSX package preflight, resource limits, disabled external-link caches, and deterministic CSV field-size enforcement.
- Added safe boundary and adversarial regression coverage for D59–D63 without creating large or dangerous fixtures.
- Approved and documented release-readiness decisions D64–D66.
- Pinned the validated direct runtime and development dependencies.
- Documented the dependency-audit procedure and result, atomicity-versus-durability boundary, trusted input directory, residual TOCTOU risk, and manual Excel smoke-test checklist.
- Added and programmatically validated a fully fictitious mixed CSV/XLSX demonstration dataset.
- Added a dedicated demo runner and regression test for documented summary, classification, traceability, and formula-safety expectations.
- Completed the portfolio-oriented README, LinkedIn and Upwork copy, GitHub presentation suggestions, and real-screenshot plan.
- Added four reproducible programmatic portfolio renderings derived from the real demonstration workbook, with deterministic content validation independent from the manual Excel evidence plan.

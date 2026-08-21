# Python Sales Data Automation

A defensive Python workflow that consolidates sales records from CSV and Excel, validates and classifies every record, and publishes an auditable Excel report.

## What problem does this solve?

Operations teams often receive sales data in separate CSV and Excel files. Those files can contain duplicated orders, missing fields, malformed values, inconsistent text, extra columns, and records that must remain traceable to their source. Manual consolidation is slow and can silently lose or misclassify data.

This project turns that workflow into a repeatable pipeline with explicit structural failures, record-level validation, duplicate review, source metadata, deterministic summaries, and a validated workbook output.

## What does the application do?

```text
CSV and XLSX inputs
    -> structural and resource validation
    -> field normalization
    -> record validation
    -> global duplicate detection
    -> valid / invalid / duplicate classification
    -> summary calculation
    -> validated Excel report
```

The report contains exactly `Summary`, `Valid Records`, `Invalid Records`, and `Duplicates`, in that order. Invalid and duplicate classifications may overlap, so a problematic repeated order remains visible in both review views.

## Key capabilities

- Deterministic, non-recursive CSV and XLSX discovery.
- UTF-8 CSV loading with exact schema validation and a 128 KiB field limit.
- Multi-worksheet XLSX loading with auxiliary-sheet classification.
- Native Excel date handling and exact `Decimal` monetary normalization.
- Required-field, email, date, amount, and status validation.
- Global, case-sensitive duplicate detection with every repeated occurrence retained.
- `source_file`, `source_sheet`, and physical `source_row` traceability.
- Preservation of approved extra columns.
- Exact paid-order summary calculation from valid, unique records.
- Deterministic four-sheet Excel report with headers, filters, frozen rows, and stable dates.

## Reliability and security

- XML parsing requires active `defusedxml` protection.
- XLSX packages receive ZIP/OOXML preflight before `openpyxl.load_workbook()`.
- Content Types, workbook relationships, worksheet targets, CRC integrity, compression methods, expansion ratios, and operational dimensions are checked without extracting members.
- External-link caches are disabled and external resources are not accessed.
- Formula-like input is exported as protected text and verified after saving.
- Unsupported or unrepresentable data fails explicitly instead of being truncated or silently coerced.
- The saved workbook is reopened and compared with an independently built logical model before publication.
- Publication uses a validated temporary file and `os.replace` so normal pre-replacement failures preserve the previous report.

See [release readiness](docs/RELEASE_READINESS.md) for the exact security, atomicity, durability, dependency, and trust-boundary claims.

## Architecture

- `src.processor`: source discovery, structural loading, traceability, and processing coordination.
- `src.xlsx_safety`: protected XLSX ZIP/OOXML preflight and operational limits.
- `src.validator`: normalization and row-level validation.
- `src.summary`: immutable summary calculation.
- `src.exporter`: logical workbook construction, formula safety, validation, and atomic publication.
- `src.main`: application orchestration, logging, and exit status.

Approved rules and architectural decisions are documented in [business rules](docs/BUSINESS_RULES.md), [architecture](docs/ARCHITECTURE.md), the [decision log](docs/DECISIONS.md), and the [testing strategy](docs/TESTING.md).

## Installation

Version 1 was validated with Python 3.14.4, openpyxl 3.1.5, defusedxml 0.7.1, and pytest 9.1.1.

```bash
git clone https://github.com/juliaexterkoetter/python-data-automation.git
cd python-data-automation
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements-dev.txt
.venv/bin/python -m pip check
```

Formal dependency-specific support for Python 3.14 is not verified for every runtime dependency; compatibility is demonstrated by the complete project suite in the validated environment.

## Running

Place trusted operator-controlled `.csv` and `.xlsx` files directly in `data/input/`, then run:

```bash
.venv/bin/python -m src.main
```

The application writes `data/output/sales_report.xlsx`. A structural input or export failure returns a non-zero exit status and does not publish an apparently successful partial report.

## Demo

The repository includes a small, fully fictitious mixed-format scenario under `data/demo/input/`.

```bash
.venv/bin/python -m scripts.run_demo
```

The generated workbook is written to `data/demo/output/sales_report.xlsx`. Expected counts and records are documented in [the demo guide](data/demo/README.md). Generated output is intentionally ignored by Git.

## Tests

```bash
.venv/bin/python -m pytest
.venv/bin/python -m compileall src tests
.venv/bin/python -m pip check
git diff --check
```

The suite covers the complete CSV/XLSX pipeline, validation, duplicates, classification overlap, summaries, formula-injection protection, package hardening, workbook round-trip validation, fault injection, and atomic publication behavior.

## Limitations

- Operational limits are intentionally sized for small and medium local automation workloads; see [business rules](docs/BUSINESS_RULES.md).
- `data/input/` is a trusted local operator directory. Concurrent hostile local writes and the residual TOCTOU window are outside Version 1.
- Atomic replacement is not a power-loss, kernel-crash, or physical-storage durability guarantee.
- The dependency versions work in the validated Python 3.14.4 environment, but formal Python 3.14 support is not independently verified for every dependency.
- Automated tests validate workbook structure and values, but the Microsoft Excel for Windows smoke test remains manual until a person completes [the checklist](docs/EXCEL_SMOKE_TEST.md).

## Project documentation

- [Project overview](docs/PROJECT.md)
- [Requirements](docs/requirements.md)
- [Business rules](docs/BUSINESS_RULES.md)
- [Architecture](docs/ARCHITECTURE.md)
- [Decision log](docs/DECISIONS.md)
- [Development workflow](docs/DEVELOPMENT.md)
- [Testing strategy](docs/TESTING.md)
- [Release readiness](docs/RELEASE_READINESS.md)
- [Manual Excel smoke test](docs/EXCEL_SMOKE_TEST.md)
- [Portfolio materials](docs/PORTFOLIO_MATERIALS.md)
- [Screenshot plan](docs/SCREENSHOT_PLAN.md)
- [Task backlog](docs/TASKS.md)

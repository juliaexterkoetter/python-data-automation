# Python Data Automation

A Python application for consolidating, validating, cleaning, and reporting sales records received in CSV and Excel files, with an emphasis on correctness, traceability, maintainability, and explicit error handling.

## Development Status

The approved Version 1 pipeline and release-readiness baseline are implemented: CSV and XLSX discovery and resource-bounded structural loading, record normalization, row-level validation, duplicate detection, record classification, summary calculation, and failure-safe Excel report publication. Demonstration assets and presentation-oriented content remain separate follow-up work.

Version 1 was validated with Python 3.14.4, openpyxl 3.1.5, defusedxml 0.7.1, and pytest 9.1.1. The full Python 3.14 support status of every runtime dependency is not independently verified; compatibility is demonstrated by the project test suite in the validated environment.

## Documentation

- [Project overview](docs/PROJECT.md)
- [Requirements](docs/requirements.md)
- [Business rules](docs/BUSINESS_RULES.md)
- [Architecture](docs/ARCHITECTURE.md)
- [Decision log](docs/DECISIONS.md)
- [Development workflow](docs/DEVELOPMENT.md)
- [Testing strategy](docs/TESTING.md)
- [Release readiness](docs/RELEASE_READINESS.md)
- [Manual Excel smoke test](docs/EXCEL_SMOKE_TEST.md)
- [Task backlog](docs/TASKS.md)

## Reproducible Setup

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements-dev.txt
.venv/bin/python -m pytest
```

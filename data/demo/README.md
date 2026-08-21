# Demonstration Scenario

All names, emails, identifiers, dates, and amounts in this directory are fictitious and exist only to demonstrate the application.

## Run

```bash
.venv/bin/python -m scripts.run_demo
```

The command reads `data/demo/input/` and writes `data/demo/output/sales_report.xlsx`.

## Expected Summary

| Metric | Expected value |
| --- | ---: |
| Total Records | 11 |
| Valid Records | 5 |
| Invalid Records | 4 |
| Duplicate Records | 4 |
| Total Paid Amount (USD) | 34.46 |

Invalid and duplicate counts overlap. Both occurrences of `BOTH-200` are invalid and duplicated. Both occurrences of `DUP-100` are valid at row level but excluded from valid output because the identifier is repeated.

## Behaviors Demonstrated

- CSV and XLSX input in one run;
- two data worksheets and one auxiliary worksheet;
- paid, pending, cancelled, and refunded statuses;
- valid, invalid, duplicate, and invalid-plus-duplicate records;
- textual identifier `0007` with leading zeros;
- zero amount;
- `ROUND_HALF_UP` normalization (`10.005` to `10.01` and `4.445` to `4.45`);
- email normalization and invalid email reporting;
- preserved `note` extra column;
- exact source file, worksheet, and physical row metadata;
- CSV formula-like text and an XLSX formula retained as protected, non-executable text.

The generated workbook is not versioned. Validate it programmatically and complete the manual Microsoft Excel checklist before using screenshots in a portfolio.

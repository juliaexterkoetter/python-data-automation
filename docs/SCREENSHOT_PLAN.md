# Portfolio Screenshot Plan

Use only the fictitious demonstration data. Before capturing, hide usernames, unrelated applications, local personal paths, notifications, account details, and any real data.

## Programmatic portfolio renderings

The reproducible assets in `portfolio/screenshots/` are generated from the real `data/demo/output/sales_report.xlsx` workbook by `scripts/render_portfolio_screenshots.py`. They are presentation renderings for LinkedIn, Upwork, GitHub, and other portfolio use; they are not screenshots of the Microsoft Excel interface.

- `01-summary.png`: the five Summary metrics and overlap note.
- `02-valid-records.png`: selected columns from all five valid records.
- `03-invalid-records.png`: selected columns and validation feedback from all four invalid records.
- `04-duplicates.png`: selected columns and validation feedback from all four duplicate occurrences.

Generate and validate them with the dedicated portfolio environment:

```bash
.venv/bin/python -m pip install -r requirements-portfolio.txt
.venv/bin/python -m scripts.render_portfolio_screenshots
.venv/bin/python -m scripts.render_portfolio_screenshots --validate-only
```

Validation regenerates the expected assets in a temporary directory and requires byte-identical output. It does not use OCR or modify the source workbook.

## Manual capture and Excel smoke-test evidence

The following capture plan remains optional and pending. Real Microsoft Excel views provide manual compatibility and visual evidence and are not replaced by the programmatic renderings above. Do not claim that this evidence exists until the corresponding application check has actually been completed.

## 01-input-files

- Window: file explorer or terminal directory listing.
- Show: `orders.csv` and `orders.xlsx` inside `data/demo/input/`.
- Visible content: filenames only; optionally a small safe preview of the fictitious CSV.
- Hide: absolute personal paths and unrelated files.
- Suggested filename: `01-input-files.png`.
- Caption: “Fictitious CSV and multi-worksheet Excel inputs used by the automation.”

## 02-terminal-success

- Window: terminal after `.venv/bin/python -m scripts.run_demo`.
- Show: discovery counts, loaded-record count, classification summary, and successful publication message.
- Hide: username, machine name, unrelated command history, and absolute personal paths when possible.
- Suggested filename: `02-terminal-success.png`.
- Caption: “One command validates, consolidates, classifies, summarizes, and publishes the report.”

## 03-summary

- Window: Microsoft Excel for Windows.
- Worksheet: `Summary`.
- Show: all five metrics and the overlap note.
- Visible values: 11 total, 5 valid, 4 invalid, 4 duplicate, USD 34.46 paid.
- Hide: Excel account details and unrelated workbooks.
- Suggested filename: `03-summary.png`.
- Caption: “Validated processing totals with independent invalid and duplicate counts.”

## 04-invalid-records

- Window: Microsoft Excel for Windows.
- Worksheet: `Invalid Records`.
- Show: `order_id`, representative business fields, `note`, source metadata, and `validation_errors`.
- Visible examples: `BAD-EMAIL`, both `BOTH-200` records, and `FORMULA-1`.
- Hide: columns only if needed for readability; never hide a value required to understand the example.
- Suggested filename: `04-invalid-records.png`.
- Caption: “Invalid records are retained with source metadata and explicit validation reasons.”

## 05-duplicates

- Window: Microsoft Excel for Windows.
- Worksheet: `Duplicates`.
- Show: both `DUP-100` and both `BOTH-200` occurrences plus source metadata and errors.
- Visible columns: `order_id`, `customer_name`, `status`, `source_file`, `source_sheet`, `source_row`, and `validation_errors`.
- Hide: nothing that would obscure why every repeated occurrence is retained.
- Suggested filename: `05-duplicates.png`.
- Caption: “Every repeated identifier remains visible, including records that are also invalid.”

The worksheet screenshots in this manual set must come from the real generated workbook opened in Microsoft Excel. Do not synthesize Excel interface screenshots.

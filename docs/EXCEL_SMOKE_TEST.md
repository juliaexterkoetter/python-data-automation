# Manual Microsoft Excel Smoke Test

Status: **WAITING FOR MANUAL EXCEL SMOKE TEST**

Use a report generated from fictitious demonstration data. Do not use client, company, or personal data in screenshots or portfolio material.

Generate the report from the repository root:

```bash
.venv/bin/python -m scripts.run_demo
```

From WSL, open the generated file with its Windows file association:

```bash
explorer.exe "$(wslpath -w data/demo/output/sales_report.xlsx)"
```

Confirm that the associated application is Microsoft Excel before performing the checklist.

## Checklist

- [ ] 1. Open `sales_report.xlsx` in Microsoft Excel for Windows without a repair or corruption warning.
- [ ] 2. Confirm the worksheets are `Summary`, `Valid Records`, `Invalid Records`, and `Duplicates`.
- [ ] 3. Confirm the worksheets appear in that exact order.
- [ ] 4. Confirm record-sheet headers are bold.
- [ ] 5. Confirm record sheets freeze the first row.
- [ ] 6. Confirm record sheets have autofilters covering their data.
- [ ] 7. Confirm dates display consistently as `YYYY-MM-DD`.
- [ ] 8. Confirm monetary values display with exactly two decimal places.
- [ ] 9. Confirm the Summary values match the expected demonstration totals.
- [ ] 10. Confirm textual IDs retain leading zeros.
- [ ] 11. Confirm `source_file`, `source_sheet`, and `source_row` are present and correct.
- [ ] 12. Confirm a record that is both invalid and duplicated appears in both applicable worksheets.
- [ ] 13. Confirm formula-like input is displayed as protected text and does not execute.
- [ ] 14. Exercise the filters on each record worksheet.
- [ ] 15. Check normal worksheet navigation and frozen-row behavior.
- [ ] 16. Save the workbook in Excel.
- [ ] 17. Close Excel.
- [ ] 18. Reopen the saved workbook without a repair or corruption warning.
- [ ] 19. Run the application again with the same demonstration input.
- [ ] 20. Confirm the existing report is replaced normally and the new report opens successfully.

Record the Excel version, Windows version, date, tester, and any warning or visual discrepancy. Do not mark the smoke test approved until every item has been checked in Microsoft Excel itself.

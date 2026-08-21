# Portfolio Materials

## LinkedIn Project

### Project title

Python Sales Data Automation and Excel Reporting

### Short description

Built a Python automation workflow that consolidates sales records from CSV and Excel files, validates and normalizes the data, identifies every duplicate occurrence, preserves invalid records for review, and produces a structured Excel report with traceable source metadata. The project emphasizes practical reliability through automated tests, exact decimal handling, explicit error reporting, protected spreadsheet text, and validated report publication.

### Skills

- Python
- Process Automation
- CSV and Excel Processing
- Data Validation
- Data Cleaning
- Data Quality
- Automated Testing
- Report Automation
- Error Handling

### Link

<https://github.com/juliaexterkoetter/python-data-automation>

## Upwork Portfolio

### Portfolio title

Python CSV and Excel Sales Data Automation

### Short description

An automated workflow for turning mixed CSV and Excel sales files into a reliable review-ready report. It consolidates files, cleans approved fields, flags invalid records, captures all duplicate orders, calculates business totals, and preserves the source of every record.

### Full case study

#### Problem

Manual consolidation of recurring CSV and Excel files is slow and error-prone. Missing values, inconsistent formatting, invalid amounts, duplicate order IDs, and mixed worksheets can make totals unreliable and make mistakes difficult to trace.

#### Solution

I built a Python pipeline that applies one documented processing contract to every supported file and produces a consistent four-sheet Excel report. Structural file problems stop the run clearly, while row-level problems remain visible for review.

#### What the automation does

- discovers CSV and XLSX inputs;
- validates file and worksheet structure;
- normalizes approved text, dates, status values, and amounts;
- validates each record without losing invalid rows;
- identifies every occurrence of duplicated order IDs;
- separates valid, invalid, and duplicate views;
- calculates a paid-order summary from valid unique records;
- attaches source file, worksheet, and row metadata;
- generates and revalidates the final Excel report.

#### Quality and reliability

The workflow uses exact decimal arithmetic, automated regression and fault-injection tests, explicit structural errors, formula-safe spreadsheet text, resource-bounded XLSX inspection, logical workbook verification, and temporary-file publication that preserves an existing report across normal failures before replacement.

#### Deliverables

- maintainable Python source code;
- automated tests;
- documented processing and validation rules;
- fictitious demonstration inputs;
- validated Excel report generation;
- operating, release, and manual Excel verification guidance.

#### Technologies

Python, openpyxl, defusedxml, pytest, CSV, XLSX/OOXML, Decimal, and standard-library filesystem and ZIP tooling.

### Skills and tags

- Python
- Automation
- Excel
- CSV
- Data Processing
- Data Cleaning
- Data Validation
- Report Automation

### Screenshot captions

- `01-input-files.png`: “Fictitious CSV and Excel source files ready for one automated run.”
- `02-terminal-success.png`: “Clear processing logs and successful report publication.”
- `03-summary.png`: “Business totals calculated from validated, unique records.”
- `04-invalid-records.png`: “Invalid records retained with explicit reasons and source traceability.”
- `05-duplicates.png`: “All duplicate occurrences preserved for review.”

## GitHub Presentation

### Repository description

Defensive Python automation for validating CSV/XLSX sales data and generating an auditable Excel report.

### Suggested topics

- python
- automation
- excel
- csv
- data-validation
- data-cleaning
- openpyxl
- reporting
- ooxml
- pytest

The README provides the final project overview, demo usage, architecture, reliability boundaries, installation, and testing commands. Repository settings, topics, releases, and tags must be changed manually or through a separately authorized GitHub operation.

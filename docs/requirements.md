# Python Data Automation - Requirements

## Problem

A company receives sales data from multiple CSV and Excel files.

The files may contain:
- duplicated records;
- missing values;
- invalid numeric values;
- inconsistent text formatting;
- different order statuses.

The goal is to automate the consolidation, validation, cleaning and reporting process using Python.

## Input

The application must read files from:

data/input/

Supported formats:
- CSV
- XLSX

Expected fields:

- order_id
- customer_name
- email
- order_date
- amount
- status

## Business Rules

### Unique orders

`order_id` identifies an order.

Duplicated order IDs must not silently disappear.

Duplicated records must be stored separately for review.

### Email normalization

Emails must:
- have leading and trailing spaces removed;
- be converted to lowercase.

### Amount validation

`amount` must be numeric.

Invalid amounts must not crash the application.

Records with invalid amounts must be moved to the invalid records report.

### Required fields

The following fields are required:

- order_id
- customer_name
- order_date
- amount
- status

Records missing required information must be marked as invalid.

### Status

Valid statuses:

- paid
- pending
- cancelled
- refunded

Unknown statuses must be marked as invalid.

## Processing

The application must:

1. Discover supported files in `data/input`.
2. Load all records.
3. Combine them into one dataset.
4. Normalize relevant fields.
5. Validate each record.
6. Separate invalid records.
7. Detect duplicated order IDs.
8. Produce a clean dataset containing valid, unique records.
9. Calculate summary information.
10. Generate an Excel report.

## Output

The application must generate:

data/output/sales_report.xlsx

The Excel workbook must contain:

### Summary

Information such as:

- total records received;
- valid records;
- invalid records;
- duplicate records;
- total amount from paid orders.

### Valid Records

All records that passed validation and are not duplicates.

### Invalid Records

Invalid records plus a column explaining the validation error.

### Duplicates

Records identified as duplicated by `order_id`.

## Reliability

The program should:

- use logging;
- handle malformed input files gracefully;
- avoid silently discarding records;
- produce clear error messages;
- be structured into reusable modules.

## Tests

Automated tests should cover at least:

- duplicate detection;
- email normalization;
- invalid amount handling;
- missing required fields;
- invalid status values;
- summary calculations.

## Scope

The application should provide a reliable and maintainable solution for automating the processing of sales data.

The implementation should prioritize:

- correctness;
- readability;
- maintainability;
- data traceability;
- clear error handling;
- automated testing;
- practical business requirements.

Unnecessary architectural complexity should be avoided.

## Related Documentation

- Detailed approved processing rules are defined in `BUSINESS_RULES.md`.
- The approved implementation structure is described in `ARCHITECTURE.md`.
- Decision status is recorded in `DECISIONS.md`.
- Testing expectations are expanded in `TESTING.md`.

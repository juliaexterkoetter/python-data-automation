# Project Overview

## Purpose

Provide a reliable and maintainable Python application that automates the processing of sales data received in CSV and Excel files.

## Business Problem

Sales records arrive in multiple files and may contain duplicates, missing information, invalid numeric values, inconsistent text formatting, and unsupported order statuses. Manual consolidation and review are error-prone and make it difficult to trace problematic records back to their source.

## Users

The intended users are business or operations personnel who need a consolidated sales report and technical operators responsible for running and supporting the process. More specific user roles have not yet been defined.

## Expected Outcome

Each execution should classify the received records, retain invalid and duplicate data for review, calculate the required summary information, and produce an auditable Excel report without silently losing records.

## High-Level Workflow

1. Discover supported input files.
2. Load and combine their records.
3. Normalize fields covered by approved rules.
4. Validate each record.
5. Classify invalid and duplicate records.
6. Produce the valid, unique dataset.
7. Calculate summary information.
8. Export the Excel report.

Detailed functional requirements are defined in `requirements.md`. Approved processing rules are organized in `BUSINESS_RULES.md`.

## Inputs

- CSV files in `data/input/`.
- XLSX files in `data/input/`.
- The expected fields are defined in `requirements.md`.

## Outputs

- `data/output/sales_report.xlsx`.
- Summary, valid-record, invalid-record, and duplicate-record worksheets as defined in `requirements.md`.

## Reliability Priorities

- Correctness and explicit error handling.
- No silent record loss.
- Source-to-output traceability.
- Clear, maintainable implementation.
- Automated verification of business behavior.
- Practical solutions without unnecessary architectural complexity.

## Current Non-Goals

- A user interface or web service.
- A database-backed system.
- Real-time or distributed processing.
- Features beyond the file-processing workflow in `requirements.md`.

These non-goals describe the current scope and may change only through an approved decision.

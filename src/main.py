"""Application entry point for the complete Version 1 reporting pipeline."""

from __future__ import annotations

import logging
from pathlib import Path

from src.exporter import ReportExportError, export_report
from src.processor import (
    StructuralInputError,
    discover_supported_files,
    load_supported_files,
    process_records,
)
from src.summary import calculate_summary


LOGGER = logging.getLogger(__name__)
DEFAULT_INPUT_DIR = Path("data/input")
DEFAULT_OUTPUT_PATH = Path("data/output/sales_report.xlsx")


def run(
    input_dir: Path = DEFAULT_INPUT_DIR,
    output_path: Path = DEFAULT_OUTPUT_PATH,
) -> int:
    """Process supported sources and publish the complete Excel report."""
    try:
        source_files = discover_supported_files(input_dir)
        csv_count = sum(path.suffix.lower() == ".csv" for path in source_files)
        xlsx_count = sum(path.suffix.lower() == ".xlsx" for path in source_files)
        LOGGER.info("Discovered %d CSV input file(s).", csv_count)
        LOGGER.info("Discovered %d XLSX input file(s).", xlsx_count)
        records = load_supported_files(source_files)
        LOGGER.info(
            "Successfully loaded %d record(s) from %d input file(s).",
            len(records),
            len(source_files),
        )
        result = process_records(records)
        LOGGER.info(
            "Processed %d record(s): %d valid, %d invalid, %d duplicate.",
            len(result.records),
            len(result.valid_records),
            len(result.invalid_records),
            len(result.duplicate_records),
        )
        summary = calculate_summary(result)
        LOGGER.info(
            "Summary: %d total, %d valid, %d invalid, %d duplicate, "
            "paid amount %s.",
            summary.total_records,
            summary.valid_records,
            summary.invalid_records,
            summary.duplicate_records,
            summary.total_paid_amount,
        )
        LOGGER.info("Generating Excel report at '%s'.", output_path)
        export_report(result, summary, output_path)
        LOGGER.info("Successfully published Excel report at '%s'.", output_path)
    except StructuralInputError as error:
        LOGGER.error("Input processing failed: %s", error)
        return 1
    except ReportExportError as error:
        LOGGER.error("%s", error)
        return 1

    return 0


def main() -> None:
    """Configure operational logging and exit with the processing status."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    raise SystemExit(run())


if __name__ == "__main__":
    main()

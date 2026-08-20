"""Application entry point for the currently implemented input pipeline."""

from __future__ import annotations

import logging
from pathlib import Path

from src.processor import (
    StructuralInputError,
    discover_supported_files,
    load_supported_files,
    process_records,
)


LOGGER = logging.getLogger(__name__)
DEFAULT_INPUT_DIR = Path("data/input")


def run(input_dir: Path = DEFAULT_INPUT_DIR) -> int:
    """Discover, load, and process supported input sources."""
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
    except StructuralInputError as error:
        LOGGER.error("Input processing failed: %s", error)
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

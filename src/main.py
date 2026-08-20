"""Application entry point for the currently implemented input pipeline."""

from __future__ import annotations

import logging
from pathlib import Path

from src.processor import StructuralInputError, discover_csv_files, load_csv_files


LOGGER = logging.getLogger(__name__)
DEFAULT_INPUT_DIR = Path("data/input")


def run(input_dir: Path = DEFAULT_INPUT_DIR) -> int:
    """Discover and structurally load CSV input sources."""
    try:
        csv_files = discover_csv_files(input_dir)
        LOGGER.info("Discovered %d CSV input file(s).", len(csv_files))
        records = load_csv_files(csv_files)
        LOGGER.info(
            "Successfully loaded %d record(s) from %d CSV file(s).",
            len(records),
            len(csv_files),
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

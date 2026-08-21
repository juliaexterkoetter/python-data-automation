"""Generate the report for the repository's fictitious demonstration data."""

import logging
from pathlib import Path

from src.main import run


DEMO_INPUT_DIR = Path("data/demo/input")
DEMO_OUTPUT_PATH = Path("data/demo/output/sales_report.xlsx")


def main() -> None:
    """Run the production pipeline against the isolated demonstration paths."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    raise SystemExit(run(DEMO_INPUT_DIR, DEMO_OUTPUT_PATH))


if __name__ == "__main__":
    main()

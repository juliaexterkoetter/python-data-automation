from __future__ import annotations

import logging
from pathlib import Path

import pytest
from openpyxl import Workbook

from src.main import run


VALID_CSV = (
    "order_id,customer_name,email,order_date,amount,status\n"
    "00123,Julia,julia@example.com,2026-08-20,149.90,paid\n"
)


def test_run_returns_zero_and_logs_success(tmp_path: Path, caplog) -> None:
    (tmp_path / "orders.csv").write_text(VALID_CSV, encoding="utf-8")

    with caplog.at_level(logging.INFO):
        status = run(tmp_path)

    assert status == 0
    assert "Discovered 1 CSV input file(s)." in caplog.text
    assert "Successfully loaded 1 record(s)" in caplog.text
    assert "Processed 1 record(s): 1 valid, 0 invalid, 0 duplicate." in caplog.text


def test_run_returns_nonzero_and_logs_structural_failure(
    tmp_path: Path,
    caplog,
) -> None:
    (tmp_path / "orders.csv").write_text(
        "order_id,status\n00123,paid\n", encoding="utf-8"
    )

    with caplog.at_level(logging.ERROR):
        status = run(tmp_path)

    assert status == 1
    assert "Input processing failed" in caplog.text
    assert "missing required columns" in caplog.text


def test_run_returns_nonzero_when_input_directory_is_missing(
    tmp_path: Path,
    caplog,
) -> None:
    with caplog.at_level(logging.ERROR):
        status = run(tmp_path / "missing")

    assert status == 1
    assert "Input directory does not exist" in caplog.text


def test_run_returns_nonzero_when_no_supported_input_exists(
    tmp_path: Path,
    caplog,
) -> None:
    (tmp_path / "ignored.txt").write_text("ignored", encoding="utf-8")

    with caplog.at_level(logging.ERROR):
        status = run(tmp_path)

    assert status == 1
    assert "No supported input files found" in caplog.text


def test_run_logs_and_returns_nonzero_for_filesystem_errors(
    tmp_path: Path,
    caplog,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    permission_error = PermissionError("access denied")

    def raise_permission_error(_path: Path):
        raise permission_error

    monkeypatch.setattr(Path, "iterdir", raise_permission_error)

    with caplog.at_level(logging.ERROR):
        status = run(tmp_path)

    assert status == 1
    assert "Input processing failed" in caplog.text
    assert "Could not access input directory" in caplog.text


def test_run_does_not_mask_unexpected_programming_errors(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def raise_programming_error(_path: Path):
        raise RuntimeError("unexpected bug")

    monkeypatch.setattr(Path, "iterdir", raise_programming_error)

    with pytest.raises(RuntimeError, match="unexpected bug"):
        run(tmp_path)


def test_run_succeeds_and_logs_record_level_validation_failures(
    tmp_path: Path,
    caplog,
) -> None:
    (tmp_path / "orders.csv").write_text(
        "order_id,customer_name,email,order_date,amount,status\n"
        ",,invalid,01/02/2026,-1.00,unknown\n",
        encoding="utf-8",
    )

    with caplog.at_level(logging.INFO):
        status = run(tmp_path)

    assert status == 0
    assert "Processed 1 record(s): 0 valid, 1 invalid, 0 duplicate." in caplog.text


def test_run_classifies_duplicates_across_multiple_csv_files(
    tmp_path: Path,
    caplog,
) -> None:
    (tmp_path / "a.csv").write_text(VALID_CSV, encoding="utf-8")
    (tmp_path / "b.csv").write_text(VALID_CSV, encoding="utf-8")

    with caplog.at_level(logging.INFO):
        status = run(tmp_path)

    assert status == 0
    assert "Processed 2 record(s): 0 valid, 0 invalid, 2 duplicate." in caplog.text


def test_run_processes_csv_and_xlsx_together(tmp_path: Path, caplog) -> None:
    (tmp_path / "a.csv").write_text(VALID_CSV, encoding="utf-8")
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Orders"
    worksheet.append(
        ["order_id", "customer_name", "email", "order_date", "amount", "status"]
    )
    worksheet.append(["00123", "Ada", None, "2026-08-21", 20, "pending"])
    workbook.save(tmp_path / "b.xlsx")
    workbook.close()

    with caplog.at_level(logging.INFO):
        status = run(tmp_path)

    assert status == 0
    assert "Discovered 1 CSV input file(s)." in caplog.text
    assert "Discovered 1 XLSX input file(s)." in caplog.text
    assert "Processed 2 record(s): 0 valid, 0 invalid, 2 duplicate." in caplog.text


def test_run_returns_nonzero_for_structurally_invalid_xlsx(
    tmp_path: Path,
    caplog,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(["order_id", "status"])
    worksheet.append(["00123", "paid"])
    workbook.save(tmp_path / "orders.xlsx")
    workbook.close()

    with caplog.at_level(logging.ERROR):
        status = run(tmp_path)

    assert status == 1
    assert "XLSX structural error" in caplog.text
    assert "missing required columns" in caplog.text

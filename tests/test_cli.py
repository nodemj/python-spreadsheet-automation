"""End-to-end CLI tests for sheetclean."""

from __future__ import annotations

from pathlib import Path

from sheetclean.cli import main


def test_cli_end_to_end_csv_to_xlsx_and_md(tmp_path: Path) -> None:
    input_csv = tmp_path / "messy.csv"
    input_csv.write_text(
        "Order Date,Customer,Total Amount\n"
        "01/15/2025, Alice ,\"$1,000.00\"\n"
        "2025/01/15,Alice,$1000.00\n"
        "bad-date,Bob,oops\n",
        encoding="utf-8",
    )
    out_xlsx = tmp_path / "out.xlsx"
    out_md = tmp_path / "report.md"
    out_log = tmp_path / "run.log"

    exit_code = main([
        str(input_csv),
        "-o",
        str(out_xlsx),
        "--report-md",
        str(out_md),
        "--log-file",
        str(out_log),
    ])

    assert exit_code == 0
    assert out_xlsx.exists()
    assert out_md.exists()
    assert out_log.exists()
    log_text = out_log.read_text(encoding="utf-8")
    assert "invalid_date" in log_text or "Unable to parse date" in log_text


def test_cli_strict_mode_returns_2_on_validation_errors(tmp_path: Path) -> None:
    input_csv = tmp_path / "bad.csv"
    input_csv.write_text("Order Date,Total Amount\nbad-date,$10.00\n", encoding="utf-8")

    exit_code = main([
        str(input_csv),
        "-o",
        str(tmp_path / "out.csv"),
        "--no-log-file",
        "--strict",
    ])

    assert exit_code == 2

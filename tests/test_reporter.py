"""Tests for Excel, CSV, and Markdown report generation."""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pandas as pd

from sheetclean.cleaner import clean_dataframe
from sheetclean.reporter import format_summary_markdown, write_cleaned_output


def test_write_excel_workbook_with_summary_and_errors(tmp_path: Path) -> None:
    raw = pd.DataFrame(
        {
            "Order Date": ["2025-02-01", "bad-date"],
            "Amount": ["$100.00", "$50.00"],
        }
    )
    result = clean_dataframe(raw)
    out_xlsx = tmp_path / "cleaned.xlsx"

    write_cleaned_output(result, out_xlsx)

    assert out_xlsx.exists()
    wb = openpyxl.load_workbook(out_xlsx)
    assert wb.sheetnames == ["Cleaned_Data", "Summary", "Validation_Errors"]

    data_ws = wb["Cleaned_Data"]
    assert data_ws["A1"].value == "Order Date"
    assert data_ws["B2"].value == 100.0
    assert data_ws["B2"].number_format == "$#,##0.00"

    errors_ws = wb["Validation_Errors"]
    assert errors_ws["C2"].value == "invalid_date"
    assert errors_ws["D2"].value == "bad-date"


def test_format_summary_markdown(tmp_path: Path) -> None:
    raw = pd.DataFrame(
        {
            "Payment Date": ["2025-03-01", "2025-03-05"],
            "Total": ["$10.00", "$20.00"],
        }
    )
    result = clean_dataframe(raw)
    md = format_summary_markdown(
        result,
        input_path=tmp_path / "in.csv",
        output_path=tmp_path / "out.xlsx",
    )

    assert "# SheetClean Summary Report" in md
    assert "| Output Rows | 2 |" in md
    assert "$30.00" in md

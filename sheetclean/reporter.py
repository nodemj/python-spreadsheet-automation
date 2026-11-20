"""Formatted Excel, CSV, and Markdown summary report generation."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from sheetclean.cleaner import CleaningResult

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
SUBHEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
SECTION_FONT = Font(name="Calibri", size=12, bold=True, color="1F4E78")
BOLD_FONT = Font(name="Calibri", size=11, bold=True)
THIN_BORDER = Border(
    bottom=Side(style="thin", color="D9D9D9"),
)
CURRENCY_FORMAT = "$#,##0.00"


def _style_header_row(ws: Worksheet, row_idx: int = 1, fill: PatternFill = HEADER_FILL, font: Font = HEADER_FONT) -> None:
    """Apply header styling to all populated cells in a row."""
    for cell in ws[row_idx]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="left", vertical="center")


def _autofit_columns(ws: Worksheet, min_width: int = 10, max_width: int = 48) -> None:
    """Auto-fit worksheet column widths based on cell string lengths."""
    for col_idx, column_cells in enumerate(ws.columns, start=1):
        max_len = 0
        for cell in column_cells:
            val = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(val))
        adjusted = max(min_width, min(max_len + 4, max_width))
        ws.column_dimensions[get_column_letter(col_idx)].width = adjusted


def _append_df_rows(ws: Worksheet, df: pd.DataFrame) -> None:
    """Write DataFrame headers and rows to a worksheet, converting NaN/NA to None."""
    ws.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        cleaned_row = [None if pd.isna(v) else v for v in row]
        ws.append(cleaned_row)


def _write_cleaned_data_sheet(wb: Workbook, result: CleaningResult) -> None:
    """Populate and format the Cleaned_Data worksheet."""
    ws = wb.active
    ws.title = "Cleaned_Data"
    _append_df_rows(ws, result.df)

    _style_header_row(ws, row_idx=1)
    ws.freeze_panes = "A2"
    if len(result.df.columns) > 0:
        last_col = get_column_letter(len(result.df.columns))
        ws.auto_filter.ref = f"A1:{last_col}{max(len(result.df) + 1, 1)}"

    col_index_map = {col: idx + 1 for idx, col in enumerate(result.df.columns)}
    for cur_col in result.currency_columns:
        col_idx = col_index_map.get(cur_col)
        if col_idx is None:
            continue
        for row_idx in range(2, len(result.df) + 2):
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.value is not None:
                cell.number_format = CURRENCY_FORMAT

    _autofit_columns(ws)


def _write_summary_sheet(wb: Workbook, result: CleaningResult) -> None:
    """Populate and format the Summary worksheet with run metrics and column stats."""
    ws = wb.create_sheet(title="Summary")
    stats = result.stats

    ws.append(["SheetClean Summary Report", ""])
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    ws.append([])

    ws.append(["Metric", "Value"])
    _style_header_row(ws, row_idx=3)

    metric_rows: list[tuple[str, Any]] = [
        ("Input Rows", stats["input_rows"]),
        ("Empty Rows Removed", stats["empty_rows_removed"]),
        ("Duplicates Removed", stats["duplicates_removed"]),
        ("Output Rows", stats["output_rows"]),
        ("Validation Errors", stats["validation_errors"]),
        ("Processing Time (s)", stats["duration_seconds"]),
    ]
    for label, value in metric_rows:
        ws.append([label, value])
        ws.cell(row=ws.max_row, column=1).font = BOLD_FONT
        ws.cell(row=ws.max_row, column=1).border = THIN_BORDER
        ws.cell(row=ws.max_row, column=2).border = THIN_BORDER

    col_summaries: dict[str, dict[str, Any]] = stats.get("column_summaries", {})
    currency_items = [(k, v) for k, v in col_summaries.items() if v["type"] == "currency"]
    date_items = [(k, v) for k, v in col_summaries.items() if v["type"] == "date"]

    if currency_items:
        ws.append([])
        ws.append(["Currency Column Summary"])
        ws.cell(row=ws.max_row, column=1).font = SECTION_FONT
        header_row = ws.max_row + 1
        ws.append(["Column", "Valid Count", "Missing", "Total", "Average", "Min", "Max"])
        _style_header_row(ws, row_idx=header_row, fill=SUBHEADER_FILL, font=BOLD_FONT)

        for col_name, info in currency_items:
            ws.append([
                col_name,
                info["valid_count"],
                info["missing_count"],
                info["sum"],
                info["mean"],
                info["min"],
                info["max"],
            ])
            for c_idx in (4, 5, 6, 7):
                cell = ws.cell(row=ws.max_row, column=c_idx)
                if cell.value is not None:
                    cell.number_format = CURRENCY_FORMAT

    if date_items:
        ws.append([])
        ws.append(["Date Column Summary"])
        ws.cell(row=ws.max_row, column=1).font = SECTION_FONT
        header_row = ws.max_row + 1
        ws.append(["Column", "Valid Count", "Missing", "Earliest Date", "Latest Date"])
        _style_header_row(ws, row_idx=header_row, fill=SUBHEADER_FILL, font=BOLD_FONT)

        for col_name, info in date_items:
            ws.append([
                col_name,
                info["valid_count"],
                info["missing_count"],
                info["min_date"],
                info["max_date"],
            ])

    _autofit_columns(ws)


def _write_errors_sheet(wb: Workbook, result: CleaningResult) -> None:
    """Populate and format the Validation_Errors worksheet."""
    ws = wb.create_sheet(title="Validation_Errors")
    ws.append(["Row", "Column", "Issue Type", "Raw Value", "Message"])
    _style_header_row(ws, row_idx=1)
    ws.freeze_panes = "A2"

    if not result.issues:
        ws.append([None, None, "none", None, "No validation errors detected."])
    else:
        for issue in result.issues:
            ws.append([
                issue.row_number,
                issue.column,
                issue.issue_type,
                issue.raw_value,
                issue.message,
            ])

    _autofit_columns(ws)


def write_cleaned_output(result: CleaningResult, output_path: Path) -> None:
    """Write cleaned data and formatted summary report to CSV or Excel (.xlsx)."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    suffix = output_path.suffix.lower()

    if suffix == ".csv":
        result.df.to_csv(output_path, index=False)
        return

    if suffix in {".xlsx", ".xlsm"}:
        wb = Workbook()
        _write_cleaned_data_sheet(wb, result)
        _write_summary_sheet(wb, result)
        _write_errors_sheet(wb, result)
        wb.save(output_path)
        return

    raise ValueError(
        f"Unsupported output format '{suffix}'. Use .xlsx or .csv."
    )


def format_summary_markdown(
    result: CleaningResult,
    input_path: Path,
    output_path: Path,
) -> str:
    """Render a Markdown summary report of the cleaning run."""
    stats = result.stats
    lines = [
        "# SheetClean Summary Report",
        "",
        f"- **Input File:** `{input_path}`",
        f"- **Output File:** `{output_path}`",
        f"- **Processing Time:** {stats['duration_seconds']}s",
        "",
        "## Row & Quality Metrics",
        "",
        "| Metric | Count |",
        "| --- | ---: |",
        f"| Input Rows | {stats['input_rows']} |",
        f"| Empty Rows Removed | {stats['empty_rows_removed']} |",
        f"| Duplicates Removed | {stats['duplicates_removed']} |",
        f"| Output Rows | {stats['output_rows']} |",
        f"| Validation Errors | {stats['validation_errors']} |",
    ]

    col_summaries: dict[str, dict[str, Any]] = stats.get("column_summaries", {})
    currency_items = [(k, v) for k, v in col_summaries.items() if v["type"] == "currency"]
    date_items = [(k, v) for k, v in col_summaries.items() if v["type"] == "date"]

    if currency_items:
        lines.extend([
            "",
            "## Currency Columns",
            "",
            "| Column | Valid | Missing | Total | Average | Min | Max |",
            "| --- | ---: | ---: | ---: | ---: | ---: | ---: |",
        ])
        for col_name, info in currency_items:
            min_v = f"${info['min']:,.2f}" if info["min"] is not None else "—"
            max_v = f"${info['max']:,.2f}" if info["max"] is not None else "—"
            lines.append(
                f"| {col_name} | {info['valid_count']} | {info['missing_count']} "
                f"| ${info['sum']:,.2f} | ${info['mean']:,.2f} | {min_v} | {max_v} |"
            )

    if date_items:
        lines.extend([
            "",
            "## Date Columns",
            "",
            "| Column | Valid | Missing | Earliest | Latest |",
            "| --- | ---: | ---: | --- | --- |",
        ])
        for col_name, info in date_items:
            lines.append(
                f"| {col_name} | {info['valid_count']} | {info['missing_count']} "
                f"| {info['min_date'] or '—'} | {info['max_date'] or '—'} |"
            )

    if result.issues:
        lines.extend([
            "",
            "## Validation Errors",
            "",
            "| Row | Column | Issue Type | Raw Value | Message |",
            "| ---: | --- | --- | --- | --- |",
        ])
        for issue in result.issues:
            lines.append(
                f"| {issue.row_number} | {issue.column} | `{issue.issue_type}` "
                f"| `{issue.raw_value}` | {issue.message} |"
            )

    lines.append("")
    return "\n".join(lines)

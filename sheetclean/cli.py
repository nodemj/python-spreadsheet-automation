"""Single-command CLI interface for sheetclean."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from sheetclean.cleaner import CleaningConfig, clean_file
from sheetclean.logging_cfg import setup_logging
from sheetclean.reporter import format_summary_markdown, write_cleaned_output


def _parse_csv_list(raw: str | None) -> list[str] | None:
    """Parse a comma-separated CLI option into a trimmed list of strings."""
    if raw is None:
        return None
    items = [item.strip() for item in raw.split(",") if item.strip()]
    return items or None


def build_parser() -> argparse.ArgumentParser:
    """Build the CLI argument parser."""
    parser = argparse.ArgumentParser(
        prog="sheetclean",
        description=(
            "Clean messy CSV/Excel files, remove duplicates, standardize dates and "
            "currencies, log validation errors, and generate a formatted summary report."
        ),
    )
    parser.add_argument(
        "input_file",
        type=Path,
        help="Path to the input .csv, .xlsx, .xlsm, or .xls file.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Output path (.xlsx or .csv). Defaults to <input_stem>_cleaned.xlsx.",
    )
    parser.add_argument(
        "--sheet",
        default="0",
        help="Excel sheet name or 0-based index (default: 0).",
    )
    parser.add_argument(
        "--date-cols",
        default=None,
        help="Comma-separated date column names (auto-detected if omitted).",
    )
    parser.add_argument(
        "--currency-cols",
        default=None,
        help="Comma-separated currency column names (auto-detected if omitted).",
    )
    parser.add_argument(
        "--dedup-cols",
        default=None,
        help="Comma-separated columns used for duplicate detection (default: all).",
    )
    parser.add_argument(
        "--date-format",
        default="%Y-%m-%d",
        help="Target strftime format for dates (default: %%Y-%%m-%%d).",
    )
    parser.add_argument(
        "--snake-case",
        action="store_true",
        help="Normalize column headers to snake_case.",
    )
    parser.add_argument(
        "--report-md",
        type=Path,
        default=None,
        help="Optional path to save a Markdown summary report.",
    )
    parser.add_argument(
        "--log-file",
        type=Path,
        default=None,
        help="Optional path for the error/audit log file (default: <output_stem>.log).",
    )
    parser.add_argument(
        "--no-log-file",
        action="store_true",
        help="Disable writing a log file to disk.",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit with code 2 if any validation errors are detected.",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Enable verbose debug logging.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    """CLI entry point."""
    parser = build_parser()
    args = parser.parse_args(argv)

    input_path: Path = args.input_file
    output_path: Path = args.output or input_path.with_name(f"{input_path.stem}_cleaned.xlsx")

    log_file: Path | None = None
    if not args.no_log_file:
        log_file = args.log_file or output_path.with_suffix(".log")

    logger = setup_logging(log_file=log_file, verbose=args.verbose)

    sheet_name: str | int = int(args.sheet) if args.sheet.isdigit() else args.sheet
    config = CleaningConfig(
        date_columns=_parse_csv_list(args.date_cols),
        currency_columns=_parse_csv_list(args.currency_cols),
        dedup_subset=_parse_csv_list(args.dedup_cols),
        date_format=args.date_format,
        snake_case_headers=args.snake_case,
    )

    try:
        result = clean_file(input_path, config=config, sheet_name=sheet_name)
        write_cleaned_output(result, output_path)
    except Exception as exc:
        logger.error("Failed to process '%s': %s", input_path, exc)
        return 1

    if args.report_md is not None:
        args.report_md.parent.mkdir(parents=True, exist_ok=True)
        args.report_md.write_text(
            format_summary_markdown(result, input_path, output_path),
            encoding="utf-8",
        )
        logger.info("Saved Markdown summary report: %s", args.report_md)

    stats = result.stats
    print(
        f"\nSheetClean complete: {output_path}\n"
        f"  Input rows:         {stats['input_rows']}\n"
        f"  Empty rows removed: {stats['empty_rows_removed']}\n"
        f"  Duplicates removed: {stats['duplicates_removed']}\n"
        f"  Output rows:        {stats['output_rows']}\n"
        f"  Validation errors:  {stats['validation_errors']}\n"
        f"  Elapsed:            {stats['duration_seconds']}s"
    )
    if log_file is not None:
        print(f"  Log file:           {log_file}")

    if args.strict and stats["validation_errors"] > 0:
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())

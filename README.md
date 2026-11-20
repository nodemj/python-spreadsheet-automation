# SheetClean

Automated Python & Pandas CLI tool that cleans messy CSV and Excel files, removes duplicates, standardizes date and currency formats, logs validation errors, and generates a formatted summary report.

## Features

- **Multi-format input/output:** Reads `.csv`, `.xlsx`, `.xlsm`, `.xls` and writes formatted `.xlsx` workbooks or `.csv`.
- **Smart cleanup:** Normalizes headers, trims whitespace, collapses internal spaces, converts null-like placeholders (`N/A`, `null`, `-`) to empty values, and drops empty rows.
- **Deduplication:** Removes exact or subset-key duplicates after normalization.
- **Date & currency standardization:** Auto-detects (or accepts explicit) date and currency columns, converts mixed date formats to ISO `YYYY-MM-DD` (configurable), and parses symbols/accounting negatives like `($1,250.50)` into clean numeric values.
- **Validation & error logging:** Logs unparseable dates and invalid currency values to both a `.log` file and a `Validation_Errors` tab in the output workbook.
- **Formatted summary report:** Generates styled `Cleaned_Data`, `Summary`, and `Validation_Errors` Excel sheets, plus an optional Markdown report (`--report-md`).

## Quick Start

```bash
# Install in editable mode
pip install -e .

# Clean a messy file and generate a formatted Excel report
sheetclean examples/messy_sales.csv -o output/sales_cleaned.xlsx --report-md output/summary.md

# Or run as a module
python3 -m sheetclean examples/messy_sales.csv -o output/sales_cleaned.xlsx
```

## CLI Options

- `-o, --output`: Output path (`.xlsx` or `.csv`). Defaults to `<input_stem>_cleaned.xlsx`.
- `--sheet`: Excel sheet name or 0-based index (default: `0`).
- `--date-cols`: Comma-separated date column names (auto-detected if omitted).
- `--currency-cols`: Comma-separated currency column names (auto-detected if omitted).
- `--dedup-cols`: Comma-separated columns used for duplicate detection (default: all).
- `--date-format`: Target `strftime` format for dates (default: `%Y-%m-%d`).
- `--snake-case`: Normalize column headers to `snake_case`.
- `--report-md`: Optional path to save a Markdown summary report.
- `--log-file`: Optional path for the audit log file (default: `<output_stem>.log`).
- `--strict`: Exit with code `2` if any validation errors are detected.
- `-v, --verbose`: Enable debug logging.

## Running Tests

```bash
python3 -m pytest -v
```

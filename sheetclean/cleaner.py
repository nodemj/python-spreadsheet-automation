"""Core Pandas cleaning, standardization, deduplication, and validation logic."""

from __future__ import annotations

import logging
import re
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd

logger = logging.getLogger("sheetclean")

DATE_HEADER_KEYWORDS = ("date", "time", "created", "updated", "dob", "timestamp")
CURRENCY_HEADER_KEYWORDS = (
    "amount",
    "price",
    "total",
    "cost",
    "revenue",
    "fee",
    "balance",
    "salary",
    "spend",
    "payment",
    "subtotal",
    "tax",
)
NULL_PLACEHOLDERS = {"", "na", "n/a", "null", "none", "nan", "-", "--"}
CURRENCY_PREFIX_RE = re.compile(
    r"^(?:USD|EUR|GBP|CAD|AUD|JPY|INR|CHF)\s*", re.IGNORECASE
)
CURRENCY_SYMBOL_RE = re.compile(r"[$€£¥₹,\s]")


@dataclass
class CleaningConfig:
    """Configuration options for the spreadsheet cleaning pipeline."""

    date_columns: list[str] | None = None
    currency_columns: list[str] | None = None
    dedup_subset: list[str] | None = None
    date_format: str = "%Y-%m-%d"
    currency_decimals: int = 2
    drop_empty_rows: bool = True
    trim_whitespace: bool = True
    snake_case_headers: bool = False


@dataclass
class ValidationIssue:
    """Represents a single data validation or parsing error."""

    row_number: int
    column: str
    raw_value: str
    issue_type: str
    message: str


@dataclass
class CleaningResult:
    """Output of the cleaning pipeline including cleaned data, issues, and metrics."""

    df: pd.DataFrame
    issues: list[ValidationIssue] = field(default_factory=list)
    stats: dict[str, Any] = field(default_factory=dict)
    date_columns: list[str] = field(default_factory=list)
    currency_columns: list[str] = field(default_factory=list)


def load_dataframe(input_path: Path, sheet_name: str | int = 0) -> pd.DataFrame:
    """Load a CSV or Excel file into a raw DataFrame without premature type coercion."""
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    suffix = input_path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(input_path, dtype=object, keep_default_na=False)
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(
            input_path, sheet_name=sheet_name, dtype=object, keep_default_na=False
        )
    raise ValueError(
        f"Unsupported file format '{suffix}'. Supported formats: .csv, .xlsx, .xlsm, .xls"
    )


def _normalize_headers(columns: pd.Index, snake_case: bool) -> list[str]:
    """Normalize column names and ensure uniqueness."""
    normalized: list[str] = []
    seen: dict[str, int] = {}

    for col in columns:
        name = re.sub(r"\s+", " ", str(col).strip())
        if snake_case:
            name = re.sub(r"[^a-zA-Z0-9]+", "_", name).strip("_").lower()
        if not name:
            name = "unnamed"

        if name in seen:
            seen[name] += 1
            unique_name = f"{name}_{seen[name]}"
        else:
            seen[name] = 1
            unique_name = name
        normalized.append(unique_name)

    return normalized


def _detect_columns(columns: list[str], keywords: tuple[str, ...]) -> list[str]:
    """Auto-detect columns whose lowercase name contains any target keyword."""
    matched: list[str] = []
    for col in columns:
        lower_col = col.lower()
        if any(kw in lower_col for kw in keywords):
            matched.append(col)
    return matched


def _resolve_columns(
    requested: list[str] | None,
    available: list[str],
    keywords: tuple[str, ...],
    snake_case: bool,
) -> list[str]:
    """Resolve user-specified columns or fall back to keyword auto-detection."""
    if requested is None:
        return _detect_columns(available, keywords)

    resolved: list[str] = []
    for raw_col in requested:
        norm = re.sub(r"\s+", " ", raw_col.strip())
        if snake_case:
            norm = re.sub(r"[^a-zA-Z0-9]+", "_", norm).strip("_").lower()
        if norm in available:
            resolved.append(norm)
        elif raw_col.strip() in available:
            resolved.append(raw_col.strip())
        else:
            logger.warning("Requested column '%s' not found in dataset; skipping.", raw_col)
    return resolved


def _clean_cell_text(val: Any) -> Any:
    """Trim whitespace and convert null-like string placeholders to None."""
    if val is None or pd.isna(val):
        return None
    if isinstance(val, str):
        cleaned = re.sub(r"\s+", " ", val.strip())
        if cleaned.lower() in NULL_PLACEHOLDERS:
            return None
        return cleaned
    return val


def _parse_currency_value(val: Any, decimals: int) -> tuple[float | None, bool]:
    """Parse a messy currency value into a rounded float.

    Returns:
        Tuple of (parsed_value, is_valid). Empty/None values return (None, True).
    """
    if val is None or pd.isna(val):
        return None, True
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return round(float(val), decimals), True

    text = str(val).strip()
    if text.lower() in NULL_PLACEHOLDERS:
        return None, True

    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1].strip()

    text = CURRENCY_PREFIX_RE.sub("", text)
    text = CURRENCY_SYMBOL_RE.sub("", text)

    if text.startswith("-"):
        negative = not negative
        text = text[1:]
    elif text.startswith("+"):
        text = text[1:]

    if not text:
        return None, False

    try:
        num = float(text)
    except ValueError:
        return None, False

    if negative:
        num = -abs(num)
    return round(num, decimals), True


def _standardize_dates(
    df: pd.DataFrame,
    date_cols: list[str],
    date_format: str,
    issues: list[ValidationIssue],
) -> None:
    """Standardize date columns in-place and record unparseable values."""
    for col in date_cols:
        raw_series = df[col]
        parsed = pd.to_datetime(raw_series, errors="coerce", format="mixed")

        formatted_values: list[str | None] = []
        for idx, (raw_val, ts) in enumerate(zip(raw_series, parsed)):
            if raw_val is None or pd.isna(raw_val):
                formatted_values.append(None)
            elif pd.isna(ts):
                msg = f"Unable to parse date '{raw_val}' in column '{col}'."
                logger.warning("Row %d: %s", idx + 1, msg)
                issues.append(
                    ValidationIssue(
                        row_number=idx + 1,
                        column=col,
                        raw_value=str(raw_val),
                        issue_type="invalid_date",
                        message=msg,
                    )
                )
                formatted_values.append(None)
            else:
                formatted_values.append(ts.strftime(date_format))

        df[col] = formatted_values


def _standardize_currencies(
    df: pd.DataFrame,
    currency_cols: list[str],
    decimals: int,
    issues: list[ValidationIssue],
) -> None:
    """Standardize currency columns in-place to numeric floats and log errors."""
    for col in currency_cols:
        parsed_values: list[float | None] = []
        for idx, raw_val in enumerate(df[col]):
            num, is_valid = _parse_currency_value(raw_val, decimals)
            if not is_valid:
                msg = f"Invalid currency value '{raw_val}' in column '{col}'."
                logger.warning("Row %d: %s", idx + 1, msg)
                issues.append(
                    ValidationIssue(
                        row_number=idx + 1,
                        column=col,
                        raw_value=str(raw_val),
                        issue_type="invalid_currency",
                        message=msg,
                    )
                )
            parsed_values.append(num)

        df[col] = pd.Series(parsed_values, index=df.index, dtype="float64")


def _build_column_summaries(
    df: pd.DataFrame,
    date_cols: list[str],
    currency_cols: list[str],
) -> dict[str, dict[str, Any]]:
    """Compute summary statistics for standardized currency and date columns."""
    summaries: dict[str, dict[str, Any]] = {}

    for col in currency_cols:
        valid = df[col].dropna()
        summaries[col] = {
            "type": "currency",
            "valid_count": int(valid.count()),
            "missing_count": int(df[col].isna().sum()),
            "sum": round(float(valid.sum()), 2) if not valid.empty else 0.0,
            "mean": round(float(valid.mean()), 2) if not valid.empty else 0.0,
            "min": round(float(valid.min()), 2) if not valid.empty else None,
            "max": round(float(valid.max()), 2) if not valid.empty else None,
        }

    for col in date_cols:
        valid = df[col].dropna()
        summaries[col] = {
            "type": "date",
            "valid_count": int(valid.count()),
            "missing_count": int(df[col].isna().sum()),
            "min_date": str(valid.min()) if not valid.empty else None,
            "max_date": str(valid.max()) if not valid.empty else None,
        }

    return summaries


def clean_dataframe(
    raw_df: pd.DataFrame,
    config: CleaningConfig | None = None,
) -> CleaningResult:
    """Clean, validate, deduplicate, and standardize a DataFrame.

    Args:
        raw_df: Input DataFrame to clean.
        config: Optional ``CleaningConfig`` controlling detection and formatting.

    Returns:
        ``CleaningResult`` containing the cleaned DataFrame, validation issues, and stats.
    """
    start_time = time.perf_counter()
    cfg = config or CleaningConfig()
    df = raw_df.copy()
    input_rows = len(df)
    issues: list[ValidationIssue] = []

    df.columns = _normalize_headers(df.columns, snake_case=cfg.snake_case_headers)
    available_cols = list(df.columns)

    if cfg.trim_whitespace:
        for col in available_cols:
            df[col] = df[col].map(_clean_cell_text)

    empty_rows_removed = 0
    if cfg.drop_empty_rows:
        before = len(df)
        df = df.dropna(how="all").reset_index(drop=True)
        empty_rows_removed = before - len(df)

    date_cols = _resolve_columns(
        cfg.date_columns, available_cols, DATE_HEADER_KEYWORDS, cfg.snake_case_headers
    )
    currency_cols = _resolve_columns(
        cfg.currency_columns,
        available_cols,
        CURRENCY_HEADER_KEYWORDS,
        cfg.snake_case_headers,
    )
    # Prevent a column from being treated as both date and currency
    currency_cols = [c for c in currency_cols if c not in date_cols]

    _standardize_dates(df, date_cols, cfg.date_format, issues)
    _standardize_currencies(df, currency_cols, cfg.currency_decimals, issues)

    dedup_cols = None
    if cfg.dedup_subset:
        dedup_cols = _resolve_columns(
            cfg.dedup_subset, available_cols, (), cfg.snake_case_headers
        ) or None

    before_dedup = len(df)
    df = df.drop_duplicates(subset=dedup_cols, keep="first").reset_index(drop=True)
    duplicates_removed = before_dedup - len(df)

    duration = round(time.perf_counter() - start_time, 4)
    col_summaries = _build_column_summaries(df, date_cols, currency_cols)

    stats: dict[str, Any] = {
        "input_rows": input_rows,
        "empty_rows_removed": empty_rows_removed,
        "duplicates_removed": duplicates_removed,
        "output_rows": len(df),
        "validation_errors": len(issues),
        "date_columns_cleaned": date_cols,
        "currency_columns_cleaned": currency_cols,
        "duration_seconds": duration,
        "column_summaries": col_summaries,
    }

    logger.info(
        "Cleaned %d rows -> %d rows (removed %d empty, %d duplicates, %d validation errors) in %.3fs",
        input_rows,
        len(df),
        empty_rows_removed,
        duplicates_removed,
        len(issues),
        duration,
    )

    return CleaningResult(
        df=df,
        issues=issues,
        stats=stats,
        date_columns=date_cols,
        currency_columns=currency_cols,
    )


def clean_file(
    input_path: Path,
    config: CleaningConfig | None = None,
    sheet_name: str | int = 0,
) -> CleaningResult:
    """Load and clean a CSV or Excel file."""
    logger.info("Loading input file: %s", input_path)
    raw_df = load_dataframe(input_path, sheet_name=sheet_name)
    return clean_dataframe(raw_df, config=config)

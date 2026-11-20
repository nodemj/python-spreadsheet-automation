"""Unit tests for sheetclean.cleaner."""

from __future__ import annotations

import pandas as pd

from sheetclean.cleaner import CleaningConfig, clean_dataframe


def test_clean_dataframe_dedup_dates_currency_and_validation() -> None:
    raw = pd.DataFrame(
        {
            "  Order ID ": ["1001", "1002", "1001", "", "1003", "1004"],
            " Order Date ": [
                "01/15/2025",
                "2025-01-16",
                "2025/01/15",
                "   ",
                "not-a-date",
                "17-Jan-2025",
            ],
            "Customer Name": [
                " Alice   Smith ",
                "Bob Jones",
                "Alice Smith",
                None,
                "Carol Lee",
                "Dave Kim",
            ],
            "Total Amount": [
                "$1,250.50",
                "($200.00)",
                "$1250.50",
                "",
                "EUR 450.25",
                "invalid_amt",
            ],
        }
    )

    result = clean_dataframe(raw)

    assert result.stats["input_rows"] == 6
    assert result.stats["empty_rows_removed"] == 1
    assert result.stats["duplicates_removed"] == 1
    assert result.stats["output_rows"] == 4
    assert result.stats["validation_errors"] == 2

    assert list(result.df.columns) == [
        "Order ID",
        "Order Date",
        "Customer Name",
        "Total Amount",
    ]
    assert result.df.loc[0, "Order Date"] == "2025-01-15"
    assert result.df.loc[0, "Customer Name"] == "Alice Smith"
    assert result.df.loc[0, "Total Amount"] == 1250.50
    assert result.df.loc[1, "Total Amount"] == -200.00
    assert pd.isna(result.df.loc[2, "Order Date"])
    assert result.df.loc[2, "Total Amount"] == 450.25
    assert pd.isna(result.df.loc[3, "Total Amount"])

    issue_types = {issue.issue_type for issue in result.issues}
    assert issue_types == {"invalid_date", "invalid_currency"}


def test_snake_case_headers_and_custom_dedup_subset() -> None:
    raw = pd.DataFrame(
        {
            "User Email": ["a@example.com", "a@example.com", "b@example.com"],
            "Signup Date": ["2025-01-01", "2025-01-02", "2025-01-03"],
            "Fee Amount": ["$10", "$20", "$30"],
        }
    )
    cfg = CleaningConfig(
        snake_case_headers=True,
        dedup_subset=["user_email"],
        date_format="%d/%m/%Y",
    )

    result = clean_dataframe(raw, config=cfg)

    assert list(result.df.columns) == ["user_email", "signup_date", "fee_amount"]
    assert result.stats["duplicates_removed"] == 1
    assert result.stats["output_rows"] == 2
    assert result.df.loc[0, "signup_date"] == "01/01/2025"
    assert result.df.loc[0, "fee_amount"] == 10.0

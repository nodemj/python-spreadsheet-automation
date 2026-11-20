# SheetClean Summary Report

- **Input File:** `examples/messy_sales.csv`
- **Output File:** `examples/output/messy_sales_cleaned.xlsx`
- **Processing Time:** 0.0098s

## Row & Quality Metrics

| Metric | Count |
| --- | ---: |
| Input Rows | 7 |
| Empty Rows Removed | 1 |
| Duplicates Removed | 1 |
| Output Rows | 5 |
| Validation Errors | 2 |

## Currency Columns

| Column | Valid | Missing | Total | Average | Min | Max |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Total Amount | 3 | 2 | $1,500.75 | $500.25 | $-200.00 | $1,250.50 |

## Date Columns

| Column | Valid | Missing | Earliest | Latest |
| --- | ---: | ---: | --- | --- |
| Order Date | 4 | 1 | 2025-01-15 | 2025-01-20 |

## Validation Errors

| Row | Column | Issue Type | Raw Value | Message |
| ---: | --- | --- | --- | --- |
| 5 | Order Date | `invalid_date` | `not-a-date` | Unable to parse date 'not-a-date' in column 'Order Date'. |
| 6 | Total Amount | `invalid_currency` | `pending` | Invalid currency value 'pending' in column 'Total Amount'. |

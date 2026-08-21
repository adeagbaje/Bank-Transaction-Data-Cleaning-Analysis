# Bank Transaction Data Cleaning & Analysis

End-to-end data cleaning and analysis project built in Excel: transforms a messy, nested transaction dataset into a structured, analysis-ready table with a documented QA log and summary dashboard.

## Overview

This project takes a raw dataset of 127 customer transaction records — where transaction details were stored as nested JSON-like objects inside single cells, names were uncapitalized, dates were ISO-8601 strings with microsecond precision, and missing dates of birth were encoded as the placeholder string `'n/a'` — and transforms it into a clean, structured dataset ready for analysis.

## What I did

- **Parsed and flattened nested transaction objects** using Power Query, extracting transaction ID, date, time, hour, amount, mode, channel, and bank into individual columns
- **Standardized text fields** with Excel formulas — trimmed whitespace and applied consistent Title Case to first and last names
- **Converted placeholder nulls** (`'n/a'`) to true missing values rather than silently imputing them, preserving data integrity
- **Derived new fields**, including customer age and age band, calculated from date of birth relative to transaction date
- **Identified and documented data quality issues** in a full audit log — including the key finding that the dataset only contained 4 distinct transactions duplicated across 127 customer rows, and that the `tx_mode` field had zero variance (every row was `'Debit'`)
- **Built a summary dashboard** using PivotTables and PivotCharts, visualizing transaction value by bank, the web vs. mobile split, customer age band distribution, and transaction timing by hour of day

## Tools

Excel — formulas (`TRIM`, `PROPER`, date/time parsing) and Power Query for parsing, reshaping, and cleaning; PivotTables/PivotCharts for the dashboard.

## Files

| Sheet | Description |
|---|---|
| `Raw Data` | Original, unprocessed records as received |
| `Cleaned Data` | Final structured, analysis-ready dataset |
| `Data Quality Log` | Issue-by-issue documentation of what was found and how it was resolved |
| `Summary` | Key completeness and transaction-value statistics |
| `Dashboard` | Charts: transaction value by bank, web vs. mobile split, records by age band, transactions by hour of day |

## Key findings

- 127 customer records mapped to only **4 distinct transactions** — a ~97% duplication rate at the transaction level, even though every customer row was unique
- Only **35% of records** had a usable date of birth; age-based figures reflect that subset only, and missing DOBs were left as `NULL` rather than imputed
- The `tx_mode` field carried no analytical value, since every transaction was a debit

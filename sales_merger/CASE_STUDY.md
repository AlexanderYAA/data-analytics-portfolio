# Case study: Sales Report Automation

> Portfolio project built on synthetic data. No real business or customer data is used.

## Upwork portfolio fields

**Project title** (≤ 70 characters)
Python Sales Report Automation: 36 Messy Files to One Excel Report

**Your role**
Data analyst & Python developer

**Project description** (≤ 600 characters)
Portfolio scenario: a small retail shop gets monthly sales exports from three order channels (till, phone log, website), each in a different format. I built a Python script that merges a year of these files in one run: it recognises different column names, date and number formats, removes duplicates, rejects invalid rows with a clear reason, and writes a formatted Excel report with live formulas plus a clean CSV for Power BI. Totals were verified against the source data to the cent. Synthetic data.

**Skills**
Python · pandas · openpyxl · Microsoft Excel · Data Cleaning · Automation · Power BI

**Files to attach**
`examples/sample_report.pdf` (first page as the cover image), `examples/sample_report.xlsx`, a link to the code repository.

---

## Full case study

### Problem
A small retail shop takes orders in-store, by phone and through its website. Every month each channel produces its own file, and none of them look alike:

- **Till export:** comma CSV, `2025-01-31` dates.
- **Phone log, kept by hand in Excel:** title lines above the table, `31.01.2025` dates, a TOTAL line at the bottom, its own column names (`Order #`, `Qty`, `Price, $`, `Client`).
- **Website export:** semicolon CSV, `01/31/2025` dates, decimal commas, its own spelling of categories and customer types. Each export repeats the last orders of the previous one, and test orders without a price slip in.

A year means 36 files. Merging them by hand is slow, and the typical errors (double-counted orders, a misread date, last year's orders in this year's total) go unnoticed.

### Solution
A Python script (pandas + openpyxl) that runs with one double-click:

1. Reads every CSV/XLSX file in a folder and detects the encoding, delimiter and header row by itself.
2. Maps different column names, category spellings and customer types to one standard, using a simple `config.json` that a non-programmer can edit.
3. Reads each file's date format, decimal commas and currency signs.
4. Validates every row. Duplicates, rows without a price, dates outside the reporting year and TOTAL lines are excluded, and each one is logged with the file, row number and reason.
5. Writes one Excel report with **live formulas**: revenue, cost of goods and gross profit are calculated in the workbook, and every summary uses SUMIFS/COUNTIFS over the data. If the client changes a cost % on the Assumptions sheet, the whole report updates.
6. Also saves a clean `Orders.csv` that feeds a Power BI dashboard directly.

Report sheets: Summary (KPIs, processed files, monthly chart), By month, By channel, By category, Data, Issues, Assumptions.

### Result
- 36 files with 952 order lines → 914 clean orders in under 2 seconds.
- 50 problems found and explained on the Issues sheet: 33 duplicated website orders, 12 TOTAL lines, 3 orders from 2024/2026, 2 test orders without a price.
- **Verified:** the merged data matches the source dataset row by row. Revenue $58,482.14, gross profit $31,712.54 (54.2% margin), 914 orders, average order $63.98, 45.6% returning customers, exactly the control figures of the Excel and Power BI dashboards built on the same data. The report's formulas were recalculated in Microsoft Excel with no errors.

### What this means for a client
- A monthly report becomes a double-click instead of hours of copy-pasting.
- Nothing is dropped silently: every excluded row has a reason the client can check.
- A new file format, category or cost % means editing a config file, not paying for new code.
- The clean output can feed Power BI or any other dashboard.

### Tools
Python 3, pandas, openpyxl, Microsoft Excel, Power BI.

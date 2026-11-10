# Sales Report Automation

A Python script that turns a year of messy monthly sales exports into one clean Excel report with live formulas, plus a ready-to-use table for Power BI, in a single run.

> Portfolio project. All data is synthetic. It is the same dataset as my Excel and Power BI *Retail Sales Dashboard*: this script is the step that prepares its data. See [CASE_STUDY.md](CASE_STUDY.md) for the short case-study version.

## The problem

A small retail shop takes orders through three channels, and each channel keeps its own monthly file in its own format:

| Channel | File | Quirks |
|---|---|---|
| In-store | `instore_2025-MM.csv`: till export | Comma CSV, `2025-01-31` dates. The January and December exports also picked up orders from 2024 / 2026 |
| Phone | `phone_2025-MM.xlsx`: order log kept by hand | Title lines above the table, `31.01.2025` dates, headers like `Price, $` and `Client: Repeat`, a `TOTAL` line at the bottom |
| Online | `online_2025-MM.csv`: website export | Semicolon CSV, `01/31/2025` dates, decimal comma `46,86`, its own spelling of categories and customer types, each export repeats the last orders of the previous one, an occasional test order without a price |

That is 36 files a year. Merging them by hand takes hours, and duplicates, misread dates or orders from the wrong year slip in silently.

## What the script does

1. Reads every `.csv` / `.xlsx` file in `input/`, detecting the encoding, delimiter and header row automatically.
2. Matches column names to a standard set (`Order #`, `OrderID` and `Order ID` all become one column).
3. Normalises dates (it picks the format that fits the whole file and warns if a file is ambiguous), numbers, category names and customer types.
4. Takes the channel from the file name (`phone_2025-03.xlsx` → *Phone*).
5. Logs every problem instead of dropping it silently:
   - **Error**: an empty or unreadable price, quantity or date, or a date outside the reporting period. The row is excluded.
   - **Warning**: a duplicate order (removed), an unknown category, customer type or channel, an ambiguous date format, or a category without a COGS %.
   - **Info**: `TOTAL` and blank lines (skipped), returns with negative quantity (kept).
6. Writes two files to `output/`:
   - `sales_report_2025-01_to_2025-12.xlsx`, a formatted report with **live formulas**:

     | Sheet | Contents |
     |---|---|
     | Summary | KPIs (revenue, gross profit, margin, orders, average order, returning customers, issues), processed files, a monthly revenue chart |
     | By month | Revenue per month per channel, gross profit, margin, orders |
     | By channel | Revenue, gross profit, margin, orders, units, average order, returning customers, share |
     | By category | The same, by product category |
     | Data | All clean orders. Revenue, COGS and gross profit are formulas; each row keeps its source file and row number |
     | Issues | Every problem found, with the file, row and what was done |
     | Assumptions | COGS % per category and the reporting period |

     Every summary is SUMIFS / COUNTIFS over the Data sheet, and COGS is looked up from Assumptions. Change a cost % or fix an order in Excel and the whole report recalculates.
   - `Orders.csv`, the clean table in the layout the Power BI dashboard loads.

Sample run:

```
Files: 36   orders: 914   revenue: $58,482.14   gross profit: $31,712.54   period: 2025-01-01 to 2025-12-31
Issues: Warning 33, Info 12, Error 5
```

## Verification

- The merged `Orders.csv` is identical, row by row, to the source data the 36 files were made from.
- The report was recalculated in Microsoft Excel with no formula errors. Every figure matches the Excel and Power BI dashboards built on that data: revenue $58,482, gross profit $31,713, margin 54.2%, 914 orders, average order $63.98, 45.6% returning customers, plus all monthly, channel and category totals.
- Live-formula test: changing the Bouquets COGS from 45% to 50% on the Assumptions sheet lowered gross profit by exactly $1,479.61 (5% of Bouquets revenue).

**What "Issues logged" means:** the number of problems found in the source files, each listed on the Issues sheet. In the sample: 33 orders repeated across website exports, 12 TOTAL lines in the phone logs, 3 orders from 2024 / 2026, and 2 test orders without a price. The totals are correct because of this cleaning, and the client can see exactly what was wrong in their files.

## Examples

`examples/` holds the output of a sample run, recalculated in Excel so the values show in any viewer:

- `sample_report.xlsx`: the report.
- `sample_report.pdf`: the main sheets as a PDF.
- `Orders.csv`: the clean table for Power BI.

## Adapting it to a new client: `config.json`, no code changes

- `columns`: the header names to recognise for each standard column.
- `categories`, `customers`: spellings that mean the same thing.
- `cogs_pct`: cost of goods as a share of revenue, per category (copied to the Assumptions sheet).
- `channels`: file name prefix → channel name.
- `period`: the reporting period. Orders outside it are excluded and logged.
- `date_formats`: the date formats to try, in order of priority.
- `report_note`: an optional line under the report title (used here for the synthetic-data note).

## How to run

**One click (Windows):** put files into `input/`, double-click `run_report.bat`. The first run sets up the Python packages, then the report folder opens.

**Command line:**

```bash
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python merge_sales.py --input path\to\files --output path\to\reports
```

To rebuild the 36 sample files from `source/Orders.csv`, run `python make_sample_data.py`.

## Stack

Python 3.10+, pandas, openpyxl.

## Possible next steps for a real client

- Run on a schedule (Windows Task Scheduler) and email the report.
- Read files straight from Google Drive / OneDrive or the shop's website API.
- Point Power BI at `output/Orders.csv` so the dashboard refreshes after every run.

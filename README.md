# Data Analytics & Automation Portfolio

Messy spreadsheets in, clear dashboards and automated reports out. These projects show the full path for a small retail business: cleaning and merging raw sales files, analysing them, and presenting the results.

> All projects use the same **synthetic** dataset (914 orders of a small retail shop in 2025, three sales channels, four product categories). No real business or customer data is used.

| Project | What it shows | Tools |
|---|---|---|
| [Excel Sales Dashboard](excel_dashboard/) | A one-page dashboard built with formulas only: one channel selector updates every KPI, table, chart and text insight | Excel (SUMIFS, INDEX/MATCH, charts) |
| [Sales Report Automation](sales_merger/) | A script that merges 36 monthly files in different formats into one Excel report with live formulas, and logs every rejected row with a reason | Python, pandas, openpyxl, Excel |
| [Interactive Sales Dashboard](html_dashboard/) | A single-file dashboard with filters, auto-generated insights and CSV export. [Live demo](https://alexanderyaa.github.io/data-analytics-portfolio/html_dashboard/) | HTML, JavaScript, Chart.js |
| [Retail Sales Analysis](python_report/) | A Jupyter notebook that answers six business questions and ends with recommendations, plus a PDF report for the owner | Python, pandas, matplotlib, Jupyter |

All projects were cross-checked against each other and against the same control figures: revenue $58,482.14, gross profit $31,712.54, margin 54.2%, 914 orders, average order $63.98.

## How the pieces fit together

1. **Sales Report Automation** turns raw exports into a clean `Orders.csv` and an Excel report.
2. That clean table feeds the **Interactive Dashboard** (and a Power BI version of the same report).
3. The **Sales Analysis** notebook digs into the same data to explain *why* the numbers look the way they do and what to do next.

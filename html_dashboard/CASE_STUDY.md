# Case study: Interactive HTML Sales Dashboard

> Portfolio project built on synthetic data. No real business or customer data is used.

## Upwork portfolio fields

**Project title** (≤ 70 characters)
Interactive Sales Dashboard in One HTML File (No Power BI Needed)

**Your role**
Data analyst & dashboard developer

**Project description** (≤ 600 characters)
Portfolio scenario: a small retail shop needs to explore its sales without Power BI licences. I built an interactive dashboard as a single HTML file that opens in any browser, on desktop or phone. It filters by channel, category, customer type and months, shows revenue, gross profit, margin and average order, writes key insights automatically, and exports the filtered orders to CSV. Every figure was cross-checked against an independent Python calculation. Synthetic data.

**Skills**
Data Visualization · Dashboard · JavaScript · HTML · Chart.js · Python · Data Analysis

**Files to attach**
`screenshots/dashboard_desktop.png` as the cover image, plus a link to the live page (GitHub Pages) once it is published.

---

## Full case study

### Problem
Small businesses often keep sales in spreadsheets and want a dashboard they can share with a partner or manager. Power BI or Tableau mean licences, accounts and a learning curve, and a static PDF cannot be filtered.

### Solution
One self-contained HTML file with the data embedded:

- Six KPI cards: revenue, gross profit, margin, orders, average order value, returning customers.
- Monthly revenue stacked by channel with a gross profit line, revenue by channel, by weekday and by category, plus a category table.
- Filters for channel, category, customer type and month range. Charts are clickable: a click on a channel or category bar filters the whole page.
- Automatic key insights that rewrite themselves for the current filters.
- A CSV export of the filtered orders.
- A responsive layout that works on a phone.

A short Python build script embeds the data, so the dashboard can be refreshed from a new export in one command. It can also take the clean output of an automated data-cleaning script.

### Result
- 914 orders, 3 channels, 4 categories in a 46 KB file that opens instantly.
- **Verified:** KPIs matched an independent pandas calculation for eight filter combinations, and the unfiltered view matches the Excel and Power BI versions of the same report to the dollar (revenue $58,482, gross profit $31,713, margin 54.2%).

### Tools
HTML, CSS, JavaScript, Chart.js, Python.

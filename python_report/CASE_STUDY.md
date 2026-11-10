# Case study: Retail Sales Analysis in Python

> Portfolio project built on synthetic data. No real business or customer data is used.

## Upwork portfolio fields

**Project title** (≤ 70 characters)
Sales Data Analysis in Python: Seasonality, Profit Drivers & Advice

**Your role**
Data analyst

**Project description** (≤ 600 characters)
Portfolio scenario: a small retail shop has a year of order data but no clear picture of what drives profit. I analysed 914 orders in Python (pandas, matplotlib) and answered six business questions: seasonality, category profit, channel differences, customer loyalty, weekday patterns and order size. The report ends with six concrete recommendations and comes as a PDF for the owner and a Jupyter notebook with all the code. Synthetic data.

**Skills**
Data Analysis · Python · pandas · Matplotlib · Jupyter Notebook · Data Visualization · Business Analysis

**Files to attach**
`figures/1_monthly_revenue.png` as the cover image, `sales_analysis_report.pdf`, and a link to the notebook on GitHub.

---

## Full case study

### Problem
A small shop records every order but only ever looks at the monthly total. It does not know which weeks, products, channels or customers really make its profit, so planning is based on gut feeling.

### Approach
1. **Data check:** missing values, duplicates, date range, cost assumptions. Nothing is analysed before the data is trusted.
2. **Profit, not just revenue:** gross profit per order from each category's cost-of-goods %.
3. **Six questions,** each answered with a chart, a table and a short plain-English conclusion.
4. **Recommendations** tied directly to the numbers.

### What the analysis found (synthetic data)
- Three months (Feb, May, Dec) bring 35% of revenue, and the peaks are concentrated in single holiday weeks: Valentine's week is 47% of February.
- A niche category (events and weddings) is 3% of orders but 25% of gross profit. One month without event orders becomes the weakest month of the year.
- Gifts are 22% of orders but 10% of profit, which makes them better as add-ons than as a product line.
- Phone customers spend the most per order and return most often.
- Sunday has the fewest orders but the largest baskets.

### Recommendations delivered
Plan stock and staff for three short peaks; grow the events business on purpose; sell gifts as add-ons; look after phone customers; raise the value of repeat orders; review Sunday hours against Sunday profit.

### Deliverables
- A PDF/HTML report without code for the business owner.
- The full Jupyter notebook with code and outputs.
- Chart images ready for slides.

### Tools
Python, pandas, matplotlib, Jupyter, nbconvert.

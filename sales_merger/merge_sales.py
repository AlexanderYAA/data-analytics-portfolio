"""Merge monthly sales exports from several order channels into one report.

Usage:
    python merge_sales.py                      # input/ -> output/
    python merge_sales.py --input D:/sales --output D:/reports

Every CSV / XLSX file in the input folder is read, its columns are matched
to a standard set using config.json, dates, numbers, categories and customer
types are normalised, problem rows are logged, and the result is written to:

  output/sales_report_<from>_to_<to>.xlsx   formatted report for people
  output/Orders.csv                          clean table for Power BI
"""

import argparse
import csv
import json
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

REQUIRED = ["order_id", "date", "category", "units", "unit_price"]
OPTIONAL = ["customer"]
HERE = Path(__file__).parent


# ---------------------------------------------------------------- issues log

class Issues:
    """Collects every problem found, so nothing is dropped silently."""

    COLUMNS = ["Severity", "File", "Row", "Order ID", "Problem", "Action"]

    def __init__(self):
        self.rows = []

    def add(self, severity, file, row, order_id, problem, action):
        self.rows.append(dict(zip(self.COLUMNS,
                                  [severity, file, row, order_id, problem, action])))

    def frame(self):
        return pd.DataFrame(self.rows, columns=self.COLUMNS)


# ---------------------------------------------------------------- reading

def norm(text):
    return " ".join(str(text).strip().lower().split())


def build_lookup(mapping):
    """{"standard": ["alias", ...]} -> {"alias": "standard"}"""
    lookup = {}
    for standard, aliases in mapping.items():
        lookup[norm(standard)] = standard
        for alias in aliases:
            lookup[norm(alias)] = standard
    return lookup


def read_csv(path):
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "cp1251", "latin-1"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    try:
        sep = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|").delimiter
    except csv.Error:
        sep = ","
    df = pd.read_csv(path, sep=sep, dtype=str, encoding=encoding,
                     keep_default_na=False, skip_blank_lines=True)
    return df, 0


def read_excel(path, col_lookup):
    """Find the header row: hand-kept logs often have title lines above the table."""
    for df in pd.read_excel(path, sheet_name=None, header=None, dtype=object).values():
        for i in range(min(15, len(df))):
            hits = sum(norm(v) in col_lookup for v in df.iloc[i] if pd.notna(v))
            if hits >= 3:
                body = df.iloc[i + 1:]
                body.columns = [str(v) if pd.notna(v) else f"col{j}"
                                for j, v in enumerate(df.iloc[i])]
                return body.dropna(how="all"), i
    return None, None


def parse_number(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace("$", "").replace("\u00a0", "").replace(" ", "")
    if not s:
        return None
    if "," in s and "." in s:          # 1,234.50
        s = s.replace(",", "")
    elif "," in s:                     # 46,86
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def parse_dates(series, formats, file, issues):
    """Use the first format that reads every date in the file.

    If several formats fit (e.g. 03/04/2025 can be 4 March or 3 April),
    the first one in config.json wins and the file gets a warning.
    """
    values = series.where(series.astype(str).str.strip() != "")
    filled = values.notna()
    if values[filled].map(lambda v: isinstance(v, datetime)).all() and filled.any():
        return pd.to_datetime(values)
    text = values.astype(str)
    fits = [f for f in formats
            if pd.to_datetime(text[filled], format=f, errors="coerce").notna().all()]
    if not fits:
        return pd.to_datetime(text, errors="coerce", dayfirst=True)
    if len(fits) > 1:
        issues.add("Warning", file, "", "",
                   f"Dates fit several formats: {', '.join(fits)}",
                   f"Used {fits[0]} - check the dates")
    return pd.to_datetime(text, format=fits[0], errors="coerce")


def is_filler(order_id):
    """Blank lines and TOTAL lines that people add to hand-kept logs."""
    s = norm(order_id)
    return s in ("", "nan", "none") or s.startswith(("total", "итого"))


def load_file(path, cfg, lookups, issues):
    if path.suffix.lower() == ".csv":
        df, header_row = read_csv(path)
    else:
        df, header_row = read_excel(path, lookups["columns"])
    if df is None:
        issues.add("Error", path.name, "", "", "Header row not found", "File skipped")
        return None, 0

    rename = {c: lookups["columns"][norm(c)] for c in df.columns
              if norm(c) in lookups["columns"]}
    df = df.rename(columns=rename)
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        issues.add("Error", path.name, "", "",
                   f"Missing columns: {', '.join(missing)}", "File skipped")
        return None, 0
    for col in OPTIONAL:
        if col not in df.columns:
            df[col] = ""

    df = df[REQUIRED + OPTIONAL].copy()
    df["source_row"] = df.index + header_row + 2   # row number as seen in Excel
    df["source_file"] = path.name

    filler = df["order_id"].map(is_filler)
    for _, r in df[filler].iterrows():
        issues.add("Info", path.name, r.source_row, "",
                   f"Blank or total line ('{r.order_id}')", "Skipped")
    df = df[~filler]
    read_count = len(df)

    prefix = path.stem.split("_")[0].lower()
    channel = cfg["channels"].get(prefix)
    if channel is None:
        channel = prefix.title()
        issues.add("Warning", path.name, "", "", f"Unknown channel prefix '{prefix}'",
                   f"Counted as channel '{channel}'; add it to config.json")
    df["channel"] = channel

    df["order_id"] = df["order_id"].astype(str).str.strip()
    df["date"] = parse_dates(df["date"], cfg["date_formats"], path.name, issues)
    df["units"] = df["units"].map(parse_number)
    df["unit_price"] = df["unit_price"].map(parse_number)
    raw_cat = df["category"].fillna("").astype(str)
    raw_cust = df["customer"].fillna("").astype(str)
    df["category"] = raw_cat.map(lambda v: lookups["categories"].get(norm(v)))
    df["customer"] = raw_cust.map(lambda v: lookups["customers"].get(norm(v)))

    period = cfg.get("period") or {}
    start = pd.Timestamp(period["from"]) if period.get("from") else None
    end = pd.Timestamp(period["to"]) if period.get("to") else None

    keep = []
    for idx, r in df.iterrows():
        bad = [name for name in ("date", "units", "unit_price") if pd.isna(r[name])]
        if bad:
            issues.add("Error", path.name, r.source_row, r.order_id,
                       f"Empty or unreadable: {', '.join(bad)}", "Row excluded")
            continue
        if (start is not None and r.date < start) or (end is not None and r.date > end):
            issues.add("Error", path.name, r.source_row, r.order_id,
                       f"Date {r.date:%Y-%m-%d} is outside the reporting period "
                       f"{period.get('from', '…')} – {period.get('to', '…')}",
                       "Row excluded")
            continue
        if r.category is None:
            issues.add("Warning", path.name, r.source_row, r.order_id,
                       f"Unknown category '{raw_cat[idx]}'",
                       "Counted as 'Other'; add it to config.json")
            df.at[idx, "category"] = "Other"
        if r.customer is None:
            if raw_cust[idx].strip():
                issues.add("Warning", path.name, r.source_row, r.order_id,
                           f"Unknown customer type '{raw_cust[idx]}'",
                           "Counted as 'Unknown'; add it to config.json")
            df.at[idx, "customer"] = "Unknown"
        if r.units < 0:
            issues.add("Info", path.name, r.source_row, r.order_id,
                       "Negative units (return)", "Kept, reduces revenue")
        keep.append(idx)
    return df.loc[keep], read_count


# ---------------------------------------------------------------- merging

def merge(input_dir, cfg):
    issues = Issues()
    lookups = {name: build_lookup(cfg[name])
               for name in ("columns", "categories", "customers")}
    files = sorted(p for p in input_dir.iterdir()
                   if p.suffix.lower() in (".csv", ".xlsx", ".xls")
                   and not p.name.startswith("~$"))
    frames, stats = [], []
    for path in files:
        df, read_count = load_file(path, cfg, lookups, issues)
        if df is not None and len(df):
            frames.append(df)
        stats.append({"File": path.name,
                      "Channel": df["channel"].iloc[0] if df is not None and len(df) else "",
                      "Rows read": read_count})
    stats = pd.DataFrame(stats, columns=["File", "Channel", "Rows read"])
    if not frames:
        return None, issues, stats

    data = pd.concat(frames, ignore_index=True)
    dup = data.duplicated(subset=["order_id"], keep="first")
    for _, r in data[dup].iterrows():
        issues.add("Warning", r.source_file, r.source_row, r.order_id,
                   "Duplicate order (already in an earlier file)", "Row excluded")
    data = data[~dup].copy()

    stats["Rows loaded"] = stats["File"].map(
        data.groupby("source_file").size()).fillna(0).astype(int)
    stats["Issues"] = stats["File"].map(
        issues.frame().groupby("File").size()).fillna(0).astype(int)

    data["revenue"] = data["units"] * data["unit_price"]
    cogs_pct = data["category"].map(cfg.get("cogs_pct", {}))
    for cat in sorted(data.loc[cogs_pct.isna(), "category"].unique()):
        issues.add("Warning", "", "", "", f"No COGS % for category '{cat}'",
                   "Gross profit not counted for it; add it to config.json")
    data["cogs"] = data["revenue"] * cogs_pct
    data["gross_profit"] = data["revenue"] - data["cogs"]
    data["month"] = data["date"].dt.to_period("M").dt.to_timestamp()
    data = data.sort_values(["date", "order_id"]).reset_index(drop=True)
    return data, issues, stats


# ---------------------------------------------------------------- report

NAVY, TEAL, LIGHT = "1F3A5F", "2A9D8F", "EEF3F8"
HEAD_FILL = PatternFill("solid", fgColor=NAVY)
HEAD_FONT = Font(bold=True, color="FFFFFF")
MONEY, MONEY2, INT, PCT, DATE = '"$"#,##0', '"$"#,##0.00', "#,##0", "0.0%", "yyyy-mm-dd"


def write_table(ws, df, top=1, formats=None, total_row=False):
    formats = formats or {}
    for j, col in enumerate(df.columns, start=1):
        c = ws.cell(row=top, column=j, value=col)
        c.fill, c.font = HEAD_FILL, HEAD_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, row in enumerate(df.itertuples(index=False), start=1):
        for j, value in enumerate(row, start=1):
            if isinstance(value, pd.Timestamp):
                value = value.to_pydatetime()
            elif pd.isna(value):
                value = None
            c = ws.cell(row=top + i, column=j, value=value)
            if df.columns[j - 1] in formats:
                c.number_format = formats[df.columns[j - 1]]
    if total_row:
        for j in range(1, len(df.columns) + 1):
            c = ws.cell(row=top + len(df), column=j)
            c.font, c.fill = Font(bold=True), PatternFill("solid", fgColor=LIGHT)
    for j, col in enumerate(df.columns, start=1):
        shown = [12 if str(v).startswith("=") else len(str(v))
                 for v in df.iloc[:, j - 1].head(500)]
        width = max([len(str(col))] + shown)
        letter = get_column_letter(j)
        ws.column_dimensions[letter].width = max(
            ws.column_dimensions[letter].width or 0, min(width + 3, 60))
    return top + len(df)


# The report is "live": the Data sheet holds the orders, Revenue / COGS /
# Gross profit there are formulas, every summary is SUMIFS / COUNTIFS over
# Data, and COGS % sits on the Assumptions sheet. Edit an order or a % in
# Excel and the whole report recalculates.
#
# Data columns:  A Date  B Channel  C Order ID  D Category  E Customer
#                F Units  G Unit price  H Revenue  I COGS  J Gross profit
#                K Source file  L Source row
D_DATE, D_CH, D_CAT, D_CUST, D_UNITS, D_REV, D_GP, D_FILE = (
    "Data!$A:$A", "Data!$B:$B", "Data!$D:$D", "Data!$E:$E",
    "Data!$F:$F", "Data!$H:$H", "Data!$J:$J", "Data!$K:$K")


def returning_formula(crit=""):
    ret = f'COUNTIFS({D_CUST},"Returning"{crit})'
    new = f'COUNTIFS({D_CUST},"New"{crit})'
    return f'=IFERROR({ret}/({ret}+{new}),"")'


def group_sheet(ws, data, key, label, data_col):
    """By channel / By category: one row per value, all cells are formulas."""
    order = data.groupby(key)["revenue"].sum().sort_values(ascending=False).index
    n = len(order)
    tot = n + 2
    rows = []
    for i, name in enumerate(order, start=2):
        crit = f",{data_col},$A{i}"
        rows.append({
            label: name,
            "Revenue": f"=SUMIFS({D_REV}{crit})",
            "Gross profit": f"=SUMIFS({D_GP}{crit})",
            "Margin": f'=IFERROR(C{i}/B{i},"")',
            "Orders": f"=COUNTIFS({data_col},$A{i})",
            "Units": f"=SUMIFS({D_UNITS}{crit})",
            "Avg order": f'=IFERROR(B{i}/E{i},"")',
            "Returning": returning_formula(crit),
            "Share": f'=IFERROR(B{i}/B${tot},"")',
        })
    rows.append({
        label: "Total",
        "Revenue": f"=SUM(B2:B{tot - 1})",
        "Gross profit": f"=SUM(C2:C{tot - 1})",
        "Margin": f'=IFERROR(C{tot}/B{tot},"")',
        "Orders": f"=SUM(E2:E{tot - 1})",
        "Units": f"=SUM(F2:F{tot - 1})",
        "Avg order": f'=IFERROR(B{tot}/E{tot},"")',
        "Returning": returning_formula(),
        "Share": f"=SUM(I2:I{tot - 1})",
    })
    fmt = {"Revenue": MONEY, "Gross profit": MONEY, "Margin": PCT, "Orders": INT,
           "Units": INT, "Avg order": MONEY2, "Returning": PCT, "Share": PCT}
    write_table(ws, pd.DataFrame(rows), formats=fmt, total_row=True)
    return tot


def build_report(data, issues, stats, out_path, cfg):
    wb = Workbook()
    issues_df = issues.frame()
    period = f"{data['date'].min():%d %b %Y} – {data['date'].max():%d %b %Y}"
    n = len(data)

    ws = wb.active
    ws.title = "Summary"
    wm = wb.create_sheet("By month")
    wc = wb.create_sheet("By channel")
    wg = wb.create_sheet("By category")
    wd = wb.create_sheet("Data")
    wi = wb.create_sheet("Issues")
    wa = wb.create_sheet("Assumptions")

    # Assumptions: COGS % per category (the Data sheet looks it up here)
    wa["A1"] = "Assumptions"
    wa["A1"].font = Font(bold=True, size=14, color=NAVY)
    wa["A2"] = "Cost of goods sold as % of revenue. Change a % and the report recalculates."
    wa["A2"].font = Font(italic=True, color="666666")
    cogs = pd.DataFrame(list(cfg.get("cogs_pct", {}).items()),
                        columns=["Category", "COGS %"])
    cogs_end = write_table(wa, cogs, top=3, formats={"COGS %": PCT})
    cogs_range = f"Assumptions!$A$4:$B${max(cogs_end, 4)}"
    p = cfg.get("period") or {}
    wa.cell(row=cogs_end + 2, column=1,
            value=f"Reporting period: {p.get('from', 'any')} – {p.get('to', 'any')} "
                  "(orders outside it are excluded and listed on the Issues sheet)")

    # Data
    out = pd.DataFrame({
        "Date": data["date"], "Channel": data["channel"], "Order ID": data["order_id"],
        "Category": data["category"], "Customer": data["customer"],
        "Units": data["units"], "Unit price": data["unit_price"]})
    rows = range(2, n + 2)
    out["Revenue"] = [f"=F{r}*G{r}" for r in rows]
    out["COGS"] = [f'=IFERROR(H{r}*VLOOKUP(D{r},{cogs_range},2,FALSE),"")' for r in rows]
    out["Gross profit"] = [f'=IF(I{r}="","",H{r}-I{r})' for r in rows]
    out["Source file"] = data["source_file"].values
    out["Source row"] = data["source_row"].values
    last = write_table(wd, out, formats={"Date": DATE, "Units": INT,
                                         "Unit price": MONEY2, "Revenue": MONEY2,
                                         "COGS": MONEY2, "Gross profit": MONEY2})
    wd.auto_filter.ref = f"A1:{get_column_letter(out.shape[1])}{last}"

    # By channel / By category
    ch_tot = group_sheet(wc, data, "channel", "Channel", D_CH)
    group_sheet(wg, data, "category", "Category", D_CAT)

    # By month: channels as columns
    channels = list(data.groupby("channel")["revenue"].sum()
                    .sort_values(ascending=False).index)
    months = sorted(data["month"].unique())
    m_tot = len(months) + 2
    L = {name: get_column_letter(j) for j, name in enumerate(
        ["Month"] + channels + ["Total", "Gross profit", "Margin", "Orders"], start=1)}
    first_ch, last_ch = L[channels[0]], L[channels[-1]]
    rows = []
    for i, m in enumerate(months, start=2):
        when = f',{D_DATE},">="&$A{i},{D_DATE},"<"&EDATE($A{i},1)'
        row = {"Month": pd.Timestamp(m)}
        for ch in channels:
            row[ch] = f"=SUMIFS({D_REV},{D_CH},{L[ch]}$1{when})"
        row["Total"] = f"=SUM({first_ch}{i}:{last_ch}{i})"
        row["Gross profit"] = f"=SUMIFS({D_GP}{when})"
        row["Margin"] = f'=IFERROR({L["Gross profit"]}{i}/{L["Total"]}{i},"")'
        row["Orders"] = f"=COUNTIFS({when[1:]})"
        rows.append(row)
    total = {"Month": "Total"}
    for col in channels + ["Total", "Gross profit", "Orders"]:
        total[col] = f"=SUM({L[col]}2:{L[col]}{m_tot - 1})"
    total["Margin"] = f'=IFERROR({L["Gross profit"]}{m_tot}/{L["Total"]}{m_tot},"")'
    rows.append(total)
    write_table(wm, pd.DataFrame(rows),
                formats={"Month": "[$-409]mmm yyyy",  # English month names in any locale
                         **{c: MONEY for c in channels + ["Total", "Gross profit"]},
                         "Margin": PCT, "Orders": INT}, total_row=True)
    wm.column_dimensions["A"].width = 12

    # Summary
    ws["A1"] = "Sales report"
    ws["A1"].font = Font(bold=True, size=18, color=NAVY)
    ws["A2"] = f"Period: {period}   ·   Generated {datetime.now():%Y-%m-%d %H:%M}"
    ws["A2"].font = Font(italic=True, color="666666")
    if cfg.get("report_note"):
        ws["A3"] = cfg["report_note"]
        ws["A3"].font = Font(bold=True, color="C0392B")
    kpis = [("Revenue", f"='By channel'!B{ch_tot}", MONEY),
            ("Gross profit", f"='By channel'!C{ch_tot}", MONEY),
            ("Gross margin", f"='By channel'!D{ch_tot}", PCT),
            ("Orders", f"='By channel'!E{ch_tot}", INT),
            ("Avg order value", f"='By channel'!G{ch_tot}", MONEY2),
            ("Returning customers", f"='By channel'!H{ch_tot}", PCT),
            ("Issues logged", "=MAX(COUNTA(Issues!$A:$A)-1,0)", INT)]
    for j, (label, formula, fmt) in enumerate(kpis, start=1):
        ws.cell(row=4, column=j, value=label).font = Font(color="666666")
        c = ws.cell(row=5, column=j, value=formula)
        c.number_format, c.font = fmt, Font(bold=True, size=16, color=NAVY)
        ws.column_dimensions[get_column_letter(j)].width = 20

    ws["A7"] = "Files processed"
    ws["A7"].font = Font(bold=True, size=12, color=NAVY)
    files = stats[["File", "Channel", "Rows read"]].copy()
    first, f_tot = 9, 9 + len(files)
    files["Rows loaded"] = [f"=COUNTIFS({D_FILE},$A{r})" for r in range(first, f_tot)]
    files["Issues"] = [f"=COUNTIFS(Issues!$B:$B,$A{r})" for r in range(first, f_tot)]
    files.loc[len(files)] = ["Total", None, f"=SUM(C{first}:C{f_tot - 1})",
                             f"=SUM(D{first}:D{f_tot - 1})",
                             "=MAX(COUNTA(Issues!$A:$A)-1,0)"]
    write_table(ws, files, top=8, total_row=True,
                      formats={"Rows read": INT, "Rows loaded": INT, "Issues": INT})

    chart = BarChart()
    chart.type, chart.grouping, chart.overlap = "col", "stacked", 100
    chart.title = "Revenue by month and channel"
    chart.y_axis.numFmt = MONEY
    chart.x_axis.delete = chart.y_axis.delete = False  # openpyxl hides axes otherwise
    chart.y_axis.majorGridlines = None
    chart.legend.position, chart.legend.overlay = "b", False
    chart.add_data(Reference(wm, min_col=2, max_col=1 + len(channels),
                             min_row=1, max_row=m_tot - 1), titles_from_data=True)
    chart.set_categories(Reference(wm, min_col=1, min_row=2, max_row=m_tot - 1))
    chart.height, chart.width = 9, 24
    ws.add_chart(chart, "G8")  # right of the files table, visible on open

    # Issues
    if issues_df.empty:
        wi["A1"] = "No issues found."
    else:
        order = {"Error": 0, "Warning": 1, "Info": 2}
        sort_keys = {"Severity": lambda s: s.map(order),
                     "File": lambda s: s,
                     "Row": lambda s: pd.to_numeric(s, errors="coerce").fillna(0)}
        issues_df = issues_df.sort_values(["Severity", "File", "Row"],
                                          key=lambda s: sort_keys[s.name](s))
        last = write_table(wi, issues_df)
        wi.auto_filter.ref = f"A1:F{last}"
        colors = {"Error": "C0392B", "Warning": "D68910", "Info": TEAL}
        for r in range(2, last + 1):
            c = wi.cell(row=r, column=1)
            c.font = Font(bold=True, color=colors.get(c.value, "000000"))

    for sheet in wb.worksheets:
        if sheet.title != "Summary":
            sheet.freeze_panes = "A2"
    wb.save(out_path)


def write_powerbi_csv(data, out_path):
    """Same layout as the Orders.csv the Power BI dashboard is built on."""
    out = data[["order_id", "date", "channel", "category", "customer",
                "units", "unit_price"]].copy()
    out.columns = ["OrderID", "Date", "Channel", "Category", "Customer",
                   "Units", "UnitPrice"]
    out["Units"] = out["Units"].astype(int)
    out = out.sort_values("OrderID")
    out.to_csv(out_path, index=False, date_format="%Y-%m-%d", encoding="utf-8")


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(description="Merge sales files into one Excel report.")
    ap.add_argument("--input", default=HERE / "input", type=Path)
    ap.add_argument("--output", default=HERE / "output", type=Path)
    ap.add_argument("--config", default=HERE / "config.json", type=Path)
    args = ap.parse_args()

    cfg = json.loads(args.config.read_text(encoding="utf-8"))
    if not args.input.is_dir():
        sys.exit(f"Input folder not found: {args.input}")

    data, issues, stats = merge(args.input, cfg)
    if data is None:
        if issues.rows:
            print(issues.frame().to_string(index=False))
        sys.exit("No usable files found.")

    args.output.mkdir(parents=True, exist_ok=True)
    name = f"sales_report_{data['date'].min():%Y-%m}_to_{data['date'].max():%Y-%m}.xlsx"
    build_report(data, issues, stats, args.output / name, cfg)
    write_powerbi_csv(data, args.output / "Orders.csv")

    counts = issues.frame()["Severity"].value_counts()
    print(f"Files: {len(stats)}   orders: {data['order_id'].nunique():,}   "
          f"revenue: ${data['revenue'].sum():,.2f}   "
          f"gross profit: ${data['gross_profit'].sum():,.2f}   "
          f"period: {data['date'].min():%Y-%m-%d} to {data['date'].max():%Y-%m-%d}")
    print("Issues: " + (", ".join(f"{k} {v}" for k, v in counts.items()) or "none"))
    print(f"Report: {args.output / name}")
    print(f"Power BI table: {args.output / 'Orders.csv'}")


if __name__ == "__main__":
    main()

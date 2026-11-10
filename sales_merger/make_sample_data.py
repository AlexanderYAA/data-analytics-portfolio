"""Split the shop's 2025 orders into the raw monthly exports the merger works on.

A small retail shop takes orders through three channels, and each channel
keeps its own monthly file in its own format:

  instore_2025-MM.csv   till export: comma CSV, ISO dates; the January and
                        December exports were run with a wrong date range and
                        also contain orders from 2024 / 2026
  phone_2025-MM.xlsx    order log kept by hand in Excel: title lines above
                        the table, d.m.Y dates, a TOTAL line at the bottom
  online_2025-MM.csv    website export: semicolon CSV, decimal comma,
                        m/d/Y dates, own spelling of categories, overlaps
                        with the previous export (duplicates), a test order

The source is source/Orders.csv - the same synthetic data as the Power BI
dashboard, so the merged result can be checked against it to the cent.
"""

from pathlib import Path

import pandas as pd

HERE = Path(__file__).parent
SOURCE = HERE / "source" / "Orders.csv"
OUT = HERE / "input"

ONLINE_CATEGORY = {
    "Bouquets": "bouquet",
    "Potted plants": "Potted Plants",
    "Gifts & add-ons": "Gifts and add-ons",
    "Events & weddings": "Events/Weddings",
}
ONLINE_CUSTOMER = {"New": "new customer", "Returning": "returning customer"}
PHONE_CUSTOMER = {"New": "First-time", "Returning": "Repeat"}


OUT_OF_PERIOD = {  # orders that belong to the 2024 / 2026 reports
    1: [("ORD-9001", "2024-12-30", "Bouquets", "New", 1, 54.90),
        ("ORD-9002", "2024-12-31", "Gifts & add-ons", "Returning", 2, 18.50)],
    12: [("ORD-9003", "2026-01-01", "Potted plants", "New", 1, 32.00)],
}


def write_instore(df, path, month):
    cols = ["OrderID", "Date", "Category", "Customer", "Units", "UnitPrice"]
    out = df[cols].copy()
    out["Date"] = out["Date"].dt.strftime("%Y-%m-%d")
    extra = pd.DataFrame(OUT_OF_PERIOD.get(month, []), columns=cols)
    out = pd.concat([p for p in (extra, out) if len(p)], ignore_index=True)
    out = out.sort_values("Date", kind="stable")
    out.to_csv(path, index=False)


def write_phone(df, path, month_label):
    log = pd.DataFrame({
        "Order #": df["OrderID"],
        "Order date": df["Date"].dt.strftime("%d.%m.%Y"),
        "Product category": df["Category"],
        "Client": df["Customer"].map(PHONE_CUSTOMER),
        "Qty": df["Units"],
        "Price, $": df["UnitPrice"],
    })
    total = pd.DataFrame([{"Order #": "TOTAL", "Qty": log["Qty"].sum()}])
    log = pd.concat([log, total], ignore_index=True)
    with pd.ExcelWriter(path) as xw:
        log.to_excel(xw, sheet_name="Phone orders", index=False, startrow=3)
        ws = xw.sheets["Phone orders"]
        ws["A1"] = "Phone orders log"
        ws["A2"] = month_label


def write_online(df, path, overlap, test_order):
    out = pd.concat([p for p in (overlap, df) if len(p)], ignore_index=True)
    if test_order:  # a test order someone placed on the site, with no price
        test = df.iloc[0].copy()
        test["OrderID"], test["UnitPrice"] = "TEST-001", float("nan")
        out.loc[len(out)] = test
    pd.DataFrame({
        "Order ID": out["OrderID"],
        "Order Date": out["Date"].dt.strftime("%m/%d/%Y"),
        "Category": out["Category"].map(ONLINE_CATEGORY),
        "Customer type": out["Customer"].map(ONLINE_CUSTOMER),
        "Units": out["Units"],
        "Unit Price": out["UnitPrice"].map(
            lambda v: "" if pd.isna(v) else f"{v:.2f}".replace(".", ",")),
    }).to_csv(path, sep=";", index=False, encoding="utf-8-sig")


def main():
    orders = pd.read_csv(SOURCE, parse_dates=["Date"])
    OUT.mkdir(exist_ok=True)
    overlap = orders.iloc[0:0]
    for month, by_month in orders.groupby(orders["Date"].dt.to_period("M")):
        tag = str(month)                                  # 2025-01
        ch = {c: g for c, g in by_month.groupby("Channel")}
        write_instore(ch["In-store"], OUT / f"instore_{tag}.csv", month.month)
        write_phone(ch["Phone"], OUT / f"phone_{tag}.xlsx", month.strftime("%B %Y"))
        write_online(ch["Online"], OUT / f"online_{tag}.csv", overlap,
                     test_order=month.month in (3, 9))
        overlap = ch["Online"].tail(3)                    # repeated in the next export
    print(f"{len(list(OUT.iterdir()))} sample files written to {OUT}")


if __name__ == "__main__":
    main()

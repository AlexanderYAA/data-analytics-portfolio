"""Build the interactive dashboard: embed the orders into template.html.

    python build_dashboard.py            # data/ -> index.html

The result is one self-contained HTML file (the only external file is the
Chart.js library from a CDN), so it can be opened locally, attached, or
hosted on GitHub Pages.
"""

import csv
import json
from pathlib import Path

HERE = Path(__file__).parent


def main():
    cogs = {r["Category"]: float(r["COGS_Pct"])
            for r in csv.DictReader(open(HERE / "data" / "Categories.csv", encoding="utf-8"))}
    channels, categories, customers = [], list(cogs), []
    rows = []
    for r in csv.DictReader(open(HERE / "data" / "Orders.csv", encoding="utf-8")):
        for value, bucket in ((r["Channel"], channels), (r["Customer"], customers)):
            if value not in bucket:
                bucket.append(value)
        year, month, day = map(int, r["Date"].split("-"))
        rows.append([
            int(r["OrderID"].split("-")[1]), year, month, day,
            channels.index(r["Channel"]), categories.index(r["Category"]),
            customers.index(r["Customer"]), int(r["Units"]),
            round(float(r["UnitPrice"]) * 100),          # cents, no float drift
        ])
    data = {
        "channels": sorted(channels, key=["In-store", "Online", "Phone"].index),
        "categories": categories,
        "customers": customers,
        "cogs": [cogs[c] for c in categories],
        "rows": rows,
    }
    # channel indexes must follow the sorted order
    remap = [data["channels"].index(c) for c in channels]
    for row in rows:
        row[4] = remap[row[4]]

    html = (HERE / "template.html").read_text(encoding="utf-8")
    html = html.replace("/*__DATA__*/null", json.dumps(data, separators=(",", ":")))
    (HERE / "index.html").write_text(html, encoding="utf-8")
    print(f"index.html written: {len(rows)} orders, {len(html) / 1024:.0f} KB")


if __name__ == "__main__":
    main()

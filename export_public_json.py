"""
Export FX Watch public data - GitHub repo version.

Reads SGDUSD_Tracker.xlsx (sheet "new") and writes docs/fx-public.json: a
plain JSON object with five series ("usd","eur","gbp","rmb","rebate"), each
a list of [iso_date, value] pairs, built fresh from the xlsx every run.

This file is served publicly via GitHub Pages (Settings > Pages > Deploy
from branch > main > /docs), so the private repo itself stays private while
this one small derived file is readable with no auth. The FX Watch Claude
artifact dashboard fetches this file directly, giving it exact parity with
the authoritative xlsx (same source, same values), with no dependency on
anyone's computer.

Conventions (must match fx_update.py / the dashboard exactly):
  - "usd"    = SGD per 1 USD   (column B)
  - "eur"    = SGD per 1 EUR   (column C)
  - "rmb"    = RMB per 1 USD   (column D)
  - "gbp"    = SGD per 1 GBP   (column F), only for dates >= 2020-01-01
  - "rebate" = rmb value minus 0.08, rounded to 4 decimals, wherever rmb
    is numeric (Heritage's RMB export-tax-rebate convention)

Date reconstruction: most of column A is a chain of "=A[prev]+1" formulas
with no cached value (openpyxl data_only=True reads these as None). Row 34
holds a literal anchor date, 2016-01-09; any row >= 35 with a None date is
reconstructed as 2016-01-09 + (row - 34) days. Rows with an actual cached
datetime (early rows, or any row Excel has since recalculated and saved)
use that value directly.
"""

import json
import datetime
import openpyxl

FILE_PATH = "SGDUSD_Tracker.xlsx"
SHEET_NAME = "new"
OUT_PATH = "docs/fx-public.json"

ANCHOR_ROW = 34
ANCHOR_DATE = datetime.date(2016, 1, 9)
GBP_START_DATE = datetime.date(2020, 1, 1)


def resolve_date(row, raw_value):
    if isinstance(raw_value, datetime.datetime):
        return raw_value.date()
    if isinstance(raw_value, datetime.date):
        return raw_value
    if row >= ANCHOR_ROW + 1:
        return ANCHOR_DATE + datetime.timedelta(days=row - ANCHOR_ROW)
    return None


def main():
    wb = openpyxl.load_workbook(FILE_PATH, data_only=True)
    ws = wb[SHEET_NAME]

    usd, eur, gbp, rmb, rebate = [], [], [], [], []

    for row in range(2, ws.max_row + 1):
        raw_date = ws.cell(row=row, column=1).value
        d = resolve_date(row, raw_date)
        if d is None:
            continue
        iso = d.isoformat()

        b = ws.cell(row=row, column=2).value
        c = ws.cell(row=row, column=3).value
        dd = ws.cell(row=row, column=4).value
        f = ws.cell(row=row, column=6).value

        if isinstance(b, (int, float)):
            usd.append([iso, b])
        if isinstance(c, (int, float)):
            eur.append([iso, c])
        if isinstance(dd, (int, float)):
            rmb.append([iso, dd])
            rebate.append([iso, round(dd - 0.08, 4)])
        if isinstance(f, (int, float)) and d >= GBP_START_DATE:
            gbp.append([iso, f])

    data = {"usd": usd, "eur": eur, "gbp": gbp, "rmb": rmb, "rebate": rebate}

    import os
    os.makedirs("docs", exist_ok=True)
    with open(OUT_PATH, "w") as out:
        json.dump(data, out, separators=(",", ":"))

    print(f"Wrote {OUT_PATH}: usd={len(usd)} eur={len(eur)} gbp={len(gbp)} rmb={len(rmb)} rebate={len(rebate)}")


if __name__ == "__main__":
    main()

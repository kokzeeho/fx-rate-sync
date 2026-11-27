"""
Export FX Watch public data - GitHub repo version.

Reads SGDUSD_Tracker.xlsx (sheet "new") and writes:
  - docs/fx-public.json: full history, five series ("usd","eur","gbp","rmb",
    "rebate"), each a list of [iso_date, value] pairs.
  - docs/fx-recent.json: same five series, but only the last N points each
    (small file, safe for fetch tools that truncate large responses) - used
    by the daily FX Watch dashboard refresh to catch up after any gap.

Both are served publicly via GitHub Pages (this repo is public; Pages
serves from main branch, /docs folder), so the FX Watch Claude artifact
dashboard can read them directly with no auth and no dependency on anyone's
computer - and since they're built straight from the same xlsx, the
dashboard gets exact parity with the authoritative Excel file.

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


def write_recent(data, n=35):
    """Small 'recent window' file (last n points per series) for fetch tools
    that truncate very large responses; fx-public.json keeps the full
    history for completeness/manual use."""
    recent = {k: v[-n:] for k, v in data.items()}
    with open("docs/fx-recent.json", "w") as out:
        json.dump(recent, out, separators=(",", ":"))
    print(f"Wrote docs/fx-recent.json: last {n} points per series")


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

    write_recent(data)


if __name__ == "__main__":
    main()

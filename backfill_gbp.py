"""
One-time backfill script: fills in the SGD/GBP column (F) for a date range
that's been sitting empty, using Frankfurter's historical time-series
endpoint. Calculates each row's date arithmetically (row 2's date + (row-2)
days) rather than reading cached formula values, since a large stretch of
cached date values in this sheet were found to be missing/stale.
"""

import sys
import datetime
import requests
import openpyxl

FILE_PATH = "SGDUSD_Tracker.xlsx"
SHEET_NAME = "new"
DATE_COL_IDX = 1
GBP_COL_IDX = 6
TEMPLATE_ROW = 2

START_DATE = "2020-01-01"
END_DATE = "2026-09-11"


def fetch_historical_gbp(start_date, end_date):
    url = f"https://api.frankfurter.dev/v1/{start_date}..{end_date}"
    r = requests.get(url, params={"base": "GBP", "symbols": "SGD"})
    r.raise_for_status()
    data = r.json()["rates"]
    return {datetime.date.fromisoformat(d): v["SGD"] for d, v in data.items()}


def copy_style(source, target):
    import copy
    target.font = copy.copy(source.font)
    target.alignment = copy.copy(source.alignment)
    target.number_format = source.number_format
    target.border = copy.copy(source.border)
    target.fill = copy.copy(source.fill)


def main():
    print(f"Fetching SGD/GBP rates from {START_DATE} to {END_DATE}...")
    gbp_by_date = fetch_historical_gbp(START_DATE, END_DATE)
    print(f"Got {len(gbp_by_date)} published rates from the API.")

    print(f"Opening {FILE_PATH}...")
    wb = openpyxl.load_workbook(FILE_PATH)
    ws = wb[SHEET_NAME]

    anchor_cell = ws.cell(row=TEMPLATE_ROW, column=DATE_COL_IDX).value
    anchor_date = anchor_cell.date() if isinstance(anchor_cell, datetime.datetime) else anchor_cell
    print(f"Anchor date at row {TEMPLATE_ROW}: {anchor_date}")

    start = datetime.date.fromisoformat(START_DATE)
    end = datetime.date.fromisoformat(END_DATE)

    start_row = TEMPLATE_ROW + (start - anchor_date).days
    end_row = TEMPLATE_ROW + (end - anchor_date).days
    print(f"Target row range: {start_row} to {end_row}")

    last_known_rate = None
    filled = 0
    carried = 0

    for row in range(start_row, end_row + 1):
        row_date = anchor_date + datetime.timedelta(days=row - TEMPLATE_ROW)

        if row_date in gbp_by_date:
            last_known_rate = gbp_by_date[row_date]
            filled += 1
        elif last_known_rate is not None:
            carried += 1
        else:
            continue

        target_cell = ws.cell(row=row, column=GBP_COL_IDX, value=last_known_rate)
        copy_style(ws.cell(row=TEMPLATE_ROW, column=GBP_COL_IDX), target_cell)

    wb.save(FILE_PATH)
    print(f"Done. {filled} rows filled with fresh published rates, {carried} rows carried forward (weekends/holidays).")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)

"""
One-time backfill script: fills in the USD/RMB column (D) for a date range
that's been sitting empty, using Frankfurter's historical time-series
endpoint (one API call covers the whole range, no key needed).

Run this locally on your Mac (it needs internet access, which your sandbox
version of Claude doesn't have). It preserves all existing formulas,
formatting, and the SGD/USD and SGD/EUR columns — only touches column D.
"""

import sys
import datetime
import requests
import openpyxl

FILE_PATH = "SGDUSD_Tracker.xlsx"  # <-- update if your filename differs
SHEET_NAME = "new"
DATE_COL_IDX = 1
RMB_COL_IDX = 4  # Column D = USD/RMB
TEMPLATE_ROW = 2  # earliest row — used as a style fallback

START_DATE = "2026-04-02"
END_DATE = "2026-09-08"


def fetch_historical_rmb(start_date, end_date):
    """One call, gets every published rate in the range. Weekends/holidays
    (ECB doesn't publish those) will be missing from the response — we carry
    forward the last known value for those, matching how columns B/C
    already behave in this tracker."""
    url = f"https://api.frankfurter.dev/v1/{start_date}..{end_date}"
    r = requests.get(url, params={"base": "USD", "symbols": "CNY"})
    r.raise_for_status()
    data = r.json()["rates"]  # {"2026-04-02": {"CNY": 7.12}, ...}
    return {datetime.date.fromisoformat(d): v["CNY"] for d, v in data.items()}


def main():
    print(f"Fetching USD/RMB rates from {START_DATE} to {END_DATE}...")
    rmb_by_date = fetch_historical_rmb(START_DATE, END_DATE)
    print(f"Got {len(rmb_by_date)} published rates from the API.")

    print(f"Opening {FILE_PATH}...")
    wb_values = openpyxl.load_workbook(FILE_PATH, data_only=True)
    ws_values = wb_values[SHEET_NAME]

    wb = openpyxl.load_workbook(FILE_PATH)  # preserves formulas
    ws = wb[SHEET_NAME]

    start = datetime.date.fromisoformat(START_DATE)
    end = datetime.date.fromisoformat(END_DATE)

    last_known_rate = None
    filled = 0
    carried = 0

    for row in range(2, ws_values.max_row + 1):
        cell_value = ws_values.cell(row=row, column=DATE_COL_IDX).value
        if cell_value is None:
            continue
        row_date = cell_value.date() if isinstance(cell_value, datetime.datetime) else cell_value

        if row_date < start:
            continue
        if row_date > end:
            break

        if row_date in rmb_by_date:
            last_known_rate = rmb_by_date[row_date]
            filled += 1
        elif last_known_rate is not None:
            carried += 1
        else:
            continue  # no rate available yet (before first published date)

        target_cell = ws.cell(row=row, column=RMB_COL_IDX, value=last_known_rate)
        template = ws.cell(row=TEMPLATE_ROW, column=RMB_COL_IDX)
        target_cell.font = template.font.copy()
        target_cell.alignment = template.alignment.copy()
        target_cell.number_format = template.number_format
        target_cell.border = template.border.copy()
        target_cell.fill = template.fill.copy()

    wb.save(FILE_PATH)
    print(f"Done. {filled} rows filled with fresh published rates, {carried} rows carried forward (weekends/holidays).")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)

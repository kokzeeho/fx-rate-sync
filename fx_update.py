"""
Daily FX rate updater — GitHub repo version.

Fetches SGD/USD, SGD/EUR, USD/RMB and writes them into today's row in the
xlsx file, which lives in this same repo. GitHub Actions commits the change
back automatically after this script runs. No cloud auth needed.
"""

import sys
import datetime
import requests
import openpyxl

FILE_PATH = "SGDUSD_Tracker.xlsx"  # <-- update this if your filename differs
SHEET_NAME = "new"
DATE_COL_IDX = 1   # Column A
RATE_COL_IDXS = [2, 3, 4]  # Columns B, C, D = SGD/USD, SGD/EUR, USD/RMB


def fetch_rates():
    """Fetch SGD/USD, SGD/EUR, USD/RMB from Frankfurter (ECB rates, no API key).
    Returns (rates_list, rate_date) — rate_date is the date the rates were
    actually published for (from the API itself), not today's clock date.
    """
    r1 = requests.get("https://api.frankfurter.dev/v1/latest", params={"base": "SGD", "symbols": "USD,EUR"})
    r1.raise_for_status()
    j1 = r1.json()
    d1 = j1["rates"]

    r2 = requests.get("https://api.frankfurter.dev/v1/latest", params={"base": "USD", "symbols": "CNY"})
    r2.raise_for_status()
    d2 = r2.json()["rates"]

    rate_date = datetime.date.fromisoformat(j1["date"])  # e.g. "2026-09-09"
    return [d1["USD"], d1["EUR"], d2["CNY"]], rate_date


def find_row_for_date(ws, target_date):
    last_filled_row = 1
    for row in range(2, ws.max_row + 1):
        cell_value = ws.cell(row=row, column=DATE_COL_IDX).value
        if cell_value is None:
            continue
        last_filled_row = row
        cell_date = cell_value.date() if isinstance(cell_value, datetime.datetime) else cell_value
        if cell_date == target_date:
            return row
    # Target date wasn't pre-filled — append after the last filled row.
    new_row = last_filled_row + 1
    ws.cell(row=new_row, column=DATE_COL_IDX, value=datetime.datetime.combine(target_date, datetime.time()))
    return new_row


def main():
    print("Fetching rates...")
    rates, rate_date = fetch_rates()
    print(f"Rates for {rate_date} (SGD/USD, SGD/EUR, USD/RMB): {rates}")

    print(f"Opening {FILE_PATH}...")
    wb = openpyxl.load_workbook(FILE_PATH)
    ws = wb[SHEET_NAME]

    row = find_row_for_date(ws, rate_date)
    print(f"Writing to row {row} (date {rate_date})...")
    for col_idx, value in zip(RATE_COL_IDXS, rates):
        ws.cell(row=row, column=col_idx, value=value)

    wb.save(FILE_PATH)
    print("Done.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)

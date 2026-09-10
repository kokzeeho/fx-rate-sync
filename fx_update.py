"""
Daily FX rate updater — GitHub repo version.

Fetches SGD/USD, SGD/EUR, USD/RMB and appends them as a new row at the bottom
of the xlsx file, which lives in this same repo. GitHub Actions commits the
change back automatically after this script runs. No cloud auth needed.

Conventions (confirmed against the tracker's historical data):
  - SGD/USD column = amount of SGD for 1 USD  (e.g. ~1.27)
  - SGD/EUR column = amount of SGD for 1 EUR  (e.g. ~1.47)
  - USD/RMB column = amount of RMB for 1 USD  (e.g. ~7.1)
"""

import sys
import datetime
import requests
import openpyxl
from openpyxl.worksheet.views import Selection

FILE_PATH = "SGDUSD_Tracker.xlsx"  # <-- update this if your filename differs
SHEET_NAME = "new"
DATE_COL_IDX = 1   # Column A
RATE_COL_IDXS = [2, 3, 4]  # Columns B, C, D = SGD/USD, SGD/EUR, USD/RMB
DATE_NUMBER_FORMAT = "d/m/yyyy"
RATE_NUMBER_FORMAT = "0.0000"


def fetch_rates():
    """Fetch SGD/USD, SGD/EUR, USD/RMB from Frankfurter (ECB rates, no API key).

    Uses base=USD and base=EUR (rather than base=SGD) so the returned rates
    are already in "amount of X for 1 USD/EUR" form, matching this tracker's
    column convention directly — no inverting needed.

    Returns (rates_list, rate_date) — rate_date is the date the rates were
    actually published for (from the API itself), not today's clock date.
    """
    r1 = requests.get("https://api.frankfurter.dev/v1/latest", params={"base": "USD", "symbols": "SGD,CNY"})
    r1.raise_for_status()
    j1 = r1.json()
    d1 = j1["rates"]

    r2 = requests.get("https://api.frankfurter.dev/v1/latest", params={"base": "EUR", "symbols": "SGD"})
    r2.raise_for_status()
    d2 = r2.json()["rates"]

    rate_date = datetime.date.fromisoformat(j1["date"])  # e.g. "2026-09-09"

    sgd_per_usd = d1["SGD"]
    rmb_per_usd = d1["CNY"]
    sgd_per_eur = d2["SGD"]

    return [sgd_per_usd, sgd_per_eur, rmb_per_usd], rate_date


def find_row_for_date(file_path, sheet_name, target_date):
    """Matches using the CACHED COMPUTED values of column A (older rows use
    formulas like '=A3928+1', not plain dates — data_only=True reads what
    Excel/Numbers last calculated, rather than the formula text itself)."""
    wb_values = openpyxl.load_workbook(file_path, data_only=True)
    ws_values = wb_values[sheet_name]

    last_filled_row = 1
    for row in range(2, ws_values.max_row + 1):
        cell_value = ws_values.cell(row=row, column=DATE_COL_IDX).value
        if cell_value is None:
            continue
        last_filled_row = row
        cell_date = cell_value.date() if isinstance(cell_value, datetime.datetime) else cell_value
        if cell_date == target_date:
            return row, True  # found an existing pre-filled row
    return last_filled_row + 1, False  # not found — append a new row


def scroll_to_row(ws, row):
    """Sets the sheet's saved view so it opens scrolled down to this row,
    with that row's date cell selected as the active cell."""
    cell_ref = f"A{row}"
    ws.sheet_view.selection = [Selection(activeCell=cell_ref, sqref=cell_ref)]
    ws.sheet_view.topLeftCell = f"A{max(row - 15, 1)}"


def main():
    print("Fetching rates...")
    rates, rate_date = fetch_rates()
    print(f"Rates for {rate_date} (SGD/USD, SGD/EUR, USD/RMB): {rates}")

    print("Finding matching row (using cached computed dates)...")
    row, found_existing = find_row_for_date(FILE_PATH, SHEET_NAME, rate_date)
    print(f"Target row: {row} (pre-filled: {found_existing})")

    print(f"Opening {FILE_PATH} to write (preserving formulas)...")
    wb = openpyxl.load_workbook(FILE_PATH)  # data_only=False (default) — keeps formulas intact
    ws = wb[SHEET_NAME]

    if not found_existing:
        date_cell = ws.cell(row=row, column=DATE_COL_IDX, value=datetime.datetime.combine(rate_date, datetime.time()))
        date_cell.number_format = DATE_NUMBER_FORMAT

    print(f"Writing rates to row {row}...")
    for col_idx, value in zip(RATE_COL_IDXS, rates):
        cell = ws.cell(row=row, column=col_idx, value=value)
        cell.number_format = RATE_NUMBER_FORMAT

    scroll_to_row(ws, row)

    wb.save(FILE_PATH)
    print("Done.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)

"""
Daily FX rate updater — GitHub repo version.

Fetches SGD/USD, SGD/EUR, USD/RMB, SGD/GBP and appends them as a new row at
the bottom of the xlsx file, which lives in this same repo. Also copies down
column E's formula (a derived RMB calculation, not a fetched rate) from the
row above, so it doesn't go blank on newly-added rows. GitHub Actions commits
the change back automatically after this script runs. No cloud auth needed.

Conventions (confirmed against the tracker's historical data):
  - SGD/USD column = amount of SGD for 1 USD  (e.g. ~1.27)
  - SGD/EUR column = amount of SGD for 1 EUR  (e.g. ~1.47)
  - USD/RMB column = amount of RMB for 1 USD  (e.g. ~7.1)
  - SGD/GBP column = amount of SGD for 1 GBP  (e.g. ~1.70)
  - Column E = a formula referencing column D, not a fetched value; copied
    down (not recalculated) for each new row.
"""

import sys
import copy
import datetime
import re
import requests
import openpyxl
from openpyxl.worksheet.views import Selection

FILE_PATH = "SGDUSD_Tracker.xlsx"
SHEET_NAME = "new"
DATE_COL_IDX = 1
RATE_COL_IDXS = [2, 3, 4, 6]  # B, C, D, F = SGD/USD, SGD/EUR, USD/RMB, SGD/GBP
FORMULA_COL_IDX = 5  # E — copied down as a formula, not a fetched value


def fetch_rates():
    r1 = requests.get("https://api.frankfurter.dev/v1/latest", params={"base": "USD", "symbols": "SGD,CNY"})
    r1.raise_for_status()
    j1 = r1.json()
    d1 = j1["rates"]

    r2 = requests.get("https://api.frankfurter.dev/v1/latest", params={"base": "EUR", "symbols": "SGD"})
    r2.raise_for_status()
    d2 = r2.json()["rates"]

    r3 = requests.get("https://api.frankfurter.dev/v1/latest", params={"base": "GBP", "symbols": "SGD"})
    r3.raise_for_status()
    d3 = r3.json()["rates"]

    rate_date = datetime.date.fromisoformat(j1["date"])

    sgd_per_usd = d1["SGD"]
    rmb_per_usd = d1["CNY"]
    sgd_per_eur = d2["SGD"]
    sgd_per_gbp = d3["SGD"]

    return [sgd_per_usd, sgd_per_eur, rmb_per_usd, sgd_per_gbp], rate_date


def find_row_for_date(file_path, sheet_name, target_date):
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
            return row, True
    return last_filled_row + 1, False


def scroll_to_row(ws, row):
    cell_ref = f"A{row}"
    ws.sheet_view.selection = [Selection(activeCell=cell_ref, sqref=cell_ref)]
    ws.sheet_view.topLeftCell = f"A{max(row - 15, 1)}"


TEMPLATE_ROW = 2


def copy_style(source_cell, target_cell, fallback_cell=None):
    style_source = source_cell
    if fallback_cell is not None and source_cell.number_format == "General":
        style_source = fallback_cell

    target_cell.font = copy.copy(style_source.font)
    target_cell.alignment = copy.copy(style_source.alignment)
    target_cell.number_format = style_source.number_format
    target_cell.border = copy.copy(style_source.border)
    target_cell.fill = copy.copy(style_source.fill)


def copy_formula_down(ws, row, col_idx):
    """Copies the formula from the cell above into this cell, rewriting any
    row references it contains to point at this row instead (e.g. a formula
    referencing D3930 on row 3930 becomes a reference to D3931 on row 3931)."""
    source_cell = ws.cell(row=row - 1, column=col_idx)
    if not isinstance(source_cell.value, str) or not source_cell.value.startswith("="):
        return  # nothing to copy — source row has no formula here
    prev_row = row - 1
    new_formula = re.sub(rf"([A-Z]+){prev_row}\b", rf"\g<1>{row}", source_cell.value)
    target_cell = ws.cell(row=row, column=col_idx, value=new_formula)
    copy_style(source_cell, target_cell)


def main():
    print("Fetching rates...")
    rates, rate_date = fetch_rates()
    print(f"Rates for {rate_date} (SGD/USD, SGD/EUR, USD/RMB, SGD/GBP): {rates}")

    print("Finding matching row (using cached computed dates)...")
    row, found_existing = find_row_for_date(FILE_PATH, SHEET_NAME, rate_date)
    print(f"Target row: {row} (pre-filled: {found_existing})")

    print(f"Opening {FILE_PATH} to write (preserving formulas)...")
    wb = openpyxl.load_workbook(FILE_PATH)
    ws = wb[SHEET_NAME]

    date_cell = ws.cell(row=row, column=DATE_COL_IDX, value=datetime.datetime.combine(rate_date, datetime.time()))
    style_source_row = row - 1 if row > 2 else TEMPLATE_ROW
    copy_style(ws.cell(row=style_source_row, column=DATE_COL_IDX), date_cell, ws.cell(row=TEMPLATE_ROW, column=DATE_COL_IDX))

    print(f"Writing rates to row {row}...")
    for col_idx, value in zip(RATE_COL_IDXS, rates):
        cell = ws.cell(row=row, column=col_idx, value=value)
        copy_style(ws.cell(row=row - 1, column=col_idx), cell, ws.cell(row=TEMPLATE_ROW, column=col_idx))

    if row > 2:
        print(f"Copying column E formula down to row {row}...")
        copy_formula_down(ws, row, FORMULA_COL_IDX)

    scroll_to_row(ws, row)

    wb.save(FILE_PATH)
    print("Done.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)

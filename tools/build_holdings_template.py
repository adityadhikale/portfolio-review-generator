"""Re-create templates/Holdings_Template.xlsx, the one holdings format the tool reads.

    python tools/build_holdings_template.py

The file holds no client data and is kept in git. Its layout must match crest/holdings.py.
"""
from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from crest.holdings import ASOF_CELL, ASOF_LABEL, CLIENT_CELL, CLIENT_LABEL, COLUMNS, FIRST_ROW, HEADER_ROW  # noqa: E402

OUT = ROOT / "templates" / "Holdings_Template.xlsx"
FONT = "Arial"
NAVY = PatternFill("solid", fgColor="1F3864")
INPUT = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LAST_ROW = FIRST_ROW + 499          # formatting and checks cover 500 holding rows

GUIDE = [
    ("How to fill the CREST holdings template", None),
    ("", None),
    ("Step", "What to do"),
    ("1", "Save a copy of this file and name it after the client."),
    ("2", f"Cell {CLIENT_CELL}: type the client's name exactly as it should appear on the deck."),
    ("3", f"Cell {ASOF_CELL}: type the date of the holdings statement as dd/mm/yyyy (or enter a real Excel date)."),
    ("4", f"From row {FIRST_ROW}, paste one holding per row. The holdings end at the first blank row."),
    ("5", "Upload the file (or put it in the input\\ folder and run python run.py)."),
    ("", None),
    ("Column", "Rule"),
    ("ISIN", "12 characters, e.g. INE002A01018. Give an ISIN or a Symbol (both is best)."),
    ("Symbol", "NSE symbol (e.g. RELIANCE) or BSE code (e.g. 540565)."),
    ("Company Name", "Optional. Used for matching only when neither ISIN nor Symbol is found."),
    ("Quantity", "A number above 0."),
    ("Average Price", "Average buy price per share, a number above 0, no currency symbol."),
    ("", None),
    ("Good to know", ""),
    ("Prices", "There is no price column: prices come live from the CREST Universal Sheet."),
    ("Duplicates", "Rows for the same stock are merged at the weighted average price."),
    ("Errors", "If anything is wrong the run stops and names the exact cell to fix."),
    ("Layout", f"Keep the labels in {CLIENT_LABEL[0]}:{ASOF_LABEL[0]} and the headers on row {HEADER_ROW} with these names."),
]


def build(path: Path = OUT) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Holdings"
    bold = Font(name=FONT, size=10, bold=True)
    plain = Font(name=FONT, size=10)

    ws["B1"] = "CREST holdings template"
    ws["B1"].font = Font(name=FONT, size=14, bold=True, color="1F3864")
    for (lab_cell, lab), val_cell in ((CLIENT_LABEL, CLIENT_CELL), (ASOF_LABEL, ASOF_CELL)):
        ws[lab_cell] = lab
        ws[lab_cell].font = bold
        ws[val_cell].fill = INPUT
        ws[val_cell].border = BOX
        ws[val_cell].font = plain
    ws[ASOF_CELL].number_format = "dd/mm/yyyy"

    for j, label in enumerate(COLUMNS.values(), start=2):
        c = ws.cell(HEADER_ROW, j, label)
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        c.fill = NAVY
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BOX
    for r in range(FIRST_ROW, LAST_ROW + 1):
        for j in range(2, 2 + len(COLUMNS)):
            ws.cell(r, j).font = plain
        ws.cell(r, 5).number_format = "#,##0.####"
        ws.cell(r, 6).number_format = "#,##0.00"
    for letter, w in zip("ABCDEF", (2, 18, 16, 40, 14, 16)):
        ws.column_dimensions[letter].width = w

    pos = DataValidation(type="decimal", operator="greaterThan", formula1="0", allow_blank=True,
                         showErrorMessage=True, errorTitle="Number above 0", error="Enter a number above 0.")
    pos.add(f"E{FIRST_ROW}:F{LAST_ROW}")
    ws.add_data_validation(pos)
    ws.freeze_panes = f"A{FIRST_ROW}"

    g = wb.create_sheet("How to fill")
    for i, (a, b) in enumerate(GUIDE, start=1):
        g.cell(i, 1, a).font = plain
        if b is not None:
            g.cell(i, 2, b).font = plain
            g.cell(i, 2).alignment = Alignment(wrap_text=True, vertical="top")
        if a in ("Step", "Column", "Good to know"):
            g.cell(i, 1).font = g.cell(i, 2).font = bold
    g["A1"].font = Font(name=FONT, size=14, bold=True, color="1F3864")
    g.column_dimensions["A"].width = 16
    g.column_dimensions["B"].width = 95

    wb.active = 0
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


if __name__ == "__main__":
    print(f"Written {build()}")

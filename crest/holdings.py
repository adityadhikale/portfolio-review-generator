"""Standard holdings template reader.

Every client's holdings are copied into templates/Holdings_Template.xlsx before a run:

    B2 "Client Name"   C2 <client name>
    B3 "As On Date"    C3 <dd/mm/yyyy, or a real Excel date>
    row 5              ISIN | Symbol | Company Name | Quantity | Average Price
    row 6 onwards      one holding per row, ending at the first blank row

Any problem stops the run with the exact cell to fix. Prices are never read from the template: they always come
from the Universal Sheet.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import openpyxl

from .util import PipelineError, normalize_symbol, to_number

CLIENT_LABEL, CLIENT_CELL = ("B2", "Client Name"), "C2"
ASOF_LABEL, ASOF_CELL = ("B3", "As On Date"), "C3"
HEADER_ROW = 5
FIRST_ROW = HEADER_ROW + 1
COLUMNS = {"isin": "ISIN", "symbol": "Symbol", "name": "Company Name", "quantity": "Quantity", "avg_price": "Average Price"}
ISIN_RE = re.compile(r"^[A-Z]{2}[A-Z0-9]{9}[0-9]$")


@dataclass
class Holding:
    symbol: str            # normalised symbol (the ISIN when no symbol is given): the key used everywhere downstream
    raw_symbol: str        # as typed (symbol, else ISIN)
    quantity: float
    avg_price: float       # weighted average after merging duplicates
    name: str | None = None
    source_rows: list[int] = field(default_factory=list)  # 1-based rows in the template
    isin: str | None = None


@dataclass
class HoldingsFile:
    client: str
    asof: str                          # dd/mm/yyyy
    sheet: str
    raw_rows: list[list]               # the sheet, cell values, unchanged
    holdings: list[Holding]
    notes: list[str]


def _norm(v) -> str:
    return re.sub(r"\s+", " ", str(v or "")).strip().lower()


def _blank(v) -> bool:
    return v is None or str(v).strip() == ""


def _asof(v, errors: list[str]) -> str | None:
    if isinstance(v, datetime):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, date):
        return v.strftime("%d/%m/%Y")
    if _blank(v):
        errors.append(f"Cell {ASOF_CELL}: the as-on date is blank. Type the date of the holdings (dd/mm/yyyy).")
        return None
    s = str(v).strip()
    try:
        return datetime.strptime(s, "%d/%m/%Y").strftime("%d/%m/%Y")
    except ValueError:
        errors.append(f"Cell {ASOF_CELL}: '{s}' is not a date. Type it as dd/mm/yyyy, e.g. 08/10/2026.")
        return None


def _pick_sheet(wb):
    for ws in wb.worksheets:
        if _norm(ws[CLIENT_LABEL[0]].value) == CLIENT_LABEL[1].lower():
            return ws
    raise PipelineError(
        f"This is not a CREST holdings template: no sheet has '{CLIENT_LABEL[1]}' in cell {CLIENT_LABEL[0]}. "
        "Copy the holdings into templates\\Holdings_Template.xlsx and upload that file.")


def read_template(path: Path) -> HoldingsFile:
    path = Path(path)
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise PipelineError(f"{path.name}: the holdings file must be the CREST holdings template (.xlsx).")
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:  # noqa: BLE001 - corrupt / not really an Excel file
        raise PipelineError(f"{path.name} could not be opened as an Excel file ({e}).") from e
    ws = _pick_sheet(wb)
    errors: list[str] = []
    notes: list[str] = []

    if _norm(ws[ASOF_LABEL[0]].value) != ASOF_LABEL[1].lower():
        errors.append(f"Cell {ASOF_LABEL[0]}: should read '{ASOF_LABEL[1]}' (found '{ws[ASOF_LABEL[0]].value or ''}').")
    client = str(ws[CLIENT_CELL].value or "").strip()
    if not client:
        errors.append(f"Cell {CLIENT_CELL}: the client name is blank. Type the client's name.")
    asof = _asof(ws[ASOF_CELL].value, errors)

    header = {_norm(c.value): c.column for c in ws[HEADER_ROW] if not _blank(c.value)}
    col: dict[str, int] = {}
    for key, label in COLUMNS.items():
        if label.lower() in header:
            col[key] = header[label.lower()]
        else:
            errors.append(f"Row {HEADER_ROW}: column '{label}' is missing. The headers must be "
                          f"{' | '.join(COLUMNS.values())}.")
    if len(col) < len(COLUMNS):
        raise PipelineError(f"{path.name} (sheet '{ws.title}'): " + " ".join(errors))

    L = openpyxl.utils.get_column_letter
    merged: list[Holding] = []
    by_isin: dict[str, Holding] = {}
    by_sym: dict[str, Holding] = {}
    r = FIRST_ROW
    last = ws.max_row
    while r <= last:
        get = lambda k: ws.cell(r, col[k]).value  # noqa: E731
        vals = {k: get(k) for k in COLUMNS}
        if all(_blank(v) for v in vals.values()):
            break
        cell = lambda k: f"{L(col[k])}{r}"  # noqa: E731
        isin = str(vals["isin"]).strip().upper() if not _blank(vals["isin"]) else None
        sym = str(vals["symbol"]).strip() if not _blank(vals["symbol"]) else None
        if isinstance(vals["symbol"], float) and vals["symbol"].is_integer():
            sym = str(int(vals["symbol"]))  # BSE code typed as a number
        name = str(vals["name"]).strip() if not _blank(vals["name"]) else None
        row_ok = True
        if not isin and not sym:
            errors.append(f"Row {r}: give an ISIN ({cell('isin')}) or a Symbol ({cell('symbol')}).")
            row_ok = False
        if isin and not ISIN_RE.match(isin):
            errors.append(f"Cell {cell('isin')}: '{isin}' is not an ISIN (12 characters, e.g. INE002A01018).")
            row_ok = False
        nums = {}
        for k in ("quantity", "avg_price"):
            v = vals[k]
            n = None if isinstance(v, bool) else to_number(v)
            if n is None or n <= 0:
                errors.append(f"Cell {cell(k)}: {COLUMNS[k]} must be a number above 0 "
                              f"(found {'a blank' if _blank(v) else repr(v)}).")
                row_ok = False
            nums[k] = n
        if row_ok:
            h = Holding(normalize_symbol(sym or isin), sym or isin, nums["quantity"], nums["avg_price"], name, [r], isin)
            m = (by_isin.get(isin) if isin else None) or by_sym.get(h.symbol)
            if m is None:
                merged.append(h)
                m = h
            else:
                tot = m.quantity + h.quantity
                m.avg_price = (m.quantity * m.avg_price + h.quantity * h.avg_price) / tot
                m.quantity = tot
                m.source_rows.append(r)
                m.isin = m.isin or isin
                m.name = m.name or name
                notes.append(f"Rows {m.source_rows} are the same stock ({m.raw_symbol}): merged at the weighted "
                             "average price")
            if isin:
                by_isin[isin] = m
            by_sym[h.symbol] = m
        r += 1

    after = [i for i in range(r + 1, last + 1)
             if any(not _blank(ws.cell(i, c).value) for c in col.values())]
    if after:
        notes.append(f"WARN: row {r} is blank, so the holdings end there; rows {after[0]}-{after[-1]} below it are "
                     "NOT in this review. Delete the blank row if they are holdings.")
    if not merged and not errors:
        errors.append(f"Row {FIRST_ROW}: no holdings. Enter one holding per row from row {FIRST_ROW}.")
    if errors:
        shown = errors[:20] + ([f"... and {len(errors) - 20} more"] if len(errors) > 20 else [])
        raise PipelineError(f"{path.name} (sheet '{ws.title}') needs fixing: " + " ".join(shown))

    raw_rows = [list(row) for row in ws.iter_rows(values_only=True)]
    wb.close()
    return HoldingsFile(client, asof, ws.title, raw_rows, merged, notes)

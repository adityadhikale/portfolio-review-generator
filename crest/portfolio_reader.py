"""Read holdings back from a Portfolio Excel (for Step B, the deck).

The Portfolio sheet is found by the hidden marker in A1 ("CREST Portfolio v1"), and its columns by header name, so
a copy edited by hand still reads correctly: rows deleted or added, columns moved, or the file downloaded again from
Google Drive. Quantity and Average Price typed in the Portfolio sheet win. A row with no Key is matched by its
Ticker, then by its Name. Prices, sector, views, results and ratings are never read from here: Step B takes them
fresh from the Universal Sheet.
"""
from __future__ import annotations

import re
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

from .excel_build import MARKER, MARKER_CELL
from .holdings import HoldingsFile, Holding, Merger, parse_date_cell, positive
from .util import PipelineError, normalize_symbol

NEEDED = {"name": "Name", "quantity": "Quantity Available", "avg_price": "Average Price"}
OPTIONAL = {"key": "Key", "ticker": "Ticker"}


def _norm(v) -> str:
    return re.sub(r"\s+", " ", str(v or "")).strip().lower()


def _blank(v) -> bool:
    return v is None or str(v).strip() == ""


def _has_marker(ws) -> bool:
    return str(ws[MARKER_CELL].value or "").strip() == MARKER


def _header_row(ws) -> int | None:
    want = {v.lower() for v in NEEDED.values()}
    for r in range(1, min(ws.max_row, 30) + 1):
        if want <= {_norm(c.value) for c in ws[r]}:
            return r
    return None


def file_kind(path: Path) -> str:
    """'portfolio' (a Portfolio Excel built by Step A) or 'template' (a holdings template)."""
    path = Path(path)
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise PipelineError(f"{path.name}: upload a holdings template or a Portfolio Excel (.xlsx).")
    try:
        wb = openpyxl.load_workbook(path)
    except Exception as e:  # noqa: BLE001 - corrupt / not really an Excel file
        raise PipelineError(f"{path.name} could not be opened as an Excel file ({e}).") from e
    try:
        for ws in wb.worksheets:
            if _has_marker(ws):
                return "portfolio"
        if "Portfolio" in wb.sheetnames and _header_row(wb["Portfolio"]):
            return "portfolio"          # built before the marker existed
        return "template"
    finally:
        wb.close()


def _portfolio_sheet(wb):
    for ws in wb.worksheets:
        if _has_marker(ws):
            return ws
    if "Portfolio" in wb.sheetnames:
        return wb["Portfolio"]
    raise PipelineError("No Portfolio sheet found. Upload the Portfolio Excel the tool built (or a holdings template).")


def _labelled(ws, label: str, above: int):
    """Value to the right of the cell reading `label` in the summary block above the table, and its cell name."""
    for r in range(1, above):
        for c in ws[r]:
            if _norm(c.value) == label.lower():
                v = ws.cell(r, c.column + 1)
                return v.value, v.coordinate
    return None, None


def read_portfolio_excel(path: Path) -> HoldingsFile:
    path = Path(path)
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = _portfolio_sheet(wb)
    hr = _header_row(ws)
    if hr is None:
        raise PipelineError(f"{path.name} (sheet '{ws.title}'): no header row with "
                            f"{', '.join(repr(v) for v in NEEDED.values())}. Keep those column headers.")
    header = {_norm(c.value): c.column for c in ws[hr] if not _blank(c.value)}
    col = {k: header[v.lower()] for k, v in {**NEEDED, **OPTIONAL}.items() if v.lower() in header}
    errors: list[str] = []
    notes: list[str] = []

    cv, ccell = _labelled(ws, "Client", hr)
    client = str(cv or "").strip()
    if not client:
        errors.append(f"Cell {ccell or 'F3'}: the client name is blank (label 'Client' in the summary block).")
    av, acell = _labelled(ws, "Data as of", hr)
    asof = parse_date_cell(av, acell or "F4", errors)

    get = lambda r, k: ws.cell(r, col[k]).value if k in col else None  # noqa: E731
    cell = lambda r, k: f"{get_column_letter(col[k])}{r}"  # noqa: E731
    merger = Merger(notes)
    by_ticker = by_name = 0
    for r in range(hr + 1, ws.max_row + 1):
        key, name, ticker = get(r, "key"), get(r, "name"), get(r, "ticker")
        qty, avg = get(r, "quantity"), get(r, "avg_price")
        if _norm(name) == "total" or all(_blank(v) for v in (key, name, qty, avg)):
            break
        key = "" if _blank(key) else str(key).strip()
        ticker = "" if _blank(ticker) or str(ticker).startswith("=") else str(ticker).strip()
        name = "" if _blank(name) else str(name).strip()
        if not (key or ticker or name):
            errors.append(f"Row {r}: no Key, Ticker or Name - say which stock this row is.")
            continue
        q = positive(qty, cell(r, "quantity"), NEEDED["quantity"], errors)
        a = positive(avg, cell(r, "avg_price"), NEEDED["avg_price"], errors)
        if q is None or a is None:
            continue
        if key:
            sym, raw = normalize_symbol(key), key
        elif ticker:
            sym, raw = normalize_symbol(ticker), ticker
            by_ticker += 1
        else:
            sym, raw = normalize_symbol(name), name
            by_name += 1
        merger.add(Holding(sym, raw, q, a, name or None, [r]))

    if not merger.holdings and not errors:
        errors.append(f"Sheet '{ws.title}' has no holdings below the header row {hr}.")
    if errors:
        shown = errors[:20] + ([f"... and {len(errors) - 20} more"] if len(errors) > 20 else [])
        raise PipelineError(f"{path.name} (sheet '{ws.title}') needs fixing: " + " ".join(shown))
    if by_ticker or by_name:
        notes.append(f"{by_ticker + by_name} row(s) in the Portfolio sheet have no Key (added by hand): matched by "
                     f"Ticker ({by_ticker}) or Name ({by_name})")

    raw_rows = [list(row) for row in ws.iter_rows(values_only=True)]
    step_a = _step_a_holdings(wb)
    wb.close()
    return HoldingsFile(client, asof, ws.title, raw_rows, merger.holdings, notes, step_a)


def _step_a_holdings(wb) -> dict[str, str] | None:
    """Key -> company name for every holding Step A had: Holdings_Normalized (Key, Company Name), with the Universal
    name from Universal_Snapshot where there is one. None when the workbook has no Holdings_Normalized sheet."""
    if "Holdings_Normalized" not in wb.sheetnames:
        return None
    out = {}
    hn = wb["Holdings_Normalized"]
    for r in range(2, hn.max_row + 1):
        k = hn.cell(r, 1).value
        if not _blank(k):
            out[normalize_symbol(k)] = str(hn.cell(r, 4).value or "").strip()
    if "Universal_Snapshot" in wb.sheetnames:
        us = wb["Universal_Snapshot"]
        for r in range(2, us.max_row + 1):
            k, n = us.cell(r, 1).value, us.cell(r, 2).value
            if not _blank(k) and normalize_symbol(k) in out and not _blank(n):
                out[normalize_symbol(k)] = str(n).strip()
    return out

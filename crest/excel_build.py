"""Build '[Client] Portfolio.xlsx' with openpyxl. Every calculated cell is a live formula.

Sheets
  Portfolio            summary block (B2:C6) + holdings table (header row 8, data from row 9), as in reference_portfolio.xlsx
  Raw_Holdings         the holdings template as uploaded, cell-for-cell
  Holdings_Normalized  the holdings the pipeline used (duplicate rows merged)
  Universal_Snapshot   the Universal Sheet rows matched to this client (no other stocks)
  Exceptions           everything that needs a human look
  References           AI-written Sector / Result Line / Commentary with source links, and AI Rating / Sentiment
                       with the AI's reason
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.indexed_list import IndexedList

from .universal import value_number

FONT = "Arial"
INR = '#,##0.00;[Red]-#,##0.00'  # amounts: no currency symbol
PCT = '0.00%;[Red]-0.00%'
QTY = '#,##0.##'
MCAP = '#,##0.00'
HDR_FILL = PatternFill("solid", fgColor="0D0D0D")
SUB_FILL = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="D9D9D9")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FIRST_DATA_ROW = 9
MARKER_CELL, MARKER = "A1", "CREST Portfolio v1"  # hidden (column A): tells Step B this is a Portfolio Excel
HEADER_ROW = 8
PORTFOLIO_COLS = ["Key", "Name", "Ticker", "Sector", "Market Cap (Cr)", "Quantity Available", "Average Price",
                  "Invested Value", "CMP", "Current Value", "Unrealized P&L", "Unrealized P&L (%)",
                  "% Allocation", "Views"]
SNAP_COLS = ["Key", "Name", "Ticker (CREST)", "Sector", "Market Cap (Cr)", "CMP", "Views", "Short Name",
             "Deck Group", "Result Line", "Rating", "Sentiment", "Management Commentary", "Is ETF",
             "Universal Row", "Ticker (as in Universal)", "Rating Source"]
NORM_COLS = ["Key", "Symbol / ISIN (as typed)", "ISIN", "Company Name", "Quantity", "Average Price",
             "Source Rows (Raw_Holdings)"]


@dataclass
class ExcelLayout:
    path: Path
    first_row: int
    last_row: int
    total_row: int
    keys_in_order: list


def _hdr(ws, row, cols, start_col=1):
    for j, name in enumerate(cols, start=start_col):
        c = ws.cell(row=row, column=j, value=name)
        c.font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX


def _font_all(ws, size=10):
    """Arial everywhere. Unstyled cells already use the workbook default (Arial 10); only cells given another
    font explicitly are changed, reusing one Font object per (bold, italic, colour)."""
    cache = {}
    for row in ws.iter_rows():
        for c in row:
            if not c.has_style:
                continue
            f = c.font
            if f is not None and f.name == FONT:
                continue
            b, i = (bool(f.b), bool(f.i)) if f is not None else (False, False)
            color = f.color if f is not None else None
            key = (b, i, getattr(color, "rgb", None), getattr(color, "theme", None))
            if key not in cache:
                cache[key] = Font(name=FONT, size=size, bold=b, italic=i, color=color)
            c.font = cache[key]


def build_workbook(path: Path, *, client: str, asof: str, period_label: str, parse, priced: list, snapshot_rows: list, exception_rows: list, deck_group_of,
                   google_finance: bool = True) -> ExcelLayout:
    """priced: list of Matched sorted by current value desc (only rows with a price).
    snapshot_rows: list of Matched (all matched, incl. those without a price)."""
    wb = Workbook()
    # Workbook default font = Arial 10, so plain cells need no per-cell font (setting fonts cell by cell made
    # the build take ~20 s per workbook).
    wb._fonts = IndexedList([Font(name=FONT, size=10)])
    wb.calculation.fullCalcOnLoad = True
    ws = wb.active
    ws.title = "Portfolio"
    raw = wb.create_sheet("Raw_Holdings")
    norm = wb.create_sheet("Holdings_Normalized")
    snap = wb.create_sheet("Universal_Snapshot")
    exc = wb.create_sheet("Exceptions")

    # ---------------- Raw_Holdings: unchanged upload
    for r in parse.raw_rows:
        raw.append(list(r))
    for j in range(1, (max(len(r) for r in parse.raw_rows) if parse.raw_rows else 1) + 1):
        raw.column_dimensions[get_column_letter(j)].width = 16
    _font_all(raw)

    # ---------------- Holdings_Normalized
    _hdr(norm, 1, NORM_COLS)
    holdings = [m.holding for m in snapshot_rows] + [e.holding for e in exception_rows if getattr(e, "holding", None)]
    seen = set()
    nrow = 2
    for h in holdings:
        if h.symbol in seen:
            continue
        seen.add(h.symbol)
        norm.append([h.symbol, h.raw_symbol, h.isin, h.name, h.quantity, h.avg_price,
                     ", ".join(str(x) for x in h.source_rows)])
        nrow += 1
    n_last = nrow - 1
    for r in range(2, n_last + 1):
        norm.cell(r, 5).number_format = QTY
        norm.cell(r, 6).number_format = INR
    chk = n_last + 2
    norm.cell(chk, 4, "Qty x Avg Price total").font = Font(name=FONT, bold=True)
    norm.cell(chk, 6, f"=SUMPRODUCT(E2:E{n_last},F2:F{n_last})").number_format = INR
    norm.cell(1, 9, "The holdings template's rows after merging rows for the same stock (weighted average price).")
    for j, w in enumerate([14, 18, 16, 34, 12, 14, 22], start=1):
        norm.column_dimensions[get_column_letter(j)].width = w
    norm.freeze_panes = "A2"
    _font_all(norm)

    # ---------------- Universal_Snapshot (matched rows only)
    _hdr(snap, 1, SNAP_COLS)
    for m in snapshot_rows:
        u = m.uni
        mcap = value_number(u, "mcap")
        mcap_display = "ETF" if u.is_etf else (mcap if mcap is not None else "")
        snap.append([m.holding.symbol, u.values["name"], u.ticker_crest, u.values["sector"], mcap_display,
                     value_number(u, "cmp"), u.values["views"],
                     u.values.get("short_name"), deck_group_of(u), u.values["result"], u.rating or None,
                     u.sentiment or None, u.values["commentary"], "Yes" if u.is_etf else "No",
                     u.row, u.values["ticker"], getattr(u, "rating_source", "") or None])
    s_last = max(2, len(snapshot_rows) + 1)
    for r in range(2, s_last + 1):
        snap.cell(r, 5).number_format = MCAP
        snap.cell(r, 6).number_format = INR
        for c in (10, 13):
            snap.cell(r, c).alignment = Alignment(wrap_text=False)
    for j, w in enumerate([13, 34, 16, 22, 14, 11, 16, 24, 30, 40, 10, 10, 60, 7, 9, 16, 15], start=1):
        snap.column_dimensions[get_column_letter(j)].width = w
    snap.freeze_panes = "B2"
    _font_all(snap)

    # ---------------- Portfolio
    n = len(priced)
    first, last = FIRST_DATA_ROW, FIRST_DATA_ROW + max(n, 1) - 1
    total = last + 1
    ws[MARKER_CELL] = MARKER
    ws["B2"] = "Summary"
    ws["B2"].font = Font(name=FONT, bold=True, size=12)
    labels = [("Invested Value", f"=SUM(H{first}:H{last})", INR),
              ("Present Value", f"=SUM(J{first}:J{last})", INR),
              ("Unrealized P&L", "=C4-C3", INR),
              ("Unrealized P&L Pct.", "=IFERROR(C5/C3,0)", PCT)]
    for i, (lab, f, fmt) in enumerate(labels, start=3):
        a, b = ws.cell(i, 2, lab), ws.cell(i, 3, f)
        a.font = Font(name=FONT, bold=True)
        a.fill = SUB_FILL
        b.number_format = fmt
        a.border = b.border = BOX
    info = [("Client", client), ("Data as of", asof), ("Period", period_label),
            ("Price source", "Universal Sheet CMP")]
    for i, (lab, val) in enumerate(info, start=3):
        ws.cell(i, 5, lab).font = Font(name=FONT, bold=True, color="595959")
        ws.cell(i, 6, val).font = Font(name=FONT, color="595959")

    _hdr(ws, HEADER_ROW, PORTFOLIO_COLS)
    SN = "Universal_Snapshot"
    rng = lambda col: f"{SN}!${col}$2:${col}${s_last}"  # noqa: E731
    look = lambda col, r: f'INDEX({rng(col)},MATCH($A{r},{rng("A")},0))'  # noqa: E731
    keys = []
    for i, m in enumerate(priced):
        r = first + i
        keys.append(m.holding.symbol)
        u, h = m.uni, m.holding
        # Market Cap and CMP: Google-Sheets GOOGLEFINANCE formulas in the form Google itself exports
        # (__xludf.DUMMYFUNCTION). Opened in Google Sheets they become live GOOGLEFINANCE formulas; in Excel /
        # LibreOffice the function is unknown, so IFERROR shows the fallback value from this run's data.
        gref = f"C{r}" if u.ticker_exchange == "NSE" else f'SUBSTITUTE(C{r},""BSE:"",""BOM:"")'  # Google uses BOM: for BSE
        mcap = value_number(u, "mcap")
        mcap_fb = '"ETF"' if u.is_etf else (repr(round(mcap, 2)) if mcap is not None else '""')
        price = value_number(u, "cmp")
        row = [h.symbol,
               u.values["name"],
               f'=IFERROR({look("C", r)}&"","")',
               u.values["sector"],
               (f'=IFERROR(__xludf.DUMMYFUNCTION("GOOGLEFINANCE({gref},""marketcap"")/10^7"),{mcap_fb})'
                if google_finance else ("ETF" if u.is_etf else (round(mcap, 2) if mcap is not None else ""))),
               h.quantity,
               h.avg_price,
               f"=F{r}*G{r}",
               (f'=IFERROR(__xludf.DUMMYFUNCTION("GOOGLEFINANCE({gref},""price"")"),{price!r})'
                if google_finance else price),
               f"=F{r}*I{r}",
               f"=J{r}-H{r}",
               f"=K{r}/H{r}",
               f"=J{r}/SUM($J${first}:$J${last})",
               f'=IFERROR({look("G", r)}&"","")']
        for j, v in enumerate(row, start=1):
            c = ws.cell(r, j, v)
            c.font = Font(name=FONT, size=10)
            c.border = BOX
        for j, fmt in ((5, MCAP), (6, QTY), (7, INR), (8, INR), (9, INR), (10, INR), (11, INR), (12, PCT), (13, PCT)):
            ws.cell(r, j).number_format = fmt
        ws.cell(r, 5).alignment = Alignment(horizontal="right")
    ws.cell(total, 2, "Total").font = Font(name=FONT, bold=True)
    for col, fmt in (("H", INR), ("J", INR), ("K", INR), ("M", PCT)):
        c = ws[f"{col}{total}"]
        c.value = f"=SUM({col}{first}:{col}{last})"
        c.number_format = fmt
        c.font = Font(name=FONT, bold=True)
    ws[f"L{total}"] = f"=IFERROR(K{total}/H{total},0)"
    ws[f"L{total}"].number_format = PCT
    ws[f"L{total}"].font = Font(name=FONT, bold=True)
    for j in range(2, 15):
        ws.cell(total, j).fill = SUB_FILL
        ws.cell(total, j).border = BOX
    widths = {"A": 13, "B": 40, "C": 17, "D": 24, "E": 15, "F": 11, "G": 13, "H": 15, "I": 12, "J": 15,
              "K": 15, "L": 12, "M": 11, "N": 22}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.column_dimensions["A"].hidden = True
    ws.row_dimensions[HEADER_ROW].height = 30
    ws.freeze_panes = f"C{first}"
    ws.auto_filter.ref = f"B{HEADER_ROW}:N{last}"
    ws[f"A{HEADER_ROW}"].comment = Comment("Hidden key column: normalised symbol (or ISIN) used by every INDEX/MATCH lookup.", "CREST")

    # ---------------- Exceptions
    _hdr(exc, 1, ["Severity", "Type", "Symbol", "Detail", "Quantity", "Average Price", "Invested (Qty x Avg)"])
    for i, e in enumerate(exception_rows, start=2):
        h = getattr(e, "holding", None)
        exc.cell(i, 1, getattr(e, "severity", "WARN"))
        exc.cell(i, 2, e.kind)
        exc.cell(i, 3, e.symbol)
        exc.cell(i, 4, e.detail)
        if h is not None:
            exc.cell(i, 5, h.quantity).number_format = QTY
            exc.cell(i, 6, h.avg_price).number_format = INR
            exc.cell(i, 7, f"=E{i}*F{i}").number_format = INR
    if not exception_rows:
        exc.cell(2, 1, "None")
    for j, w in enumerate([9, 18, 16, 90, 10, 13, 18], start=1):
        exc.column_dimensions[get_column_letter(j)].width = w
    exc.freeze_panes = "A2"
    _font_all(exc)

    _references(wb.create_sheet("References"), snapshot_rows)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return ExcelLayout(path, first, last, total, keys)


def _references(ref, matched):
    """References: every value the AI wrote for this client's stocks, with its sources (the same list the website
    shows under 'AI values to review'), then every AI Rating / Sentiment with the AI's reason."""
    from .ai_research import ai_rated, ai_sourced
    research, ratings = ai_sourced(matched), ai_rated(matched)
    bold = Font(name=FONT, bold=True, size=12)
    grey = Font(name=FONT, color="595959")
    link = Font(name=FONT, color="0563C1", underline="single")
    ref["A1"] = "AI-sourced values and their references"
    ref["A1"].font = bold
    ref["A2"] = ("Written by the AI, not typed by CREST. Check each against its source before the deck goes out. To "
                 "correct one, type over it in the Universal Sheet and clear that row's Data Source (or Rating Source).")
    ref["A2"].font = grey

    r = 4
    ref.cell(r, 1, "Sector, Result Line and Management Commentary").font = Font(name=FONT, bold=True)
    r += 1
    _hdr(ref, r, ["Symbol", "Company", "Field", "Value", "Quarter", "New this run", "Source 1", "Source 2", "Source 3"])
    if not research:
        r += 1
        ref.cell(r, 1, "None: no AI-written Sector, Result Line or Commentary in this portfolio.").font = grey
    for x in research:
        r += 1
        urls = [u.strip() for u in str(x["sources"]).split("|") if u.strip().lower().startswith(("http://", "https://"))]
        for j, v in enumerate([x["symbol"], x["name"], x["field"], x["value"], x["period"],
                               "Yes" if x["new"] else "No"], start=1):
            ref.cell(r, j, v).alignment = Alignment(vertical="top", wrap_text=j == 4)
        if not urls:
            ref.cell(r, 7, "none recorded").font = grey
        for j, u in enumerate(urls[:3], start=7):
            c = ref.cell(r, j, u)
            c.hyperlink, c.font = u, link
            c.alignment = Alignment(vertical="top")

    r += 3
    ref.cell(r, 1, "Rating and Sentiment").font = Font(name=FONT, bold=True)
    ref.cell(r, 4, "Judged by the AI from the stock's Result Line and Management Commentary (above, or in "
                   "Universal_Snapshot); no web sources.").font = grey
    r += 1
    _hdr(ref, r, ["Symbol", "Company", "Rating · Sentiment", "Reason"])
    if not ratings:
        r += 1
        ref.cell(r, 1, "None: no AI-given Rating or Sentiment in this portfolio.").font = grey
    for x in ratings:
        r += 1
        rs = " · ".join(v for v in (x["rating"], x["sentiment"]) if v)
        for j, v in enumerate([x["symbol"], x["name"], rs, x["reason"] or "none recorded"], start=1):
            ref.cell(r, j, v).alignment = Alignment(vertical="top", wrap_text=j == 4)

    for j, w in enumerate([13, 30, 22, 60, 10, 12, 40, 40, 40], start=1):
        ref.column_dimensions[get_column_letter(j)].width = w
    _font_all(ref)

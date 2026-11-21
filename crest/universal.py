"""Universal Sheet loader and schema validation.

The rows come from Google Sheets or a local .xlsx (crest/sheet_source.py); the parsing and checks here are the same
for both. Schema (sheet 'Universal' or the first sheet; the header row is auto-detected in the first 10 rows):
  Required : Name, Ticker, Sector, Market Cap (Cr), CMP, Views
  Quarterly: Result Line, Rating, Sentiment, Management Commentary
  Optional : Symbol, Short Name, Exchange, Instrument Type, Deck Group
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

from .util import PipelineError, normalize_name, normalize_symbol, parse_ticker, to_number

SCHEMA = {
    # std key: (display header, aliases, required)
    "name":       ("Name", ["Name", "Company", "Company Name", "Stock Name"], True),
    "ticker":     ("Ticker", ["Ticker", "Ticker Symbol"], True),
    "sector":     ("Sector", ["Sector", "Industry"], True),
    "mcap":       ("Market Cap (Cr)", ["Market Cap (Cr)", "Market Cap", "MCap (Cr)", "Mcap", "Market Capitalisation (Cr)"], True),
    "cmp":        ("CMP", ["CMP", "Price", "Current Price", "Current Market Price"], True),
    "views":      ("Views", ["Views", "View", "Current View", "CREST View"], True),
    "result":     ("Result Line", ["Result Line", "Result Summary", "Q1 Result Summary", "Q2 Result Summary",
                                   "Q3 Result Summary", "Q4 Result Summary", "Result", "Q Result", "Quarterly Result"], True),
    "rating":     ("Rating", ["Rating", "Result Rating"], True),
    "sentiment":  ("Sentiment", ["Sentiment"], True),
    "commentary": ("Management Commentary", ["Management Commentary", "Commentary", "Mgmt Commentary"], True),
    "symbol":     ("Symbol", ["Symbol", "NSE Symbol", "Trading Symbol", "Broker Symbol"], False),
    "short_name": ("Short Name", ["Short Name", "Display Name"], False),
    "exchange":   ("Exchange", ["Exchange"], False),
    "itype":      ("Instrument Type", ["Instrument Type", "Type", "Instrument"], False),
    "deck_group": ("Deck Group", ["Deck Group", "Deck Sector", "Group"], False),
    "result_period": ("Q Result Period", ["Q Result Period", "Result Period", "Quarter"], False),
    "isin":       ("ISIN", ["ISIN", "ISIN Code", "ISIN No"], False),
    "alt_symbols": ("Alt Symbols", ["Alt Symbols", "Alternate Symbols", "Former Symbols"], False),
    # Written by the tool when Gemini fills a blank Rating / Sentiment (crest/shared.py, through the Sheets API). "AI · model ·
    # period · basis": the rating is reused while the result line and commentary are unchanged. A Rating typed by
    # CREST has this cell blank and always wins.
    "rating_source": ("Rating Source", ["Rating Source"], False),
    "rating_reason": ("Rating Reason", ["Rating Reason"], False),
}
QUARTERLY = ("result", "rating", "sentiment", "commentary")
RATINGS = ["Excellent", "Good", "Average", "Weak"]
SENTIMENTS = ["Bullish", "Neutral", "Bearish"]


@dataclass
class UniRow:
    row: int                 # 1-based row in the Universal Sheet
    values: dict             # std key -> raw value
    original: list           # full original row (all columns)
    key: str                 # normalised trading symbol used for matching
    ticker_exchange: str
    ticker_code: str
    ticker_crest: str        # "NSE:SYM" / "BSE:540565"
    ticker_defaulted: bool
    is_etf: bool
    etf_inferred: bool
    rating: str              # validated or ""
    sentiment: str           # validated or ""
    notes: list = field(default_factory=list)


@dataclass
class Universe:
    source: str              # where the rows came from (Google Sheet, snapshot or local file)
    sheet: str
    header_row: int
    headers: list
    col_index: dict
    rows: list[UniRow]
    by_key: dict
    by_code: dict
    by_name: dict
    duplicates: dict
    warnings: list
    by_isin: dict
    origin: object = None    # the UniversalSheet it was read from (for writing AI results back), if any


def _norm(h) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(h or "").lower()).strip()


def _canon(value, allowed):
    if value is None or not str(value).strip():
        return ""
    s = re.sub(r"[^A-Za-z]+", "", str(value)).lower()
    for a in allowed:
        if s == a.lower():
            return a
    return None  # present but not valid


_ROWS_CACHE: dict = {}


def read_xlsx_rows(path: Path):
    """(sheet name, rows as lists) of a local .xlsx. Reading the 5,000-row sheet takes several seconds, so the rows are kept for
    the rest of this process (period detection, then every client) until the file changes on disk. Fresh
    UniRow objects are still built on every load, so one client's run never affects another's."""
    st = path.stat()
    key = (str(path.resolve()), st.st_mtime_ns, st.st_size)
    if key not in _ROWS_CACHE:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb["Universal"] if "Universal" in wb.sheetnames else wb.worksheets[0]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        title = ws.title
        wb.close()
        _ROWS_CACHE.clear()
        _ROWS_CACHE[key] = (title, rows)
    title, rows = _ROWS_CACHE[key]
    return title, [list(r) for r in rows]


def _sheet(src):
    """-> (UniversalSheet, SheetData) for a UniversalSheet, or a path to a local .xlsx."""
    from .sheet_source import LocalSheet, UniversalSheet
    us = src if isinstance(src, UniversalSheet) else UniversalSheet(LocalSheet(Path(src)))
    return us, us.data()


def result_periods(src) -> list:
    """Every 'Q Result Period' value in the sheet - a quick read for period detection (no full load)."""
    rows = _sheet(src)[1].rows
    aliases = {_norm(a) for a in SCHEMA["result_period"][1]}
    for i, r in enumerate(rows[:10]):
        nh = [_norm(c) for c in r]
        j = next((k for k, h in enumerate(nh) if h in aliases), None)
        if j is not None:
            return [r2[j] for r2 in rows[i + 1:] if j < len(r2)]
    return []


def load_universe(src) -> Universe:
    """src: a UniversalSheet (Google, with snapshot fallback) or a path to a local .xlsx."""
    us, d = _sheet(src)
    sheet, all_rows, label = d.title, [list(r) for r in d.rows], d.source

    best = (0, None)
    for i, r in enumerate(all_rows[:10]):
        hits = sum(1 for k, (_, al, _) in SCHEMA.items() if any(_norm(a) in {_norm(c) for c in r} for a in al))
        if hits > best[0]:
            best = (hits, i)
    if best[1] is None:
        raise PipelineError("Universal Sheet: could not find a header row (expected Name, Ticker, Sector, ...).")
    h_idx = best[1]
    headers = all_rows[h_idx]
    nh = [_norm(h) for h in headers]
    col = {}
    for k, (_, aliases, _) in SCHEMA.items():
        for a in aliases:
            if _norm(a) in nh:
                col[k] = nh.index(_norm(a))
                break
    missing_req = [SCHEMA[k][0] for k, v in SCHEMA.items() if v[2] and k not in col and k not in QUARTERLY]
    missing_q = [SCHEMA[k][0] for k in QUARTERLY if k not in col]
    if missing_req or missing_q:
        msg = []
        if missing_req:
            msg.append(f"missing required column(s): {', '.join(missing_req)}")
        if missing_q:
            msg.append(f"missing quarterly result column(s): {', '.join(missing_q)} "
                       "(Result Line / Rating / Sentiment / Management Commentary feed the results-table slides)")
        raise PipelineError(f"Universal Sheet ({label}) {'; '.join(msg)}. Header row {h_idx + 1}: {headers}. "
                            "See templates/Universal_Sheet_template.xlsx for the expected layout.")

    rows, warnings = [], []
    by_key, by_code, by_name, dups, by_isin = {}, {}, {}, {}, {}
    for i in range(h_idx + 1, len(all_rows)):
        r = all_rows[i]
        get = lambda k: r[col[k]] if k in col and col[k] < len(r) else None  # noqa: E731
        if get("name") in (None, "") and get("ticker") in (None, ""):
            continue
        vals = {k: get(k) for k in SCHEMA}
        ex, code, defaulted = parse_ticker(vals["ticker"], vals.get("exchange"))
        key = normalize_symbol(vals["symbol"]) if vals.get("symbol") else normalize_symbol(code)
        name = str(vals["name"] or "")
        itype = str(vals.get("itype") or "").strip().upper()
        if itype:
            is_etf, inferred = itype == "ETF", False
        elif str(vals.get("sector") or "").strip().upper() == "ETF":
            is_etf, inferred = True, False
        else:
            is_etf = bool(re.search(r"\bETF\b|BEES$", name, re.I) or re.search(r"BEES$", key))
            inferred = is_etf
        notes = []
        rating = _canon(vals["rating"], RATINGS)
        if rating is None:
            notes.append(f"Rating '{vals['rating']}' is not one of {RATINGS}")
            rating = ""
        sentiment = _canon(vals["sentiment"], SENTIMENTS)
        if sentiment is None:
            notes.append(f"Sentiment '{vals['sentiment']}' is not one of {SENTIMENTS}")
            sentiment = ""
        if str(vals["ticker"] or "").strip().upper().startswith("BOM:"):
            notes.append(f"Ticker '{vals['ticker']}' uses BOM: prefix - written as BSE:{code}")
        if defaulted:
            notes.append(f"Ticker '{vals['ticker']}' has no exchange - defaulted to {ex}")
        u = UniRow(i + 1, vals, r, key, ex, code, f"{ex}:{code}", defaulted, is_etf, inferred, rating, sentiment, notes)
        rows.append(u)
        if key in by_key:
            dups.setdefault(key, [by_key[key].row]).append(u.row)
        else:
            by_key[key] = u
        by_code.setdefault(normalize_symbol(code), u)
        for alt in alt_symbols(vals.get("alt_symbols")):
            by_code.setdefault(alt, u)
        isin = str(vals.get("isin") or "").strip().upper()
        if isin:
            by_isin.setdefault(isin, []).append(u)
        nn = normalize_name(name)
        if nn:
            by_name.setdefault(nn, []).append(u)
    for k, rr in dups.items():
        warnings.append(f"Universal Sheet has duplicate symbol {k} on rows {rr}; matches to it are refused (ambiguous)")
    return Universe(label, sheet, h_idx + 1, headers, col, rows, by_key, by_code, by_name, dups, warnings, by_isin, us)


def alt_symbols(raw) -> list:
    """'509496, ITDCEM' / 543748.0 -> ['509496', 'ITDCEM'] (BSE codes and former NSE symbols)."""
    if raw in (None, ""):
        return []
    if isinstance(raw, float) and raw.is_integer():
        raw = int(raw)
    out = []
    for part in re.split(r"[,;/|]", str(raw)):
        part = re.sub(r"^(\d+)\.0$", r"\1", part.strip())
        if part:
            out.append(normalize_symbol(part))
    return out


def value_number(u: UniRow, key: str):
    return to_number(u.values.get(key))

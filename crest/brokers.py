"""Broker adapter layer.

Each broker has a JSON file in config/brokers/ mapping its column headers to the standard
schema. Only symbol, quantity and average price are required; LTP and invested value are
read when present (LTP for price_source=broker_ltp, invested value for the tie-out).
Everything else in the broker file is ignored (recalculated downstream).
"""
from __future__ import annotations

import csv
import re
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

from .util import CONFIG_DIR, PipelineError, load_json, normalize_symbol, to_number


@dataclass
class Holding:
    symbol: str            # normalised
    raw_symbol: str
    quantity: float
    avg_price: float       # weighted average after merging duplicates
    ltp: float | None = None
    invested_value_broker: float | None = None
    name: str | None = None
    exchange: str | None = None
    source_rows: list[int] = field(default_factory=list)  # 1-based rows in the upload
    isin: str | None = None


@dataclass
class BrokerParse:
    broker: str
    header_row: int                    # 1-based
    column_map: dict[str, str]         # std field -> header text found
    raw_rows: list[list]               # the upload, cell values, unchanged
    holdings: list[Holding]
    notes: list[str]
    dropped: list[dict]                # rows dropped (footer, zero qty, ...)
    asof_in_file: str | None


def _read_rows(path: Path) -> list[list]:
    suf = path.suffix.lower()
    if suf in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.worksheets[0]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
        return rows
    if suf == ".csv":
        for enc in ("utf-8-sig", "cp1252"):
            try:
                with open(path, newline="", encoding=enc) as fh:
                    return [list(r) for r in csv.reader(fh)]
            except UnicodeDecodeError:
                continue
    raise PipelineError(f"Unsupported holdings file type: {path.name} (use .xlsx or .csv)")


def _norm_header(h) -> str:
    return re.sub(r"[^a-z0-9%&]+", " ", str(h or "").lower()).strip()


def load_broker_configs() -> list[dict]:
    return [load_json(p) for p in sorted((CONFIG_DIR / "brokers").glob("*.json"))]


def map_header(row, cfg) -> tuple[dict, dict]:
    """Map one candidate header row to standard fields: exact alias match first, then keyword
    containment (with exclusions, e.g. 'long term qty' is not the quantity). Each column is used once."""
    norm = [_norm_header(c) for c in row]
    used, idx, names = set(), {}, {}
    for field_name, aliases in cfg["columns"].items():           # pass 1: exact header text
        for a in aliases:
            na = _norm_header(a)
            if na in norm and norm.index(na) not in used:
                j = norm.index(na)
                idx[field_name], names[field_name] = j, str(row[j])
                used.add(j)
                break
    for field_name, kws in cfg.get("keywords", {}).items():      # pass 2: header contains a keyword
        if field_name in idx:
            continue
        excl = cfg.get("exclude", {}).get(field_name, [])
        for kw in kws:
            hit = next((j for j, h in enumerate(norm) if j not in used and h and re.search(rf"\b{re.escape(kw)}\b", h)
                        and not any(re.search(rf"\b{re.escape(x)}\b", h) for x in excl)), None)
            if hit is not None:
                idx[field_name], names[field_name] = hit, str(row[hit])
                used.add(hit)
                break
    return idx, names


def _valid(idx: dict) -> bool:
    return "quantity" in idx and "avg_price" in idx and any(k in idx for k in ("symbol", "isin", "name"))


def detect(rows: list[list], configs: list[dict], forced: str | None):
    """Pick (config, header_row_index0, col_idx, col_map): the row that maps the most fields."""
    cands = [c for c in configs if forced is None or c["name"] == forced]
    if forced and not cands:
        raise PipelineError(f"Broker config '{forced}' not found in config/brokers/")
    best = None
    for cfg in cands:
        for i, row in enumerate(rows[: cfg.get("header_search_rows", 25)]):
            if not any(isinstance(c, str) and c.strip() for c in row):
                continue
            idx, names = map_header(row, cfg)
            if _valid(idx) and (best is None or len(idx) > len(best[2])):
                best = (cfg, i, idx, names)
    if best is None:
        sample = [r for r in rows[:10] if any(c not in (None, "") for c in r)][:3]
        raise PipelineError(
            "Could not find the holdings table: no row in the first 25 has a quantity column, an average/cost price "
            f"column and a symbol, ISIN or company-name column. First rows seen: {sample}. Add the header names to "
            "config/brokers/generic_equity_holdings.json (see README: 'Adding a new broker layout').")
    return best


def parse_holdings(path: Path, forced_broker: str | None = None) -> BrokerParse:
    rows = _read_rows(path)
    if not any(c not in (None, "") for r in rows for c in r):
        raise PipelineError(f"The holdings file '{path.name}' is empty (its first sheet has no data). "
                            "Paste/download the broker data into it, save, and run again.")
    cfg, h_idx, col_idx, col_map = detect(rows, load_broker_configs(), forced_broker)
    header = rows[h_idx]

    asof = None
    if cfg.get("asof_regex"):
        for row in rows[:h_idx]:
            for c in row:
                m = re.search(cfg["asof_regex"], str(c or ""), re.I)
                if m:
                    asof = m.group(1)

    notes, dropped = [], []
    merged: dict[str, Holding] = {}
    for i in range(h_idx + 1, len(rows)):
        row = rows[i]
        get = lambda f: row[col_idx[f]] if f in col_idx and col_idx[f] < len(row) else None  # noqa: E731
        raw_sym = get("symbol") or get("isin") or get("name")
        qty = to_number(get("quantity"))
        avg = to_number(get("avg_price"))
        if raw_sym is None or str(raw_sym).strip() == "" or qty is None or isinstance(raw_sym, (int, float)) and avg is None:
            if any(c not in (None, "") for c in row):
                dropped.append({"row": i + 1, "reason": "not a holding row (footer/blank/total)", "values": row})
            continue
        sym = normalize_symbol(raw_sym)
        if avg is None:
            dropped.append({"row": i + 1, "reason": "missing average price", "values": row})
            notes.append(f"Row {i + 1} ({raw_sym}): no average price - dropped and listed in Exceptions")
            continue
        if qty == 0:
            dropped.append({"row": i + 1, "reason": "zero quantity", "values": row})
            continue
        h = Holding(sym, str(raw_sym).strip(), qty, avg, to_number(get("ltp")), to_number(get("invested_value")),
                    (str(get("name")).strip() if get("name") else None),
                    (str(get("exchange")).strip() if get("exchange") else None), [i + 1],
                    (str(get("isin")).strip().upper() if get("isin") else None))
        if sym in merged:
            m = merged[sym]
            tot = m.quantity + h.quantity
            m.avg_price = (m.quantity * m.avg_price + h.quantity * h.avg_price) / tot if tot else 0.0
            m.quantity = tot
            if h.invested_value_broker is not None:
                m.invested_value_broker = (m.invested_value_broker or 0) + h.invested_value_broker
            m.ltp = m.ltp if m.ltp is not None else h.ltp
            m.source_rows += h.source_rows
            notes.append(f"Duplicate symbol {sym} on rows {m.source_rows} merged (weighted average price)")
        else:
            merged[sym] = h
    holdings = [h for h in merged.values() if h.quantity != 0]
    for h in merged.values():
        if h.quantity == 0:
            dropped.append({"row": h.source_rows, "reason": "zero quantity after merging duplicates", "values": [h.raw_symbol]})
    if not holdings:
        raise PipelineError("No holdings rows found in the holdings file.")
    return BrokerParse(cfg["name"], h_idx + 1, col_map, rows, holdings, notes, dropped, asof)

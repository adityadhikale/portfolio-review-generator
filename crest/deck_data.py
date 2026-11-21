"""Deck data, read back from the recalculated Excel (single source of truth)."""
from __future__ import annotations

import math
from pathlib import Path

import openpyxl

from .validate import EXCEL_ERRORS

COLS = {"key": 1, "name": 2, "ticker": 3, "sector": 4, "mcap": 5, "qty": 6, "avg": 7, "invested": 8, "cmp": 9,
        "current": 10, "pnl": 11, "pnl_pct": 12, "alloc": 13, "views": 14}
SNAP = {"key": 1, "name": 2, "ticker": 3, "sector": 4, "mcap": 5, "cmp": 6, "views": 7, "short_name": 8,
        "deck_group": 9, "result": 10, "rating": 11, "sentiment": 12, "commentary": 13, "is_etf": 14}


def scan_errors(xlsx: Path) -> list:
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    bad = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.strip().startswith(EXCEL_ERRORS):
                    bad.append(f"{ws.title}!{c.coordinate} = {c.value}")
    return bad


def read_portfolio(xlsx: Path, layout) -> dict:
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    ws = wb["Portfolio"]
    rows = []
    for r in range(layout.first_row, layout.last_row + 1):
        rec = {k: ws.cell(r, c).value for k, c in COLS.items()}
        if rec["key"]:
            rows.append(rec)
    snap = {}
    sws = wb["Universal_Snapshot"]
    for r in range(2, sws.max_row + 1):
        rec = {k: sws.cell(r, c).value for k, c in SNAP.items()}
        if rec["key"]:
            snap[rec["key"]] = rec
    summary = {"invested": ws["C3"].value, "present": ws["C4"].value, "pnl": ws["C5"].value, "pnl_pct": ws["C6"].value}
    return {"rows": rows, "snapshot": snap, "summary": summary}


def python_values(priced) -> dict:
    """Independent recomputation of what the formulas should produce (for the cross-check)."""
    out = {}
    for m in priced:
        h, u = m.holding, m.uni
        price = _num(u.values.get("cmp"))
        out[h.symbol] = {"qty": h.quantity, "avg": h.avg_price, "invested": h.quantity * h.avg_price,
                         "cmp": price, "current": h.quantity * price}
    total = sum(v["current"] for v in out.values())
    for v in out.values():
        v["pnl"] = v["current"] - v["invested"]
        v["pnl_pct"] = v["pnl"] / v["invested"] if v["invested"] else 0
        v["alloc"] = v["current"] / total if total else 0
    return out


def _num(x):
    try:
        f = float(x)
        return None if math.isnan(f) else f
    except (TypeError, ValueError):
        return None


def compare(excel: dict, py: dict) -> list:
    mism = []
    for r in excel["rows"]:
        p = py.get(r["key"])
        if p is None:
            mism.append(f"{r['key']}: in Excel but not in Python set")
            continue
        for k, tol in (("invested", 0.01), ("current", 0.01), ("pnl", 0.01), ("pnl_pct", 1e-6), ("alloc", 1e-6),
                       ("qty", 1e-9), ("avg", 1e-9), ("cmp", 1e-9)):
            ev = _num(r[k])
            if ev is None or abs(ev - p[k]) > tol:
                mism.append(f"{r['key']} {k}: Excel {r[k]} vs Python {p[k]}")
    return mism


def build_deck_data(excel: dict, *, client, period, asof, group_order, unmapped_group, tail_pct, top_n):
    snap = excel["snapshot"]
    hold = []
    for r in excel["rows"]:
        s = snap.get(r["key"], {})
        is_etf = s.get("is_etf") == "Yes"
        hold.append({
            "symbol": r["key"], "ticker": r["ticker"], "name": r["name"],
            "short_name": s.get("short_name") or None, "sector": r["sector"], "deck_group": s.get("deck_group"),
            "is_etf": is_etf, "market_cap_cr": r["mcap"], "quantity": _num(r["qty"]), "avg_price": _num(r["avg"]),
            "cmp": _num(r["cmp"]), "invested": _num(r["invested"]), "current_value": _num(r["current"]),
            "pnl": _num(r["pnl"]), "gain_pct": round(_num(r["pnl_pct"]) * 100, 2),
            "weight_pct": round(_num(r["alloc"]) * 100, 2), "weight_exact": _num(r["alloc"]) * 100,
            "view": r["views"] or "", "rating": s.get("rating") or "", "sentiment": s.get("sentiment") or "",
            "result_line": s.get("result") or "", "management_commentary": s.get("commentary") or "",
        })
    from .util import short_display_name
    for h in hold:
        h["display_name"] = h["short_name"] or short_display_name(h["name"])
    hold.sort(key=lambda x: -x["weight_exact"])
    eq = [h for h in hold if not h["is_etf"]]
    rated = [h for h in eq if h["rating"] in ("Excellent", "Good", "Average", "Weak")]
    pos = [h for h in rated if h["rating"] in ("Excellent", "Good")]
    avg = [h for h in rated if h["rating"] == "Average"]
    weak = [h for h in rated if h["rating"] == "Weak"]
    tail = [h for h in hold if h["weight_exact"] < tail_pct]
    groups = []
    order = list(group_order) + ([unmapped_group] if unmapped_group not in group_order else [])
    for g in order:
        members = [h for h in eq if h["deck_group"] == g]
        if members:
            groups.append({"group": g, "symbols": [h["symbol"] for h in members]})
    views = {}
    for h in hold:
        views.setdefault(h["view"] or "(none)", []).append(h["display_name"])
    sm = excel["summary"]
    return {
        "run": {"client": client, "period": period.compact, "period_label": period.label,
                "quarter_end": period.quarter_end, "as_of": asof},
        "portfolio": {"invested_value": round(sm["invested"], 2), "present_value": round(sm["present"], 2),
                      "unrealized_pnl": round(sm["pnl"], 2), "unrealized_pnl_pct": round(sm["pnl_pct"] * 100, 2),
                      "holdings_count": len(hold), "equity_count": len(eq), "etf_count": len(hold) - len(eq),
                      "etfs": [h["display_name"] for h in hold if h["is_etf"]]},
        "rating_counts": {"positive": len(pos), "excellent": sum(h["rating"] == "Excellent" for h in rated),
                          "good": sum(h["rating"] == "Good" for h in rated), "average": len(avg), "weak": len(weak),
                          "rated": len(rated), "unrated_equities": len(eq) - len(rated),
                          "positive_pct": round(len(pos) / len(rated) * 100, 1) if rated else 0.0},
        "top_holdings": [{"name": h["display_name"], "symbol": h["symbol"], "weight_pct": h["weight_pct"],
                          "gain_pct": h["gain_pct"], "view": h["view"]} for h in hold[:top_n]],
        "top_n_weight_pct": round(sum(h["weight_exact"] for h in hold[:top_n]), 2),
        "tail_below_threshold": {"threshold_pct": tail_pct, "count": len(tail),
                                 "weight_pct": round(sum(h["weight_exact"] for h in tail), 2),
                                 "names": [h["display_name"] for h in tail]},
        "by_rating": {"Excellent": [h["display_name"] for h in rated if h["rating"] == "Excellent"],
                      "Good": [h["display_name"] for h in rated if h["rating"] == "Good"],
                      "Average": [h["display_name"] for h in avg], "Weak": [h["display_name"] for h in weak]},
        "by_view": views,
        "groups": groups,
        "holdings": [{k: v for k, v in h.items() if k != "weight_exact"} for h in hold],
        "_weights_exact": {h["symbol"]: h["weight_exact"] for h in hold},
    }

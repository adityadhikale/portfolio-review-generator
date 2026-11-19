"""Validation gate. Each check is PASS / WARN / FAIL; any FAIL stops the run after report.md is written."""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from .util import CONFIG_DIR, load_json

EXCEL_ERRORS = ("#N/A", "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#NUM!", "#NULL!", "Err:")


@dataclass
class Check:
    name: str
    status: str  # PASS | WARN | FAIL
    detail: str
    items: list = field(default_factory=list)


def tieout(all_holdings, tol_extra: float) -> Check:
    """Broker's own invested total vs Σ qty × avg price. Broker avg prices are rounded to 2 dp, so the
    allowed gap is Σ(qty × ₹0.005) + a small fixed amount."""
    have = [h for h in all_holdings if h.invested_value_broker is not None]
    calc = sum(h.quantity * h.avg_price for h in all_holdings)
    if not have:
        return Check("Invested value ties to broker file", "WARN",
                     f"Broker file has no invested-value column; Σ Qty × Avg Price = ₹{calc:,.2f} (not tied out)")
    broker = sum(h.invested_value_broker for h in have)
    tol = sum(abs(h.quantity) * 0.005 for h in all_holdings) + tol_extra
    diff = calc - broker
    status = "PASS" if abs(diff) <= tol and len(have) == len(all_holdings) else "FAIL"
    return Check("Invested value ties to broker file", status,
                 f"Broker INVESTMENT VALUE total ₹{broker:,.2f} vs Σ Qty × Avg Price ₹{calc:,.2f}; difference "
                 f"₹{diff:,.2f} (allowed ±₹{tol:,.2f} for 2-dp rounding of average price)")


def excel_ties(excel_invested: float, exceptions_invested: float, all_holdings) -> Check:
    calc = sum(h.quantity * h.avg_price for h in all_holdings)
    gap = excel_invested + exceptions_invested - calc
    st = "PASS" if abs(gap) < 0.05 else "FAIL"
    return Check("Excel invested + Exceptions = broker holdings", st,
                 f"Excel Portfolio invested ₹{excel_invested:,.2f} + Exceptions ₹{exceptions_invested:,.2f} "
                 f"vs Σ Qty × Avg ₹{calc:,.2f} (gap ₹{gap:,.2f})")


def allocation(sum_alloc: float, tol_pct: float) -> Check:
    st = "PASS" if abs(sum_alloc - 1) * 100 <= tol_pct else "FAIL"
    return Check("% Allocation sums to 100%", st, f"Σ % Allocation = {sum_alloc * 100:.4f}%")


def excel_errors(cells: list) -> Check:
    if cells:
        return Check("No Excel error values after recalculation", "FAIL", f"{len(cells)} error cell(s)", cells[:50])
    return Check("No Excel error values after recalculation", "PASS", "No #N/A, #REF!, #DIV/0!, #VALUE!, #NAME? found")


def python_vs_excel(mismatches: list) -> Check:
    if mismatches:
        return Check("Excel formulas agree with independent Python calculation", "FAIL",
                     f"{len(mismatches)} mismatch(es)", mismatches[:50])
    return Check("Excel formulas agree with independent Python calculation", "PASS",
                 "Invested, current value, P&L, P&L %, allocation agree to ₹0.01 / 0.0001%")


def sector_sanity(matched, mode: str) -> Check:
    rules = load_json(CONFIG_DIR / "sector_rules.json")["rules"]
    flags = []
    for m in matched:
        u = m.uni
        name = str(u.values["name"] or "")
        sector = str(u.values["sector"] or "").strip()
        if not sector:
            continue  # blank sector is reported separately as "Missing sector"
        words = set(re.findall(r"[a-z0-9]+", name.lower())) | {u.key.lower()}
        allowed, hit = set(), []
        for r in rules:
            kws = [k for k in r["keywords"] if k in words or (k == "bees" and u.key.lower().endswith("bees"))]
            if kws:
                hit += kws
                allowed |= set(r["allowed"])
        if hit and sector not in allowed:
            flags.append({"symbol": u.key, "name": name, "sector": sector, "keywords": hit,
                          "expected": sorted(allowed), "universal_row": u.row})
    # swap detector: A has B's expected sector and B has A's -> classic row-shift symptom
    swaps = []
    for i, a in enumerate(flags):
        for b in flags[i + 1:]:
            if a["sector"] in b["expected"] and b["sector"] in a["expected"]:
                swaps.append(f"{a['name']} <-> {b['name']} (Universal rows {a['universal_row']} and {b['universal_row']})")
    detail = f"{len(flags)} implausible sector(s)"
    if swaps:
        detail += "; sector values look swapped between: " + "; ".join(swaps)
    if not flags:
        return Check("Sector sanity", "PASS", "No implausible sectors by name keywords")
    return Check("Sector sanity", "FAIL" if mode == "fail" else "WARN", detail, flags)


def simple(name: str, items: list, warn_if_any=True, fail=False, ok_detail="None") -> Check:
    if not items:
        return Check(name, "PASS", ok_detail)
    return Check(name, "FAIL" if fail else ("WARN" if warn_if_any else "PASS"), f"{len(items)} item(s)", items)

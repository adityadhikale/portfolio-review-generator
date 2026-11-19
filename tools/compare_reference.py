"""Compare a run's outputs with the reference files (samples/reference_portfolio.xlsx and the reference deck)
and write compare_report.md into the run folder, with an explanation per difference category.

Usage: python tools/compare_reference.py <run_dir> [--ref-excel ...] [--ref-deck ...]
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

import openpyxl
from pptx import Presentation

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from crest.util import normalize_symbol  # noqa: E402


def _code(ticker: str) -> str:
    return normalize_symbol(str(ticker).split(":", 1)[-1])


def ref_excel(path: Path) -> dict:
    ws = openpyxl.load_workbook(path, data_only=True).active
    hdr_row = next(r for r in range(1, 20) if ws.cell(r, 3).value == "Ticker")
    hdr = [ws.cell(hdr_row, c).value for c in range(2, 15)]
    rows = {}
    order = []
    for r in range(hdr_row + 1, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(2, 15)]
        if not vals[0]:
            continue
        d = dict(zip(hdr, vals))
        rows[_code(d["Ticker"])] = d
        order.append(_code(d["Ticker"]))
    summ = {ws.cell(r, 2).value: ws.cell(r, 3).value for r in range(2, hdr_row)}
    formulas = sum(1 for row in openpyxl.load_workbook(path).active.iter_rows() for c in row
                   if isinstance(c.value, str) and c.value.startswith("="))
    return {"rows": rows, "order": order, "summary": summ, "formulas": formulas, "header_row": hdr_row}


def deck_tables(path: Path) -> tuple[list, dict]:
    prs = Presentation(path)
    titles, stocks = [], {}
    for i, s in enumerate(prs.slides, start=1):
        title = next((sh.text_frame.text for sh in s.shapes if sh.has_text_frame and sh.name == "Text 1"), "")
        titles.append(title)
        for sh in s.shapes:
            if not sh.has_table:
                continue
            for pos, r in enumerate(list(sh.table.rows)[1:], start=1):
                c = [x.text for x in r.cells]
                code = _code(c[0].split("\n")[-1].replace(" ", ""))
                group = re.sub(r"\s*\(\d/\d\)$", "", title.split(" - ", 1)[-1])
                stocks[code] = {"slide": i, "pos": pos, "group": group, "name": c[0].split("\n")[0],
                                "weight": c[1], "result": re.sub(r"\s+", " ", c[2]).strip(), "rating": c[3].strip(),
                                "commentary": re.sub(r"\s+", " ", c[4]).strip(),
                                "sentiment": re.sub(r"^[^A-Za-z]+", "", c[5]).strip(), "view": c[6].strip()}
    return titles, stocks


def slide_text(path: Path, idx: int, name: str) -> str:
    s = Presentation(path).slides[idx]
    return next((sh.text_frame.text for sh in s.shapes if sh.has_text_frame and sh.name == name), "")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("run_dir")
    ap.add_argument("--ref-excel", default=str(ROOT / "samples" / "reference_portfolio.xlsx"))
    ap.add_argument("--ref-deck", default=str(ROOT / "samples" / "reference_deck_Q1FY27.pptx"))
    a = ap.parse_args()
    run = Path(a.run_dir)
    out = run / "outputs"
    ni = json.loads((out / "narrative_input.json").read_text(encoding="utf-8"))
    ps = ni["run"]["price_source"]
    ours = {h["symbol"]: h for h in ni["holdings"]}
    xl = next(out.glob("* Portfolio.xlsx"))
    wb = openpyxl.load_workbook(xl)
    our_formulas = sum(1 for ws in wb.worksheets for row in ws.iter_rows() for c in row
                       if isinstance(c.value, str) and c.value.startswith("="))
    ref = ref_excel(Path(a.ref_excel))
    # map our broker symbols to ticker codes (IndiGrid: INDIGRID -> 540565)
    code_of = {s: _code(h["ticker"]) for s, h in ours.items()}
    by_code = {code_of[s]: h for s, h in ours.items()}

    L = ["# Comparison with the reference files", "",
         f"Run: `{run}`  ", f"Price source in this run: **{ps}**", ""]

    # ---------------- Excel
    L += ["## Excel vs reference_portfolio.xlsx", "", "### Structure", "",
          "| Aspect | Reference | This run | Why |", "|---|---|---|---|",
          f"| Formulas | {ref['formulas']} | {our_formulas} | Every calculated cell is now a live formula (the reference has none) |",
          f"| Sheets | 1 (Sheet1) | {', '.join(wb.sheetnames)} | Raw upload, normalised holdings, Universal snapshot and Exceptions added for audit |",
          f"| Table header row | {ref['header_row']} | 8 | Same layout (Summary block B2:C6, table from column B) |",
          "| Row order | not sorted by allocation (e.g. Cummins 2.93% above Divi's 4.02%) | by current value, descending | Deterministic order |",
          "| Hidden column A | none | normalised symbol (lookup key) | Needed for INDEX/MATCH to Universal_Snapshot |", ""]
    cols = [("Name", "name", None), ("Ticker", "ticker", None), ("Sector", "sector", None),
            ("Market Cap (Cr)", "market_cap_cr", 0.01), ("Quantity Available", "quantity", 1e-9),
            ("Average Price", "avg_price", 1e-9), ("Invested Value", "invested", 0.5), ("CMP", "cmp", 1e-9),
            ("Current Value", "current_value", 0.5), ("Unrealized P&L", "pnl", 0.5),
            ("Unrealized P&L (%)", "gain_pct", 0.0051), ("% Allocation", "weight_pct", 0.0051), ("Views", "view", None)]
    diffs = []
    for code, r in ref["rows"].items():
        h = by_code.get(code)
        if not h:
            diffs.append((code, "row", "present", "missing", "Not in this run"))
            continue
        for lab, key, tol in cols:
            rv, ov = r[lab], h[key]
            if key in ("gain_pct", "weight_pct") and isinstance(rv, (int, float)):
                rv = round(rv * 100, 2)
            if tol is None:
                if str(rv or "").strip() != str(ov or "").strip():
                    diffs.append((code, lab, rv, ov, ""))
            else:
                try:
                    if abs(float(rv) - float(ov)) > tol:
                        diffs.append((code, lab, rv, ov, ""))
                except (TypeError, ValueError):
                    if str(rv) != str(ov):
                        diffs.append((code, lab, rv, ov, ""))
    why = {
        "Ticker": "BOM: prefix rewritten as BSE: (rule 3)",
        "Market Cap (Cr)": "ETF market cap shown as 'ETF' instead of #N/A",
        "Invested Value": "Reference rounds Qty x Avg Price to whole rupees on some rows; this run keeps the exact product",
        "CMP": f"price_source = {ps}" + (" (broker LTP), reference used a different price set" if ps == "broker_ltp" else ""),
        "Current Value": "follows CMP", "Unrealized P&L": "follows CMP / invested rounding",
        "Unrealized P&L (%)": "follows CMP", "% Allocation": "follows CMP (and the reference rounds to 4 dp)",
    }
    by_col = {}
    for d in diffs:
        by_col.setdefault(d[1], []).append(d)
    L += ["### Value differences (54 reference rows)", ""]
    if not diffs:
        L.append("None.")
    for col, ds in by_col.items():
        L += [f"**{col}** - {len(ds)} row(s). {why.get(col, '')}", "", "| Stock | Reference | This run |", "|---|---|---|"]
        for code, _c, rv, ov, _w in ds[:60]:
            L.append(f"| {code} | {rv} | {ov} |")
        L.append("")
    rs = ref["summary"]
    L += ["### Summary block", "", "| Item | Reference | This run |", "|---|---|---|",
          f"| Invested Value | {rs.get('Invested Value')} | {ni['portfolio']['invested_value']} |",
          f"| Present Value | {rs.get('Present Value')} | {ni['portfolio']['present_value']} |",
          f"| Unrealized P&L | {rs.get('Unrealized P&L')} | {ni['portfolio']['unrealized_pnl']} |",
          f"| Unrealized P&L % | {rs.get('Unrealized P&L Pct.')} | {ni['portfolio']['unrealized_pnl_pct']}% |", ""]

    # ---------------- Deck
    deck = next(out.glob("*_Review.pptx"))
    ref_titles, ref_st = deck_tables(Path(a.ref_deck))
    our_titles, our_st = deck_tables(deck)
    L += ["## Deck vs reference deck", "", "### Slides", "", "| # | reference deck | This run |", "|---|---|---|"]
    for i in range(max(len(ref_titles), len(our_titles))):
        rt = ref_titles[i] if i < len(ref_titles) else ""
        ot = our_titles[i] if i < len(our_titles) else ""
        L.append(f"| {i + 1} | {rt} | {ot}{'' if rt == ot else ' **(differs)**'} |")
    L.append("")
    L += ["Cover line - reference: `" + slide_text(Path(a.ref_deck), 0, "Text 2") + "`  ",
          "Cover line - this run: `" + slide_text(deck, 0, "Text 2") + "`", ""]
    L += ["### Slide 5 counts", "", f"Reference: {slide_text(Path(a.ref_deck), 4, 'Text 3')} / {slide_text(Path(a.ref_deck), 4, 'Text 5')} / "
          f"{slide_text(Path(a.ref_deck), 4, 'Text 7')}  ",
          f"This run: {slide_text(deck, 4, 'Text 3')} / {slide_text(deck, 4, 'Text 5')} / {slide_text(deck, 4, 'Text 7')}", ""]
    L += ["### Slide 4 card headings (computed from data in this run)", "", "| Reference | This run |", "|---|---|"]
    for n in ("Text 4", "Text 7", "Text 10", "Text 13"):
        L.append(f"| {slide_text(Path(a.ref_deck), 3, n)} | {slide_text(deck, 3, n)} |")
    L.append("")
    fields = [("group", "Group"), ("weight", "Wt."), ("result", "Result"), ("rating", "Rating"),
              ("commentary", "Commentary"), ("sentiment", "Sentiment"), ("view", "Current View"), ("name", "Display name")]
    dd = {f: [] for f, _ in fields}
    pos = []
    for code, r in ref_st.items():
        o = our_st.get(code)
        if not o:
            dd["group"].append((code, r["group"], "not in tables"))
            continue
        for f, _ in fields:
            if r[f] != o[f]:
                dd[f].append((code, r[f], o[f]))
        if (r["slide"], r["pos"]) != (o["slide"], o["pos"]):
            pos.append((code, f"slide {r['slide']} row {r['pos']}", f"slide {o['slide']} row {o['pos']}"))
    extra = [c for c in our_st if c not in ref_st]
    wwhy = {"weight": f"Weights come from the Excel % Allocation at price_source={ps}. The reference deck used a third price "
                      "set that matches neither the sample Excel CMPs nor the broker LTPs (e.g. HBL 8.78% vs 9.16% vs 8.88%).",
            "view": "Universal Sheet 'Views' is authoritative (your decision). The reference deck's views differ from the sample Excel views.",
            "result": "Copied from the deck into the test fixture, so differences are whitespace only, if any.",
            "group": "Deck Group copied from the deck into the test fixture.",
            "name": "Short Name copied from the deck into the test fixture."}
    L += ["### Results-table rows (53 equities)", ""]
    for f, lab in fields:
        L.append(f"**{lab}: {len(dd[f])} difference(s).** {wwhy.get(f, '')}")
        if dd[f]:
            L += ["", "| Stock | Reference | This run |", "|---|---|---|"] + [f"| {c} | {a_} | {b_} |" for c, a_, b_ in dd[f]]
        L.append("")
    L += [f"**Row position: {len(pos)} difference(s).** Rows are sorted by weight within each group, so a different "
          "price set reorders some rows.", ""]
    if pos:
        L += ["| Stock | Reference | This run |", "|---|---|---|"] + [f"| {c} | {a_} | {b_} |" for c, a_, b_ in pos] + [""]
    if extra:
        L += [f"Rows in this run but not in the reference tables: {extra}", ""]
    (run / "compare_report.md").write_text("\n".join(L), encoding="utf-8")
    print("written", run / "compare_report.md")


if __name__ == "__main__":
    main()

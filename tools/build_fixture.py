"""Build the TEST-ONLY Universal Sheet fixture and the blank Universal Sheet template.

The fixture uses only the supplied sample files:
  * Name, Ticker, Sector, Market Cap, CMP, Views  <- samples/reference_portfolio.xlsx (kept as-is, wrong sectors included)
  * Short Name, Result Line, Rating, Sentiment,
    Management Commentary, Deck Group             <- samples/reference_deck_Q1FY27.pptx tables
  * Symbol (only where the ticker is a BSE code)  <- samples/holdings.xlsx row with the same Qty and Avg Price
No value is invented. Run:  python tools/build_fixture.py
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from pptx import Presentation

ROOT = Path(__file__).resolve().parent.parent
SAMPLES = ROOT / "samples"
COLS = ["Name", "Ticker", "Symbol", "Alt Symbols", "Short Name", "Sector", "Deck Group", "Market Cap (Cr)", "CMP", "Views",
        "Q Result Period", "Result Line", "Rating", "Sentiment", "Management Commentary", "Exchange", "Instrument Type"]
LEGEND = [
    ("Name", "Required", "Full company / instrument name"),
    ("Ticker", "Required", "NSE:SYMBOL or BSE:CODE (BOM: is accepted and rewritten as BSE:). A bare symbol defaults to NSE and is flagged."),
    ("Symbol", "Optional", "Broker trading symbol, when it differs from the ticker code (e.g. INDIGRID for BSE:540565). Used for matching."),
    ("Alt Symbols", "Optional", "Comma-separated BSE scrip codes and former NSE symbols; also used for matching broker files."),
    ("Short Name", "Optional", "Name shown in the deck. Blank = Name without 'Limited'."),
    ("Sector", "Required", "Sector; mapped to a deck group via config/sector_groups.json"),
    ("Deck Group", "Optional", "Overrides the sector-to-group mapping for this stock (must be a group in config/sector_groups.json)"),
    ("Market Cap (Cr)", "Required", "Number, INR crore. Leave blank for ETFs."),
    ("CMP", "Required", "Current market price, INR (used when price_source = universal_cmp)"),
    ("Views", "Required", "CREST view, e.g. Hold / Sell / Add To 3% (shown in Excel and deck)"),
    ("Q Result Period", "Optional", "Quarter the result fields refer to, e.g. Q2 FY27. Must match --period (run stops otherwise)."),
    ("Result Line", "Required (quarterly)", "One line, e.g. Rev +27.8% YoY (₹3,080cr); PAT +66% YoY (₹902cr)"),
    ("Rating", "Column required (quarterly)", "Excellent / Good / Average / Weak. Blank = from config/rating_rules.json once approved, else n/a"),
    ("Sentiment", "Column required (quarterly)", "Bullish / Neutral / Bearish. Blank = from the rating via the approved rules, else n/a"),
    ("Management Commentary", "Required (quarterly)", "One or two sentences for the results table"),
    ("Exchange", "Optional", "NSE / BSE, used only when Ticker has no prefix"),
    ("Instrument Type", "Optional", "Equity / ETF. Blank = inferred (name contains ETF or symbol ends in BEES, flagged)"),
]


def _clean(s):
    return re.sub(r"\s*\n\s*", " ", (s or "")).strip()


def deck_rows(pptx: Path) -> dict:
    out = {}
    for s in Presentation(pptx).slides:
        tables = [sh for sh in s.shapes if sh.has_table]
        title = next((sh.text_frame.text for sh in s.shapes if sh.has_text_frame and sh.name == "Text 1"), "")
        if not tables:
            continue
        group = re.sub(r"\s*\(\d/\d\)\s*$", "", title.split(" - ", 1)[-1]).strip()
        for r in list(tables[0].table.rows)[1:]:
            c = [cell.text for cell in r.cells]
            lines = [x for x in c[0].split("\n") if x.strip()]
            m = re.match(r"^(NSE|BSE)\s*:\s*(\S+)$", lines[-1].strip())
            if not m:
                continue
            out[m.group(2).upper()] = {
                "short": lines[0].strip(), "result": _clean(c[2]), "rating": c[3].strip(),
                "commentary": _clean(c[4]), "sentiment": re.sub(r"^[^A-Za-z]+", "", c[5]).strip(), "group": group}
    return out


def build(out_fixture: Path, out_template: Path):
    ws = openpyxl.load_workbook(SAMPLES / "reference_portfolio.xlsx").active
    broker = [r for r in openpyxl.load_workbook(SAMPLES / "holdings.xlsx").active.iter_rows(min_row=3, values_only=True)
              if isinstance(r[1], (int, float))]
    deck = deck_rows(SAMPLES / "reference_deck_Q1FY27.pptx")
    rows = []
    for r in ws.iter_rows(min_row=9, values_only=True):
        name, ticker = r[1], r[2]
        if not name:
            continue
        code = ticker.split(":", 1)[1].upper()
        symbol = ""
        if code.isdigit():  # BSE code: link to the broker symbol through identical Qty and Avg Price
            hits = [b[0] for b in broker if b[1] == r[5] and abs(b[3] - r[6]) < 1e-9]
            symbol = hits[0] if len(hits) == 1 else ""
        d = deck.get(code, {})
        mcap = r[4] if isinstance(r[4], (int, float)) else None
        rows.append([name, ticker, symbol, "", d.get("short", ""), r[3], d.get("group", ""), mcap, r[8], r[13],
                     "Q1 FY27" if d.get("result") else "", d.get("result", ""), d.get("rating", ""), d.get("sentiment", ""), d.get("commentary", ""), "", ""])

    wb = openpyxl.Workbook()
    u = wb.active
    u.title = "Universal"
    u.append(COLS)
    for row in rows:
        u.append(row)
    note = wb.create_sheet("README")
    note["A1"] = "TEST FIXTURE ONLY - not a real Universal Sheet."
    note["A2"] = ("Built by tools/build_fixture.py from samples/reference_portfolio.xlsx (Name, Ticker, Sector, Market Cap, "
                  "CMP, Views - wrong sectors deliberately kept) and the Q1FY27 reference deck tables (Short Name, Result "
                  "Line, Rating, Sentiment, Management Commentary, Deck Group).")
    out_fixture.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_fixture)

    # ---- blank template for the real Universal Sheet
    t = openpyxl.Workbook()
    ut = t.active
    ut.title = "Universal"
    ut.append(COLS)
    hdr = PatternFill("solid", fgColor="0D0D0D")
    for c in ut[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = hdr
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for j, w in enumerate([34, 16, 12, 16, 24, 22, 30, 14, 10, 16, 12, 50, 10, 10, 60, 9, 14], start=1):
        ut.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ut.freeze_panes = "B2"
    lg = t.create_sheet("Legend")
    lg.append(["Column", "Status", "What to enter"])
    for row in LEGEND:
        lg.append(list(row))
    for c in lg[1]:
        c.font = Font(name="Arial", bold=True)
    lg.column_dimensions["A"].width = 24
    lg.column_dimensions["B"].width = 20
    lg.column_dimensions["C"].width = 110
    out_template.parent.mkdir(parents=True, exist_ok=True)
    t.save(out_template)
    return len(rows), sum(1 for r in rows if r[12])


if __name__ == "__main__":
    n, rated = build(ROOT / "tests" / "fixtures" / "universal_sample.xlsx", ROOT / "templates" / "Universal_Sheet_template.xlsx")
    print(f"fixture rows: {n}, with deck result data: {rated}")
    sys.exit(0)

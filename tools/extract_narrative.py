"""TEST HELPER: pull the narrative text out of a finished deck (e.g. the Q1FY27 reference deck) into the
narrative.json shape. Used to test narrative insertion and the number checker, and to show what a
narrative.json looks like.  Usage: python tools/extract_narrative.py <deck.pptx> <out.json>"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

from pptx import Presentation


def _t(slide, name):
    for s in slide.shapes:
        if s.name == name and s.has_text_frame:
            return s.text_frame
    return None


def _after(tf, label):
    txt = tf.text.strip()
    return txt[len(label):].strip() if txt.startswith(label) else txt


def extract(deck: Path, name_to_symbol: dict) -> dict:
    sl = Presentation(deck).slides
    s2, s3, s4, s5, s15 = sl[1], sl[2], sl[3], sl[4], sl[len(sl) - 1]
    cards4 = []
    for hn, bn in (("Text 4", "Text 5"), ("Text 7", "Text 8"), ("Text 10", "Text 11"), ("Text 13", "Text 14")):
        head = _t(s4, hn).text
        nm = head.split(" - ")[0].strip()
        sym = name_to_symbol.get(nm) or next((v for k, v in name_to_symbol.items() if k.startswith(nm)), nm)
        cards4.append({"symbol": sym, "body": _t(s4, bn).text.strip()})
    return {
        "slide2": {"headline": _t(s2, "Text 1").text.split(" - ", 1)[1].strip(),
                   "read_through": [p.text.strip() for p in _t(s2, "Text 3").paragraphs if p.text.strip()],
                   "themes": [_t(s2, n).text.strip() for n in ("Text 6", "Text 8", "Text 10", "Text 12", "Text 14", "Text 16")],
                   "portfolio_implication": _after(_t(s2, "Text 18"), "Portfolio implication:")},
        "slide3": {"title_suffix": re.split(r"\s[-–]\s", _t(s3, "Text 1").text, maxsplit=1)[1].strip(),
                   "subtitle": _t(s3, "Text 2").text.strip(),
                   "cards": [{"heading": _t(s3, h).text.strip(), "body": _t(s3, b).text.strip()} for h, b in
                             (("Text 4", "Text 5"), ("Text 7", "Text 8"), ("Text 10", "Text 11"),
                              ("Text 13", "Text 14"), ("Text 16", "Text 17"), ("Text 19", "Text 20"))]},
        "slide4": {"subtitle": _t(s4, "Text 2").text.strip(), "cards": cards4,
                   "portfolio_action": _t(s4, "Text 17").text.strip()},
        "slide5": {"best": _after(_t(s5, "Text 9"), "Best results of the quarter:"),
                   "weakest": _after(_t(s5, "Text 10"), "Weakest results of the quarter:"),
                   "broad_theme": _after(_t(s5, "Text 11"), "Broad theme:")},
        "slide15": {"our_view": _after(_t(s15, "Text 3"), "Our view:"),
                    "investment_stance": _after(_t(s15, "Text 4"), "Investment stance:"),
                    "time_horizon": _after(_t(s15, "Text 5"), "Time horizon:"),
                    "key_themes": _after(_t(s15, "Text 6"), "Key themes:"),
                    "screening_discipline": _after(_t(s15, "Text 7"), "Screening discipline:"),
                    "panel": {"investment_stance": _t(s15, "Text 10").text.strip(),
                              "time_horizon": _t(s15, "Text 12").text.strip(),
                              "approach": _t(s15, "Text 14").text.strip(),
                              "key_themes": _t(s15, "Text 16").text.strip()}},
    }


if __name__ == "__main__":
    import openpyxl
    root = Path(__file__).resolve().parent.parent
    ws = openpyxl.load_workbook(root / "tests" / "fixtures" / "universal_sample.xlsx")["Universal"]
    hdr = [c.value for c in ws[1]]
    m = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(hdr, r))
        sym = d["Symbol"] or str(d["Ticker"]).split(":")[1]
        if d["Short Name"]:
            m[d["Short Name"]] = sym
    out = extract(Path(sys.argv[1]), m)
    Path(sys.argv[2]).write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")
    print("written", sys.argv[2])

"""Fill Deck_-_template.pptx with python-pptx. Layout, branding and fonts come from the template;
this module only replaces text, clones the results-table slide per sector group and removes unused slides."""
from __future__ import annotations

import copy
import math
import re

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from .narrative import get_path

RATING_FILL = {"Excellent": "1F5A31", "Good": "70AD47", "Average": "2F5597", "Weak": "C00000"}
SENTIMENT = {"Bullish": ("▲ Bullish", "2E7D32"), "Neutral": ("▬ Neutral", "2F5597"), "Bearish": ("▼ Bearish", "C00000")}
TEXT = "262626"
NA_FILL = "A6A6A6"

# Template slide numbers (1-based) and shape names
NARRATIVE_SLIDES = (2, 3, 4, 5, 15)
TABLE_SLIDES = range(6, 15)
TABLE_PROTOTYPE = 6


# --------------------------------------------------------------------------- text helpers
def shape(slide, name):
    for s in slide.shapes:
        if s.name == name:
            return s
    raise KeyError(f"Template shape '{name}' not found on slide - template changed?")


def para_replace(p, old: str, new: str) -> bool:
    """Replace `old` in a paragraph even when it spans several runs; formatting of the first run
    touched by the match is used for the replacement text."""
    runs = p.runs
    full = "".join(r.text for r in runs)
    idx = full.find(old)
    if idx < 0:
        return False
    end = idx + len(old)
    pos = 0
    placed = False
    for r in runs:
        rs, re_ = pos, pos + len(r.text)
        pos = re_
        if re_ <= idx or rs >= end:
            continue
        before = r.text[: max(0, idx - rs)]
        after = r.text[max(0, end - rs):] if re_ > end else ""
        if not placed:
            r.text = before + new + after
            placed = True
        else:
            r.text = after
    return True


def replace_everywhere(prs, mapping: dict):
    for slide in prs.slides:
        for p in _all_paragraphs(slide):
            for old, new in mapping.items():
                while para_replace(p, old, new):
                    if old in new:
                        break


def _all_paragraphs(slide):
    for s in slide.shapes:
        if s.has_text_frame:
            yield from s.text_frame.paragraphs
        elif s.has_table:
            for row in s.table.rows:
                for c in row.cells:
                    yield from c.text_frame.paragraphs


def set_text(sh, text: str):
    """Replace a shape's whole text, keeping the first run's formatting."""
    tf = sh.text_frame
    p0 = tf.paragraphs[0]
    for extra in tf.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    runs = p0.runs
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


def fill_token(sh, token: str, value: str, unbold_value: bool = False):
    """Replace `token` inside a shape. If the token shares a run with a bold label, the value goes in
    its own non-bold run (label stays bold)."""
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            if token in r.text and r.text.strip() != token and unbold_value:
                label = r.text.replace(token, "").rstrip()
                r.text = label + " "
                nr = copy.deepcopy(r._r)
                r._r.addnext(nr)
                rpr = nr.find(qn("a:rPr"))
                if rpr is not None:
                    rpr.set("b", "0")
                nr.find(qn("a:t")).text = value
                return True
        if para_replace(p, token, value):
            return True
    return False


def bullets(sh, items: list):
    tf = sh.text_frame
    proto = copy.deepcopy(tf.paragraphs[0]._p)
    body = tf._txBody
    for p in list(tf.paragraphs):
        body.remove(p._p)
    for it in items:
        np_ = copy.deepcopy(proto)
        rs = np_.findall(qn("a:r"))
        rs[0].find(qn("a:t")).text = it
        for r in rs[1:]:
            np_.remove(r)
        body.append(np_)


# --------------------------------------------------------------------------- slide clone / delete / move
def duplicate_slide(prs, src):
    new = prs.slides.add_slide(src.slide_layout)
    for s in list(new.shapes):
        s._element.getparent().remove(s._element)
    for el in src.shapes._spTree.iterchildren():
        if el.tag in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
            continue
        new.shapes._spTree.append(copy.deepcopy(el))
    bg = src._element.cSld.find(qn("p:bg"))
    if bg is not None:
        new._element.cSld.insert(0, copy.deepcopy(bg))
    return new


def delete_slide(prs, slide):
    lst = prs.slides._sldIdLst
    for sid in list(lst):
        if prs.part.related_part(sid.rId) is slide.part:
            prs.part.drop_rel(sid.rId)
            lst.remove(sid)
            return


def move_slide(prs, slide, new_index: int):
    lst = prs.slides._sldIdLst
    for sid in list(lst):
        if prs.part.related_part(sid.rId) is slide.part:
            lst.remove(sid)
            lst.insert(new_index, sid)
            return


# --------------------------------------------------------------------------- table cells
def _rpr(size_pt: float, bold: bool, color: str, italic=False):
    rpr = etree.SubElement(etree.Element(qn("a:r")), qn("a:rPr"))
    rpr.set("lang", "en-US")
    rpr.set("sz", str(int(size_pt * 100)))
    rpr.set("b", "1" if bold else "0")
    if italic:
        rpr.set("i", "1")
    rpr.set("dirty", "0")
    sf = etree.SubElement(rpr, qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr")).set("val", color)
    for tag in ("a:latin", "a:ea", "a:cs"):
        etree.SubElement(rpr, qn(tag)).set("typeface", "Arial")
    return rpr


def write_cell(cell, lines, *, size=8, bold=False, color=TEXT, fill=None, italic=False):
    tx = cell._tc.txBody
    ps = tx.findall(qn("a:p"))
    ppr = ps[0].find(qn("a:pPr"))
    end = ps[0].find(qn("a:endParaRPr"))
    for p in ps:
        tx.remove(p)
    for line in lines:
        p = etree.SubElement(tx, qn("a:p"))
        if ppr is not None:
            p.append(copy.deepcopy(ppr))
        r = etree.SubElement(p, qn("a:r"))
        r.append(_rpr(size, bold, color, italic))
        etree.SubElement(r, qn("a:t")).text = line
        if end is not None:
            p.append(copy.deepcopy(end))
    tcpr = cell._tc.get_or_add_tcPr()
    for old in tcpr.findall(qn("a:solidFill")) + tcpr.findall(qn("a:noFill")):
        tcpr.remove(old)
    if fill:
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor.from_string(fill)


def chunk(n: int, max_single: int, max_split: int):
    if n <= max_single:
        return [n]
    k = math.ceil(n / max_split)
    base, extra = divmod(n, k)
    return [base + (1 if i < extra else 0) for i in range(k)]


def fill_table_slide(slide, title: str, rows: list, dcfg: dict):
    set_text(shape(slide, "Text 1"), title)
    gf = next(s for s in slide.shapes if s.has_table)
    tbl = gf.table._tbl
    trs = tbl.tr_lst
    proto = copy.deepcopy(trs[1])
    for tr in trs[1:]:
        tbl.remove(tr)
    avail = dcfg["table_bottom_limit_in"] - dcfg["table_top_in"] - dcfg["header_row_in"]
    row_h = Inches(min(dcfg["max_row_in"], avail / max(1, len(rows))))
    for _ in rows:
        tr = copy.deepcopy(proto)
        tr.set("h", str(int(row_h)))
        tbl.append(tr)
    gf.height = Emu(int(trs[0].h) + int(row_h) * len(rows))
    table = gf.table
    for i, h in enumerate(rows, start=1):
        c = table.rows[i].cells
        ex, code = (h["ticker"] or ":").split(":", 1)
        write_cell(c[0], [h["display_name"], f"{ex}: {code}"], bold=True)
        write_cell(c[1], [f"{h['weight_pct']:.2f}%"])
        write_cell(c[2], [h["result_line"] or "Results awaited"], italic=not h["result_line"])
        if h["rating"] in RATING_FILL:
            write_cell(c[3], [h["rating"]], bold=True, color="FFFFFF", fill=RATING_FILL[h["rating"]])
        else:
            write_cell(c[3], ["n/a"], bold=True, color="FFFFFF", fill=NA_FILL)
        write_cell(c[4], [h["management_commentary"] or "Commentary awaited"], italic=not h["management_commentary"])
        lab, col = SENTIMENT.get(h["sentiment"], ("n/a", "7F7F7F"))
        write_cell(c[5], [lab], bold=True, color=col)
        write_cell(c[6], [h["view"] or "n/a"], bold=True)


# --------------------------------------------------------------------------- main fill
def fill_deck(template, out_path, *, data: dict, narrative: dict | None, period, client: str, asof: str,
              dcfg: dict, ratings_by_ai: bool = False) -> dict:
    prs = Presentation(template)
    pend = dcfg["pending_marker"]
    pending, problems = [], []
    nv = narrative or {}

    def val(key):
        v = get_path(nv, key)
        if v in (None, "", []):
            pending.append(key)
            return None
        return v

    # ---- period + cover
    lit = {p: (period.compact if " " not in p else period.label) for p in dcfg["template_period_literals"]}
    replace_everywhere(prs, {"Q1 FY2[]": period.label, "Q1 FY[ ]": period.label, **lit})
    # The CONTENT layout footer carries a hard-coded period and client name ("Q1 FY27 Portfolio Review -
    # <client name> (CREST Portfolio)"); rewrite it for this run.
    for layout in prs.slide_layouts:
        for sh in layout.shapes:
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                full = "".join(r.text for r in p.runs)
                m = re.search(r"\|\s*(.+?)\s+Portfolio Review\s*-\s*(.+?)\s*\(CREST Portfolio\)", full)
                if m:
                    para_replace(p, m.group(1), period.label)
                    para_replace(p, m.group(2), client)
    s1 = prs.slides[0]
    for p in shape(s1, "Text 2").text_frame.paragraphs:
        if re.search(r"\[\s*Name\s*\]", "".join(r.text for r in p.runs)):
            full = "".join(r.text for r in p.runs)
            m = re.search(r"-\s*\[\s*Name\s*\]\s*·", full)
            para_replace(p, m.group(0) if m else "[ Name ]", f"- {client} ·" if m else client)
        para_replace(p, "[ dd/mm/yyyy ]", asof)

    s2, s3, s4, s5 = (prs.slides[i] for i in (1, 2, 3, 4))
    s15 = prs.slides[14]
    hl = val("slide2.headline")
    set_text(shape(s2, "Text 1"), f"{period.label} Outlook - {hl or pend}")
    rt = val("slide2.read_through")
    bullets(shape(s2, "Text 3"), rt if isinstance(rt, list) else [rt] if rt else [pend])
    themes = val("slide2.themes") or []
    for i, nm in enumerate(["Text 6", "Text 8", "Text 10", "Text 12", "Text 14", "Text 16"]):
        fill_token(shape(s2, nm), "{ }", themes[i] if i < len(themes) and themes[i] else pend)
    if themes and len(themes) != 6:
        problems.append(f"slide2.themes has {len(themes)} items; the template has 6 tiles")
    fill_token(shape(s2, "Text 18"), "{ }", val("slide2.portfolio_implication") or pend)

    # slide 3
    set_text(shape(s3, "Text 1"), f"Market Outlook - {val('slide3.title_suffix') or pend}")
    set_text(shape(s3, "Text 2"), val("slide3.subtitle") or pend)
    cards = val("slide3.cards") or []
    if cards and len(cards) != 6:
        problems.append(f"slide3.cards has {len(cards)} items; the template has 6 cards")
    for i, (hn, bn) in enumerate([("Text 4", "Text 5"), ("Text 7", "Text 8"), ("Text 10", "Text 11"),
                                  ("Text 13", "Text 14"), ("Text 16", "Text 17"), ("Text 19", "Text 20")]):
        cd = cards[i] if i < len(cards) else {}
        set_text(shape(s3, hn), (cd.get("heading") or pend).strip())
        fill_token(shape(s3, bn), "{ }", cd.get("body") or pend)

    # slide 4 (headings computed from data)
    t4 = get_path(nv, "slide4.title")
    if t4:
        set_text(shape(s4, "Text 1"), t4)
    fill_token(shape(s4, "Text 2"), "{ }", val("slide4.subtitle") or pend)
    by_sym = {h["symbol"]: h for h in data["holdings"]}
    cards4 = val("slide4.cards") or []
    if cards4 and len(cards4) != 4:
        problems.append(f"slide4.cards has {len(cards4)} items; the template has 4 cards")
    for i, (hn, bn) in enumerate([("Text 4", "Text 5"), ("Text 7", "Text 8"), ("Text 10", "Text 11"), ("Text 13", "Text 14")]):
        cd = cards4[i] if i < len(cards4) else {}
        sym = str(cd.get("symbol") or "").strip().upper()
        h = by_sym.get(sym)
        if cd and not h:
            problems.append(f"slide4.cards[{i}].symbol '{sym}' is not a holding in this portfolio")
        head = (f"{h['display_name']} - {h['weight_pct']:.2f}% weight · {h['gain_pct']:+.2f}%" if h else pend)
        set_text(shape(s4, hn), head)
        fill_token(shape(s4, bn), "{ }", cd.get("body") or pend)
    fill_token(shape(s4, "Text 17"), "{ }", val("slide4.portfolio_action") or pend)

    # slide 5: counts are data, the three lines are narrative
    rc = data["rating_counts"]
    no_ratings = rc["rated"] == 0
    for nm, n in (("Text 3", rc["positive"]), ("Text 5", rc["average"]), ("Text 7", rc["weak"])):
        n = "–" if no_ratings else n
        p = shape(s5, nm).text_frame.paragraphs[0]
        runs = p.runs
        size = next((r.font.size for r in runs if r.font.size), None)
        for r in runs:
            if "[" in r.text:
                r.text = str(n)
                if r.font.size is None and size:
                    r.font.size = size
    note = (f"OUT OF {rc['rated']} rated equity holdings, on {period.label} results"
            + (f" (quarter ended {period.quarter_end})" if period.quarter_end else "")
            + (f"; {', '.join(data['portfolio']['etfs'])} tracked separately" if data["portfolio"]["etfs"] else "")
            + ". Positive = Good or Excellent")
    if no_ratings:
        note = (f"RATINGS PENDING: none of the {data['portfolio']['equity_count']} equity holdings has a Rating yet "
                f"({period.label} results).")
    elif rc["unrated_equities"]:
        note += f"; {rc['unrated_equities']} equity holding(s) not rated yet (results awaited)"
    tb = s5.shapes.add_textbox(Emu(7178040), Emu(1234440), Emu(4572000), Emu(731520))
    tb.name = "Text 8"
    tf = tb.text_frame
    tf.word_wrap = True
    for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, side, 0)
    tf.vertical_anchor = 3  # middle
    r = tf.paragraphs[0].add_run()
    r.text = note
    r.font.size, r.font.italic = Pt(9), True
    r.font.color.rgb = RGBColor.from_string("595959")
    for key, nm in (("slide5.best", "Text 9"), ("slide5.weakest", "Text 10"), ("slide5.broad_theme", "Text 11")):
        fill_token(shape(s5, nm), "{ }", val(key) or pend)

    # slide 15
    for key, nm in (("slide15.our_view", "Text 3"), ("slide15.investment_stance", "Text 4"),
                    ("slide15.time_horizon", "Text 5"), ("slide15.key_themes", "Text 6"),
                    ("slide15.screening_discipline", "Text 7")):
        fill_token(shape(s15, nm), "{ }", val(key) or pend, unbold_value=True)
    for key, nm in (("slide15.panel.investment_stance", "Text 10"), ("slide15.panel.time_horizon", "Text 12"),
                    ("slide15.panel.approach", "Text 14"), ("slide15.panel.key_themes", "Text 16")):
        set_text(shape(s15, nm), val(key) or pend)
    # Title: the template says "Positioning for the Next 2-3 Years"; follow this review's time horizon instead.
    t15 = get_path(nv, "slide15.title_suffix") or (
        f"Positioning for the {get_path(nv, 'slide15.panel.time_horizon')}"
        if get_path(nv, "slide15.panel.time_horizon") else None)
    if t15:
        set_text(shape(s15, "Text 1"), f"Market Outlook - {t15}")
    # Source note: the template text (Screener, views from the client's Excel, '*' marks) does not describe
    # how this tool works, so it is rewritten from what the run actually used.
    rating_src = ("Rating and Sentiment from the Universal Sheet where set, otherwise derived from those results "
                  "and the commentary" if ratings_by_ai else "Rating and Sentiment from the CREST Universal Sheet")
    set_text(shape(s15, "Text 17"),
             f"Company results (revenue, PAT) and management commentary from the CREST Universal Sheet for "
             f"{period.label}. {rating_src}. Current View is CREST's own. Holdings and weights from {client}'s "
             f"holdings statement as on {asof}; prices from the CREST Universal Sheet.")

    # ---- results tables: clone the prototype per group / chunk, then drop the template table slides
    proto = prs.slides[TABLE_PROTOTYPE - 1]
    originals = [prs.slides[i - 1] for i in TABLE_SLIDES]
    eq = {h["symbol"]: h for h in data["holdings"]}
    made = []
    for g in data["groups"]:
        members = [eq[s] for s in g["symbols"]]
        sizes = chunk(len(members), dcfg["max_rows_single_slide"], dcfg["max_rows_when_split"])
        start = 0
        for k, n in enumerate(sizes, start=1):
            part = members[start:start + n]
            start += n
            suffix = f" ({k}/{len(sizes)})" if len(sizes) > 1 else ""
            sl = duplicate_slide(prs, proto)
            fill_table_slide(sl, f"Results Analysis {period.label} - {g['group']}{suffix}", part, dcfg)
            made.append((sl, g["group"], suffix, [h["symbol"] for h in part]))
    for i, (sl, *_rest) in enumerate(made):
        move_slide(prs, sl, 5 + i)
    for sl in originals:
        delete_slide(prs, sl)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)
    return {"pending": pending, "problems": problems,
            "table_slides": [{"slide": 6 + i, "group": g, "part": sfx.strip(), "symbols": syms}
                             for i, (_s, g, sfx, syms) in enumerate(made)]}

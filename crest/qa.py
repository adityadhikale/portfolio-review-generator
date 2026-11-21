"""Deck QA: leftover placeholders, house style, text-overflow estimate, render to PDF/PNG."""
from __future__ import annotations

import math
import re
from pathlib import Path

from pptx import Presentation

PLACEHOLDER = re.compile(r"\{\s*\}|\[\s*\]|\{\s*Heading\s*\}|\[\s*Name\s*\]|dd/mm/yyyy|FY2\[\]|FY\[\s*\]|202\[\]|�")
EMU_PER_PT = 12700


def _texts(prs):
    for i, slide in enumerate(prs.slides, start=1):
        for s in slide.shapes:
            if s.has_text_frame:
                yield i, s, None, s.text_frame
            elif s.has_table:
                for ri, row in enumerate(s.table.rows):
                    for ci, c in enumerate(row.cells):
                        yield i, s, (ri, ci), c.text_frame


def placeholders(prs) -> list:
    out = []
    for i, s, cell, tf in _texts(prs):
        for p in tf.paragraphs:
            t = "".join(r.text for r in p.runs)
            if PLACEHOLDER.search(t) or (s.name == "Text 3" and t.strip() == "A"):
                out.append(f"slide {i} {s.name}{'' if cell is None else ' cell ' + str(cell)}: '{t[:80]}'")
    return out


def pending(prs, marker: str) -> list:
    return sorted({f"slide {i}" for i, s, c, tf in _texts(prs) if marker in tf.text})


def style(prs) -> list:
    issues = []
    fonts = set()
    for i, s, cell, tf in _texts(prs):
        text = tf.text
        if re.search(r"\b(?!CREST\b)[Cc][Rr][Ee][Ss][Tt]\b", text):
            issues.append(f"slide {i} {s.name}: brand not written as 'CREST'")
        if s.name == "Text 1" and re.match(r"^\s*\d", text):
            issues.append(f"slide {i}: slide heading starts with a number: '{text[:60]}'")
        for p in tf.paragraphs:
            for r in p.runs:
                if r.font.name:
                    fonts.add(r.font.name)
    extra = fonts - {"Maharlika", "Arial"}
    if extra:
        issues.append(f"fonts outside the template pair (Maharlika headings / Arial body): {sorted(extra)}")
    return issues


def _need_pt(tf, width_pt, default_size):
    lines_pt = 0.0
    for p in tf.paragraphs:
        txt = "".join(r.text for r in p.runs)
        size = next((r.font.size.pt for r in p.runs if r.font.size), None)
        if size is None and p._p.find(".//{*}endParaRPr") is not None and p._p.find(".//{*}endParaRPr").get("sz"):
            size = int(p._p.find(".//{*}endParaRPr").get("sz")) / 100
        size = size or default_size
        cpl = max(1, int(width_pt / (size * 0.5)))  # Arial averages ~0.5 em per character
        n = max(1, math.ceil(len(txt) / cpl)) if txt else 1
        lines_pt += n * size * 1.2
    return lines_pt


def overflow(prs) -> list:
    """Heuristic: estimated wrapped height vs the box (or table cell) height."""
    out = []
    for i, slide in enumerate(prs.slides, start=1):
        for s in slide.shapes:
            if s.has_text_frame and s.text_frame.text.strip() and s.width and s.height:
                tf = s.text_frame
                w = s.width / EMU_PER_PT - 14
                h = s.height / EMU_PER_PT
                need = _need_pt(tf, w, 10.5)
                if need > h * 1.15 and need - h > 6:
                    out.append(f"slide {i} {s.name}: needs ~{need:.0f}pt, box is {h:.0f}pt ('{tf.text[:50]}...')")
            elif s.has_table:
                cols = [c.width / EMU_PER_PT for c in s.table.columns]
                for ri, row in enumerate(s.table.rows):
                    if ri == 0:
                        continue
                    h = row.height / EMU_PER_PT - 4
                    for ci, c in enumerate(row.cells):
                        need = _need_pt(c.text_frame, cols[ci] - 8, 8)
                        if need > h * 1.05:
                            out.append(f"slide {i} table row {ri} col {ci + 1}: needs ~{need:.0f}pt, row is {h:.0f}pt "
                                       f"('{c.text_frame.text[:40]}...')")
    return out


def layout_footer(prs, client: str, period_label: str) -> list:
    """Layout/master text (footers) must name this client and period, not the template's."""
    out = []
    for layout in prs.slide_layouts:
        for sh in layout.shapes:
            if sh.has_text_frame and "Portfolio Review" in sh.text_frame.text:
                t = sh.text_frame.text
                if client not in t or period_label not in t:
                    out.append(f"layout '{layout.name}' footer: '{t}'")
    return out


def run_qa(pptx_path: Path, marker: str, client: str = "", period_label: str = "") -> dict:
    prs = Presentation(pptx_path)
    return {"placeholders": placeholders(prs) + layout_footer(prs, client, period_label),
            "pending": pending(prs, marker), "style": style(prs),
            "overflow": overflow(prs), "slide_count": len(prs.slides)}


def render(pptx_path: Path, outdir: Path, soffice: str, dpi: int) -> tuple[Path, list, list]:
    """PPTX -> PDF (LibreOffice) -> PNG per page (PyMuPDF). Also flags text drawn outside the page."""
    from .office import convert
    pdf = convert(pptx_path, "pdf", outdir, soffice)
    pngs, offpage = [], []
    try:
        import pymupdf as fitz
    except ImportError:  # older PyMuPDF
        import fitz
    doc = fitz.open(pdf)
    img_dir = outdir / "png"
    img_dir.mkdir(parents=True, exist_ok=True)
    for n, page in enumerate(doc, start=1):
        pix = page.get_pixmap(dpi=dpi)
        f = img_dir / f"slide_{n:02d}.png"
        pix.save(f)
        pngs.append(f)
        W, H = page.rect.width, page.rect.height
        for b in page.get_text("blocks"):
            x0, y0, x1, y1, txt = b[:5]
            if y1 > H + 1 or x1 > W + 1:
                offpage.append(f"slide {n}: text runs off the page: '{txt.strip()[:50]}'")
    doc.close()
    return pdf, pngs, offpage

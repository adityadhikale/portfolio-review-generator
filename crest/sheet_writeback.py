"""Write Gemini's Rating / Sentiment back into the Universal Sheet, so the next run (any client) reads them from
the sheet and sends no request for those stocks.

The sheet's CMP and Market Cap are formulas whose last values are stored inside the file; saving it with
openpyxl would drop those values (every price would read as missing). So this edits the sheet XML directly:
only the Rating, Sentiment, Rating Source and Rating Reason cells of the given rows change (plus the two
header cells the first time); every other byte of the workbook is copied unchanged. A backup is kept in
templates/backups/ (last 5) before each write.

Only cells that are blank or were written by the tool earlier ("AI ..." in Rating Source) are written - a value
typed by CREST is never overwritten."""
from __future__ import annotations

import os
import re
import shutil
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape

from openpyxl.utils import column_index_from_string, get_column_letter

KEEP_BACKUPS = 5


def ai_tag(model: str, period: str, basis: str) -> str:
    return f"AI · {model} · {period} · {basis}"


def tag_basis(tag) -> str | None:
    """The basis hash from an 'AI · model · period · basis' tag, or None when the cell is not an AI tag."""
    s = str(tag or "").strip()
    if not s.upper().startswith("AI"):
        return None
    return s.split("·")[-1].strip() if "·" in s else ""


def _sheet_part(z: zipfile.ZipFile, sheet_name: str) -> str:
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    m = re.search(rf'<sheet\b[^>]*\bname="{re.escape(escape(sheet_name, {chr(34): "&quot;"}))}"[^>]*\br:id="([^"]+)"', wb) \
        or re.search(r'<sheet\b[^>]*\br:id="([^"]+)"', wb)
    rid = m.group(1)
    t = re.search(rf'<Relationship\b[^>]*\bId="{re.escape(rid)}"[^>]*\bTarget="([^"]+)"', rels) \
        or re.search(rf'<Relationship\b[^>]*\bTarget="([^"]+)"[^>]*\bId="{re.escape(rid)}"', rels)
    target = t.group(1).lstrip("/")
    return target if target.startswith("xl/") else "xl/" + target


_CELL = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)


def _set_cells(row_xml: str, rown: int, cells: dict) -> str:
    """cells: {column letter: text}. Replace or insert inline-string cells, keeping column order and styles."""
    m = re.match(r'(<row\b[^>]*?)(/>|>(.*)</row>)$', row_xml, re.S)
    head, inner = m.group(1), (m.group(3) or "")
    existing = [(column_index_from_string(c.group(1)), c.group(0), c.group(3)) for c in _CELL.finditer(inner)]
    by_col = {ci: (xml, attrs) for ci, xml, attrs in existing}
    style = next((re.search(r'\bs="(\d+)"', a).group(1) for _, _, a in existing if re.search(r'\bs="(\d+)"', a)), None)
    for col, text in cells.items():
        ci = column_index_from_string(col)
        s_attr = ""
        if ci in by_col:
            sm = re.search(r'\bs="(\d+)"', by_col[ci][1])
            s_attr = f' s="{sm.group(1)}"' if sm else ""
        elif style:
            s_attr = f' s="{style}"'
        by_col[ci] = (f'<c r="{col}{rown}"{s_attr} t="inlineStr"><is><t xml:space="preserve">{escape(str(text))}</t></is></c>', "")
    maxc = max(by_col)
    head = re.sub(r'\bspans="(\d+):(\d+)"', lambda s: f'spans="{s.group(1)}:{max(int(s.group(2)), maxc)}"', head)
    return head + ">" + "".join(by_col[c][0] for c in sorted(by_col)) + "</row>"


def write_cells(xlsx: Path, sheet_name: str, updates: dict) -> int:
    """updates: {row number: {column letter: text}}. Returns the number of rows changed.
    Raises PermissionError when the file is open in Excel."""
    if not updates:
        return 0
    xlsx = Path(xlsx)
    with zipfile.ZipFile(xlsx) as z:
        part = _sheet_part(z, sheet_name)
        xml = z.read(part).decode("utf-8")
        infos = z.infolist()
        blobs = {i.filename: z.read(i.filename) for i in infos if i.filename != part}

    changed = 0
    for rown in sorted(updates):
        m = re.search(rf'<row r="{rown}"(?:\s[^>]*)?(?:/>|>.*?</row>)', xml, re.S)
        if not m:
            continue
        new = _set_cells(m.group(0), rown, updates[rown])
        xml = xml[:m.start()] + new + xml[m.end():]
        changed += 1
    # the <dimension> must cover any new columns
    maxcol = max(column_index_from_string(c) for u in updates.values() for c in u)
    xml = re.sub(r'<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/>',
                 lambda d: f'<dimension ref="{d.group(1)}{d.group(2)}:'
                           f'{get_column_letter(max(column_index_from_string(d.group(3)), maxcol))}{d.group(4)}"/>',
                 xml, count=1)

    _backup(xlsx)
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=xlsx.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, "w") as out:
            for i in infos:
                data = xml.encode("utf-8") if i.filename == part else blobs[i.filename]
                out.writestr(i, data, compress_type=i.compress_type)
        os.replace(tmp, xlsx)  # PermissionError if Excel has the file open
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return changed


def _backup(xlsx: Path):
    d = xlsx.parent / "backups"
    d.mkdir(exist_ok=True)
    shutil.copy2(xlsx, d / f"{xlsx.stem}-{datetime.now():%Y%m%d-%H%M%S}{xlsx.suffix}")
    old = sorted(d.glob(f"{xlsx.stem}-*{xlsx.suffix}"))
    for f in old[:-KEEP_BACKUPS]:
        try:
            f.unlink()
        except OSError:
            pass

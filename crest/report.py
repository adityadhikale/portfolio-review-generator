"""report.md writer."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

ICON = {"PASS": "PASS", "WARN": "WARN", "FAIL": "**FAIL**"}


def _esc(s) -> str:
    return str(s).replace("|", "\\|").replace("\n", " ")


def write_report(res, path: Path):
    i = res.info
    L = [f"# CREST portfolio run - {i.get('client', '')} - {i.get('period', '')}", "",
         f"**Status: {res.status}**  ", f"Run started {i.get('started', '')}, report written "
         f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", ""]
    ps = i.get("price_source")
    L += ["## Pricing", "",
          f"**price_source = `{ps}`** - every CMP, current value, P&L and weight in this run uses "
          + ("the **CMP column of the Universal Sheet**." if ps == "universal_cmp" else
             "the **LTP column of the broker holdings file**.") + " Sources are never mixed within a run.", ""]
    L += ["## Inputs", "", "| Item | Value |", "|---|---|"]
    for k, lab in (("holdings", "Holdings file"), ("broker", "Broker config detected"), ("header_row", "Header row found"),
                   ("asof_in_file", "As-on date in holdings file"), ("asof", "As-on date (--asof)"),
                   ("universal", "Universal Sheet"), ("universal_sheet", "Universal sheet used"),
                   ("universal_rows", "Universal rows"), ("template", "Template"), ("narrative_path", "narrative.json looked for at"),
                   ("libreoffice", "LibreOffice"), ("excel_cached_values", "Excel cached values")):
        if k in i:
            L.append(f"| {lab} | {_esc(i[k])} |")
    if "column_map" in i:
        L.append(f"| Columns used | {_esc(', '.join(f'{k} = {v}' for k, v in i['column_map'].items()))} |")
    L.append("")
    if "holdings_parsed" in i:
        mb = i.get("matched_by", {})
        n_exc = sum(1 for e in res.exceptions if e.kind in ("Unmatched", "Ambiguous"))
        L += ["## Matching", "", f"- Holdings parsed (after merging duplicates, dropping zero-qty/footer rows): **{i['holdings_parsed']}**",
              f"- Matched to Universal Sheet: **{sum(mb.values())}** (by ISIN {mb.get('isin', 0)}, by symbol {mb.get('symbol', 0)}, by ticker code "
              f"{mb.get('ticker code', 0)}, by exact name {mb.get('name', 0)})",
              f"- Unmatched / ambiguous: **{n_exc}**", ""]
    L += ["## Validation and QA checks", "", "| Check | Result | Detail |", "|---|---|---|"]
    for c in res.checks:
        L.append(f"| {_esc(c.name)} | {ICON.get(c.status, c.status)} | {_esc(c.detail)} |")
    L.append("")
    for c in res.checks:
        if c.items and c.status != "PASS":
            L += [f"<details><summary>{_esc(c.name)} - {len(c.items)} item(s)</summary>", ""]
            for it in c.items:
                if isinstance(it, dict):
                    it = ", ".join(f"{k}: {v}" for k, v in it.items())
                L.append(f"- {_esc(it)}")
            L += ["", "</details>", ""]
    if res.notes:
        L += ["## Notes and warnings", ""] + [f"- {_esc(n)}" for n in res.notes] + [""]
    L += ["## Exceptions", ""]
    if res.exceptions:
        L += ["| Severity | Type | Symbol | Detail |", "|---|---|---|---|"]
        order = {"FAIL": 0, "WARN": 1, "INFO": 2}
        for e in sorted(res.exceptions, key=lambda e: order.get(e.severity, 3)):
            L.append(f"| {e.severity} | {_esc(e.kind)} | {_esc(e.symbol)} | {_esc(e.detail)} |")
    else:
        L.append("None.")
    L.append("")
    n = res.narrative
    L += ["## Narrative (slides 2-5 and 15)", ""]
    if not n:
        L.append("Not reached.")
    elif not n.get("present"):
        L.append(f"narrative.json not found - narrative fields are marked **PENDING NARRATIVE** in the deck. "
                 f"Draft it from `outputs/narrative_input.json` (blank shape in `outputs/narrative_template.json`) and "
                 f"save it as `{i.get('narrative_path')}`, then re-run.")
    else:
        L.append("narrative.json loaded.")
        for k, lab in (("unverified_numbers", "Figures not found in narrative_input.json"), ("style", "Style issues"),
                       ("problems", "Structure problems")):
            if n.get(k):
                L += ["", f"**{lab}:**"] + [f"- {_esc(x if isinstance(x, str) else ': '.join(x))}" for x in n[k]]
    if n.get("pending_fields"):
        L += ["", "**Fields marked PENDING NARRATIVE:** " + ", ".join(f"`{p}`" for p in n["pending_fields"])]
    if res.qa.get("pending"):
        L += ["", "Slides containing PENDING NARRATIVE: " + ", ".join(res.qa["pending"])]
    L.append("")
    if i.get("table_slides"):
        L += ["## Results-table slides", "", "| Slide | Group | Part | Holdings |", "|---|---|---|---|"]
        for t in i["table_slides"]:
            L.append(f"| {t['slide']} | {_esc(t['group'])} | {t['part']} | {', '.join(t['symbols'])} |")
        L.append("")
    if res.outputs:
        L += ["## Outputs", ""] + [f"- {k}: `{v}`" for k, v in res.outputs.items()] + [""]
    path.write_text("\n".join(L), encoding="utf-8")

"""report.md writer."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

ICON = {"PASS": "PASS", "WARN": "WARN", "FAIL": "**FAIL**"}
STEP_TITLES = {"portfolio": "Step A: Portfolio Excel", "deck": "Step B: Deck", "both": "Portfolio Excel and deck"}


def _esc(s) -> str:
    return str(s).replace("|", "\\|").replace("\n", " ")


def write_report(res, path: Path):
    i = res.info
    what = STEP_TITLES.get(i.get("step"), "")
    L = [f"# CREST portfolio run - {i.get('client', '')} - {i.get('period', '')}" + (f" - {what}" if what else ""), "",
         f"**Status: {res.status}**  ", f"Run started {i.get('started', '')}, report written "
         f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", ""]
    L += ["## Pricing", "", "Every CMP, current value, P&L and weight in this run uses the **CMP column of the "
          "Universal Sheet**, read at the time of this run. Neither the holdings template nor the Portfolio Excel "
          "supplies prices.", ""]
    L += ["## Inputs", "", "| Item | Value |", "|---|---|"]
    for k, lab in (("holdings", "Holdings file"), ("holdings_kind", "Holdings file type"),
                   ("holdings_sheet", "Holdings sheet used"), ("client", "Client name"), ("asof", "As-on date"),
                   ("universal", "Universal Sheet"), ("universal_read_at", "Universal Sheet read at"),
                   ("universal_snapshot", "Data snapshot used"), ("universal_sheet", "Universal sheet used"),
                   ("universal_rows", "Universal rows"), ("template", "Template"), ("narrative_path", "narrative.json looked for at"),
                   ("libreoffice", "LibreOffice"), ("excel_cached_values", "Excel cached values")):
        if k in i:
            L.append(f"| {lab} | {_esc(i[k])} |")
    L.append("")
    if "holdings_parsed" in i:
        mb = i.get("matched_by", {})
        n_exc = sum(1 for e in res.exceptions if e.kind in ("Unmatched", "Ambiguous"))
        L += ["## Matching", "", f"- Holdings read (after merging rows for the same stock): **{i['holdings_parsed']}**",
              f"- Matched to Universal Sheet: **{sum(mb.values())}** (by ISIN {mb.get('isin', 0)}, by symbol {mb.get('symbol', 0)}, by ticker code "
              f"{mb.get('ticker code', 0)}, by exact name {mb.get('name', 0)})",
              f"- Unmatched / ambiguous: **{n_exc}**", ""]
    steps = list(dict.fromkeys(getattr(c, "step", "") for c in res.checks))
    for st in steps:
        L += [f"## Checks: {st}" if st else "## Validation and QA checks", "", "| Check | Result | Detail |",
              "|---|---|---|"]
        for c in res.checks:
            if getattr(c, "step", "") == st:
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
    if i.get("step") != "portfolio":  # a Portfolio-only run has no deck or narrative
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

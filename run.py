r"""CREST client portfolio Excel + deck generator.

HOW TO USE
  1. Copy each client's holdings into the CREST holdings template (templates\Holdings_Template.xlsx):
     client name in C2, as-on date in C3, holdings from row 6. Save it in the  input\  folder.
  2. Run:   python run.py
  3. Collect the results from  output\<client>\<period>\

Everything else (period, Universal Sheet, deck template) is found automatically.
Optional overrides: python run.py --holdings file.xlsx --period Q2FY27
"""
from __future__ import annotations

import argparse
import sys
from collections import Counter
from pathlib import Path

from crest.pipeline import run
from crest.util import ROOT, PipelineError

INPUT_DIR = ROOT / "input"
OUTPUT_DIR = ROOT / "output"
DEFAULT_UNIVERSAL = ROOT / "templates" / "universal-sheet.xlsx"
DEFAULT_TEMPLATE = ROOT / "templates" / "Deck_-_template.pptx"


def input_files() -> list[Path]:
    return sorted(p for p in INPUT_DIR.glob("*")
                  if p.suffix.lower() in (".xlsx", ".xlsm") and not p.name.startswith("~$"))


def period_from_universal(path: Path) -> str:
    """Most common 'Q Result Period' in the Universal Sheet, e.g. 'Q2 FY27' -> 'Q2FY27'."""
    from crest.universal import result_periods
    c = Counter(str(v or "").replace(" ", "").upper() for v in result_periods(path))
    c.pop("", None)
    if not c:
        raise PipelineError("The Universal Sheet has no 'Q Result Period' values, so the period is unknown. "
                            "Fill that column (e.g. Q2 FY27).")
    return c.most_common(1)[0][0]


def run_one(holdings, period, a) -> bool:
    print(f"\n=== {holdings.name} ===")
    try:
        res = run(period=period, universal=a.universal, holdings=holdings, template=a.template,
                  narrative=a.narrative, runs_dir=a.runs_dir, settings_path=a.settings,
                  render_png=not a.no_render, publish_dir=a.output_dir)
    except PipelineError as e:
        print(f"  NOT DONE: {e}")
        return False
    ok = res.status.startswith("PASSED")
    print(f"  {res.info.get('client', '')} as on {res.info.get('asof', '')}: {'DONE' if ok else 'NOT DONE'} - {res.status}")
    for c in res.checks:
        if c.status != "PASS":
            print(f"  [{c.status}] {c.name}: {c.detail}")
    for n in res.notes:
        print(f"  {n}")
    if res.outputs.get("published_to"):
        print(f"  Results: {res.outputs['published_to']}")
    return ok


def main(argv=None) -> int:
    # Check details contain symbols such as the Greek sigma; a redirected Windows console (cp1252) can't print them.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")
    ap = argparse.ArgumentParser(description="Save filled holdings templates in input\\ and run: python run.py")
    ap.add_argument("--holdings", help="Only for a single file: path to a filled holdings template "
                    "(default: every .xlsx in input\\)")
    ap.add_argument("--period", help="e.g. Q2FY27 (default: from the Universal Sheet)")
    ap.add_argument("--universal", default=str(DEFAULT_UNIVERSAL))
    ap.add_argument("--template", default=str(DEFAULT_TEMPLATE))
    ap.add_argument("--narrative")
    ap.add_argument("--runs-dir")
    ap.add_argument("--output-dir", default=str(OUTPUT_DIR))
    ap.add_argument("--settings")
    ap.add_argument("--no-render", action="store_true")
    a = ap.parse_args(argv)
    try:
        period = a.period or period_from_universal(Path(a.universal))
    except PipelineError as e:
        print(f"NOT DONE: {e}")
        return 2
    if a.holdings:
        jobs = [Path(a.holdings)]
    else:
        jobs = input_files()
        if not jobs:
            print(f"Nothing to do: save each client's filled holdings template (.xlsx) in {INPUT_DIR} and run again. "
                  f"The blank template is {ROOT / 'templates' / 'Holdings_Template.xlsx'}.")
            return 2
    print(f"Period {period} - {len(jobs)} file(s)")
    results = [run_one(h, period, a) for h in jobs]
    print(f"\nFinished: {sum(results)} done, {len(results) - sum(results)} not done. Results are in {a.output_dir}")
    return 0 if all(results) else 2


if __name__ == "__main__":
    sys.exit(main())

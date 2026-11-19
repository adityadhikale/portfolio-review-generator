r"""CREST client portfolio Excel + deck generator.

HOW TO USE
  1. Put each client's broker file in the  input\  folder.  The file name is the client name,
     e.g.  input\Sharma Family.xlsx  ->  client "Sharma Family".
  2. Run:   python run.py
  3. Collect the results from  output\<client>\<period>\

Everything else (period, as-on date, Universal Sheet, template) is found automatically.
Optional overrides for one file: python run.py --client "Name" --holdings file.xlsx --period Q2FY27 --asof dd/mm/yyyy
"""
from __future__ import annotations

import argparse
import re
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
                  if p.suffix.lower() in (".xlsx", ".xlsm", ".csv") and not p.name.startswith("~$"))


def client_from_filename(p: Path) -> str:
    """'Sharma Family Portfolio.xlsx' -> 'Sharma Family'; 'Client_A_holdings.csv' -> 'Client A'."""
    s = re.sub(r"[_]+", " ", p.stem).strip()
    s = re.sub(r"\s*\b(portfolio|holdings?|equity|statement)\b\s*", " ", s, flags=re.I).strip()
    return s or p.stem


def period_from_universal(path: Path) -> str:
    """Most common 'Q Result Period' in the Universal Sheet, e.g. 'Q2 FY27' -> 'Q2FY27'."""
    from crest.universal import load_universe
    c = Counter(str(u.values.get("result_period") or "").replace(" ", "").upper() for u in load_universe(path).rows)
    c.pop("", None)
    if not c:
        raise PipelineError("The Universal Sheet has no 'Q Result Period' values, so the period is unknown. "
                            "Fill that column (e.g. Q2 FY27).")
    return c.most_common(1)[0][0]


def run_one(client, holdings, period, a) -> bool:
    print(f"\n=== {client}  ({holdings.name}) ===")
    try:
        res = run(client=client, period=period, asof=a.asof, universal=a.universal, holdings=holdings,
                  template=a.template, narrative=a.narrative, broker=a.broker, price_source=a.price_source,
                  runs_dir=a.runs_dir, settings_path=a.settings, render_png=not a.no_render,
                  publish_dir=a.output_dir)
    except PipelineError as e:
        print(f"  NOT DONE: {e}")
        return False
    ok = res.status.startswith("PASSED")
    print(f"  {'DONE' if ok else 'NOT DONE'} - {res.status}")
    for c in res.checks:
        if c.status != "PASS":
            print(f"  [{c.status}] {c.name}: {c.detail}")
    for n in res.notes:
        print(f"  {n}")
    if res.outputs.get("published_to"):
        print(f"  Results: {res.outputs['published_to']}")
    return ok


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Put client files in input\\ and run: python run.py")
    ap.add_argument("--client", help="Only for a single file: client name (default: taken from the file name)")
    ap.add_argument("--holdings", help="Only for a single file: path to the holdings file (default: every file in input\\)")
    ap.add_argument("--period", help="e.g. Q2FY27 (default: from the Universal Sheet)")
    ap.add_argument("--asof", help="dd/mm/yyyy (default: from the holdings file)")
    ap.add_argument("--universal", default=str(DEFAULT_UNIVERSAL))
    ap.add_argument("--template", default=str(DEFAULT_TEMPLATE))
    ap.add_argument("--narrative")
    ap.add_argument("--broker")
    ap.add_argument("--price-source", choices=["universal_cmp", "broker_ltp"])
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
        jobs = [(a.client or client_from_filename(Path(a.holdings)), Path(a.holdings))]
    else:
        files = input_files()
        if not files:
            print(f"Nothing to do: put the client's broker file (.xlsx or .csv) in {INPUT_DIR} and run again.")
            return 2
        jobs = [(a.client if (a.client and len(files) == 1) else client_from_filename(f), f) for f in files]
    print(f"Period {period} - {len(jobs)} file(s)")
    results = [run_one(c, h, period, a) for c, h in jobs]
    print(f"\nFinished: {sum(results)} done, {len(results) - sum(results)} not done. Results are in {a.output_dir}")
    return 0 if all(results) else 2


if __name__ == "__main__":
    sys.exit(main())

"""One run: shared setup, then Step A (Portfolio Excel), Step B (deck), or both.

    step="portfolio"  holdings template -> Portfolio Excel
    step="deck"       Portfolio Excel (or holdings template) -> PPT + PDF
    step="both"       holdings template -> Portfolio Excel -> PPT + PDF, from one read of the Universal Sheet, so
                      the Excel and the deck show identical numbers
"""
from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

from .holdings import read_template
from .portfolio_reader import file_kind, read_portfolio_excel
from . import validate as V
from .shared import Ctx, RunResult, prepare
from .sheet_source import LocalSheet, UniversalSheet
from .step_deck import build_deck
from .step_portfolio import build_portfolio
from .util import CONFIG_DIR, ROOT, PipelineError, load_json, parse_period, safe_filename

STEPS = ("portfolio", "deck", "both")
STEP_NAMES = {"shared": "Shared setup", "portfolio": "Step A: Portfolio", "deck": "Step B: Deck"}
# Result files each step writes to output/<client>/<period>/ (a step leaves the other step's files alone)
STEP_SUFFIXES = {"portfolio": {".xlsx"}, "deck": {".pptx", ".pdf"}}


def _archive_dir(base: Path, period: str, client: str) -> Path:
    root = base / safe_filename(period) / safe_filename(client)
    if (root / "outputs").exists() or (root / "inputs").exists():
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        hist = root / "_history" / stamp
        hist.mkdir(parents=True, exist_ok=True)
        for sub in ("inputs", "outputs", "report.md"):
            if (root / sub).exists():
                shutil.move(str(root / sub), str(hist / sub))
    (root / "inputs").mkdir(parents=True, exist_ok=True)
    (root / "outputs").mkdir(parents=True, exist_ok=True)
    return root


def read_holdings(path: Path, step: str):
    """-> (HoldingsFile, 'template' | 'portfolio'). Steps portfolio and both need a holdings template."""
    kind = file_kind(path)
    if kind == "portfolio" and step != "deck":
        raise PipelineError(f"{Path(path).name} is a Portfolio Excel: it can only be used to make the deck "
                            "(step 'deck'). Making a Portfolio Excel needs a filled holdings template.")
    return (read_portfolio_excel(path) if kind == "portfolio" else read_template(path)), kind


def run(*, period, universal, holdings, template=None, narrative=None, runs_dir=None, settings_path=None,
        render_png=True, publish_dir=None, step="both", progress=None) -> RunResult:
    """holdings: a filled holdings template, or (step 'deck' only) a Portfolio Excel built by Step A. The client
    name and the as-on date are read from it; prices always come from the Universal Sheet.
    universal: a UniversalSheet (crest/sheet_source.py: Google Sheets with snapshot fallback; pass the same one to
    every run of a batch so all of them use one read), or a path to a local .xlsx.
    progress: optional progress(stage) callback; stages in order: sheet, match, research, excel, check, ai, deck,
    qa (only those the chosen step uses)."""
    if step not in STEPS:
        raise PipelineError(f"step must be one of {', '.join(STEPS)}, got '{step}'")
    settings = load_json(Path(settings_path) if settings_path else CONFIG_DIR / "settings.json")
    per = parse_period(period)
    if not isinstance(universal, UniversalSheet):
        if not Path(universal).exists():
            raise PipelineError(f"Input file not found: {universal}")
        universal = UniversalSheet(LocalSheet(universal), runs_dir)
    holdings = Path(holdings)
    template = Path(template) if template else None
    if step != "portfolio" and template is None:
        raise PipelineError("The deck template is needed to make the deck.")
    used = [holdings] + ([template] if step != "portfolio" else [])
    for p in used:
        if not p.exists():
            raise PipelineError(f"Input file not found: {p}")
    hf, kind = read_holdings(holdings, step)
    res = RunResult()
    run_dir = _archive_dir(Path(runs_dir) if runs_dir else ROOT / "runs", per.compact, hf.client)
    res.run_dir = run_dir
    out = run_dir / "outputs"
    for p in used:
        shutil.copy2(p, run_dir / "inputs" / p.name)
    narrative_path = Path(narrative) if narrative else run_dir / "narrative.json"
    res.info.update(step=step, client=hf.client, period=per.label, asof=hf.asof, universal=universal.label,
                    holdings=holdings.name,
                    holdings_kind="Portfolio Excel" if kind == "portfolio" else "holdings template",
                    started=datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
    if step != "portfolio":
        res.info.update(template=template.name, narrative_path=str(narrative_path))
    ctx = Ctx(res=res, step=step, settings=settings, per=per, client=hf.client, asof=hf.asof, universal=universal,
              template=template, narrative_path=narrative_path, hf=hf, out=out, render_png=render_png,
              progress=progress)
    current = "shared"
    try:
        ctx.stage("sheet")
        _universal_info(res, universal, per)
        prepare(ctx)
        _tag(res, current)
        excel_vals = None
        if step in ("portfolio", "both"):
            current = "portfolio"
            excel_vals = build_portfolio(ctx)
            _tag(res, current)
        if step in ("deck", "both"):
            current = "deck"
            build_deck(ctx, excel_vals)     # sets res.status
            _tag(res, current)
        else:
            res.status = "PASSED WITH WARNINGS" if any(c.status == "WARN" for c in res.checks) else "PASSED"
    except PipelineError as e:
        _tag(res, current)
        res.status = "FAILED"
        res.notes.append(f"STOPPED ({STEP_NAMES[current]}): {e}")
    finally:
        from .report import write_report
        write_report(res, run_dir / "report.md")
        if publish_dir and _publish(res, Path(publish_dir) / safe_filename(hf.client) / per.compact, step):
            write_report(res, run_dir / "report.md")  # add the locked-file note
            shutil.copy2(run_dir / "report.md", res.outputs["published_to"] / "report.md")
    return res


def _universal_info(res, us: UniversalSheet, per):
    """Read the Universal Sheet (once per batch), save the snapshot this run used, and say where the data came from."""
    d = us.data()
    snap = us.save_snapshot(per.compact)
    res.info.update(universal=d.source, universal_read_at=d.read_at, universal_snapshot=str(snap))
    for n in us.notes():
        if n not in res.notes:
            res.notes.append(n)
    if d.warning:
        res.checks.append(V.Check("Universal Sheet read live from Google", "WARN", d.warning))
    elif d.live:
        res.checks.append(V.Check("Universal Sheet read live from Google", "PASS", f"{d.source}, read {d.read_at}"))


def _tag(res, step: str):
    """Mark the checks added so far that have no step yet (the report groups checks by step)."""
    for c in res.checks:
        if not c.step:
            c.step = STEP_NAMES[step]


def _publish(res, dest: Path, step: str = "both") -> bool:
    """Copy the results (Excel, deck, PDF, report, ratings) to output/<client>/<period>/. A file that is open
    in Excel / PowerPoint cannot be replaced: the new version is saved next to it as '<name> (new).<ext>' and
    a note says so, instead of the run crashing. Returns True when such a note was added.
    Old results this run did not produce are removed only if they are this step's kind of file, so a deck run
    leaves the Portfolio Excel in place and a Portfolio run leaves the deck."""
    dest.mkdir(parents=True, exist_ok=True)
    files = [Path(p) for k in ("excel", "deck", "pdf", "rating_preview", "ai_ratings")
             if (p := res.outputs.get(k)) and Path(p).exists()] + [res.run_dir / "report.md"]
    keep = {f.name for f in files}
    mine = None if step == "both" else STEP_SUFFIXES[step]
    for old in dest.iterdir():  # results from an earlier run that this run did not produce
        if (old.is_file() and old.name not in keep and " (new)" not in old.stem and not old.name.startswith("~$")
                and (mine is None or old.suffix.lower() in mine)):
            try:
                old.unlink()
            except OSError:
                pass
    locked = []
    for f in files:
        try:
            shutil.copy2(f, dest / f.name)
            stale = dest / f"{f.stem} (new){f.suffix}"
            if stale.exists():
                try:
                    stale.unlink()
                except OSError:
                    pass
        except PermissionError:
            alt = dest / f"{f.stem} (new){f.suffix}"
            shutil.copy2(f, alt)
            locked.append(f"{f.name} -> {alt.name}")
    res.outputs["published_to"] = dest
    if locked:
        res.notes.append("WARN: file(s) open in another program could not be replaced in the output folder; the new "
                         "version is saved next to them: " + "; ".join(locked) + ". Close them and re-run to tidy up.")
    return bool(locked)

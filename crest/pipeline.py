"""End-to-end run: holdings + Universal Sheet + template -> Excel, narrative_input.json, deck, report.md."""
from __future__ import annotations

import json
import shutil
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from openpyxl.utils import get_column_letter

from . import deck_data as dd
from . import validate as V
from .ai_ratings import RatingCache, rating_basis
from .ai_review import request_review
from .deck_fill import fill_deck
from .excel_build import build_workbook
from .holdings import read_template
from .narrative_api import GeminiClient, resolve_ai
from .match import match_holdings
from .narrative import SCHEMA, ManualProvider, blank_template, check_numbers, resolve_provider_name, style_issues
from .office import find_soffice, recalculate_copy
from .qa import render, run_qa
from .ratings import RatingRules
from .sheet_writeback import ai_tag, tag_basis, write_cells
from .universal import load_universe, value_number
from .util import CONFIG_DIR, ROOT, PipelineError, load_env, load_json, parse_period, safe_filename


@dataclass
class Exc:
    severity: str
    kind: str
    symbol: str
    detail: str
    holding: object = None


@dataclass
class RunResult:
    status: str = "RUNNING"
    run_dir: Path | None = None
    checks: list = field(default_factory=list)
    exceptions: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    outputs: dict = field(default_factory=dict)
    narrative: dict = field(default_factory=dict)
    qa: dict = field(default_factory=dict)
    info: dict = field(default_factory=dict)


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


def run(*, period, universal, holdings, template, narrative=None, runs_dir=None, settings_path=None,
        render_png=True, publish_dir=None) -> RunResult:
    """holdings: a filled CREST holdings template (templates/Holdings_Template.xlsx). The client name and the
    as-on date are read from it; prices always come from the Universal Sheet."""
    settings = load_json(Path(settings_path) if settings_path else CONFIG_DIR / "settings.json")
    per = parse_period(period)
    universal, holdings, template = Path(universal), Path(holdings), Path(template)
    for p in (universal, holdings, template):
        if not p.exists():
            raise PipelineError(f"Input file not found: {p}")
    hf = read_template(holdings)
    client, asof = hf.client, hf.asof
    res = RunResult()
    run_dir = _archive_dir(Path(runs_dir) if runs_dir else ROOT / "runs", per.compact, client)
    res.run_dir = run_dir
    out = run_dir / "outputs"
    for p in (universal, holdings, template):
        shutil.copy2(p, run_dir / "inputs" / p.name)
    narrative_path = Path(narrative) if narrative else run_dir / "narrative.json"
    res.info.update(client=client, period=per.label, asof=asof, universal=universal.name,
                    holdings=holdings.name, template=template.name, narrative_path=str(narrative_path),
                    started=datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
    try:
        _run_inner(res, settings, per, client, asof, universal, holdings, template, narrative_path, hf, out,
                   render_png)
    except PipelineError as e:
        res.status = "FAILED"
        res.notes.append(f"STOPPED: {e}")
    finally:
        from .report import write_report
        write_report(res, run_dir / "report.md")
        if publish_dir and _publish(res, Path(publish_dir) / safe_filename(client) / per.compact):
            write_report(res, run_dir / "report.md")  # add the locked-file note
            shutil.copy2(run_dir / "report.md", res.outputs["published_to"] / "report.md")
    return res


def _publish(res, dest: Path) -> bool:
    """Copy the results (Excel, deck, PDF, report, ratings) to output/<client>/<period>/. A file that is open
    in Excel / PowerPoint cannot be replaced: the new version is saved next to it as '<name> (new).<ext>' and
    a note says so, instead of the run crashing. Returns True when such a note was added."""
    dest.mkdir(parents=True, exist_ok=True)
    files = [Path(p) for k in ("excel", "deck", "pdf", "rating_preview", "ai_ratings")
             if (p := res.outputs.get(k)) and Path(p).exists()] + [res.run_dir / "report.md"]
    keep = {f.name for f in files}
    for old in dest.iterdir():  # results from an earlier run that this run did not produce
        if old.is_file() and old.name not in keep and " (new)" not in old.stem and not old.name.startswith("~$"):
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


def _run_inner(res, settings, per, client, asof, universal, holdings, template, narrative_path, bp, out, render_png):
    groups_cfg = load_json(CONFIG_DIR / "sector_groups.json")
    overrides = load_json(CONFIG_DIR / "name_overrides.json").get("overrides", {})

    # 1. holdings template (already read and checked by run())
    res.info.update(holdings_sheet=bp.sheet, holdings_parsed=len(bp.holdings))
    if Path(narrative_path).exists():
        shutil.copy2(narrative_path, res.run_dir / "inputs" / "narrative.json")
    res.notes += bp.notes

    # 2. universal + match
    uni = load_universe(universal)
    res.info.update(universal_rows=len(uni.rows), universal_header_row=uni.header_row, universal_sheet=uni.sheet)
    for w in uni.warnings:
        res.exceptions.append(Exc("WARN", "Universal duplicate", "", w))
    matched, mexc = match_holdings(bp.holdings, uni)
    for e in mexc:
        res.exceptions.append(Exc("WARN", e.kind, e.symbol, e.detail, e.holding))
    res.info["matched_by"] = {k: sum(1 for m in matched if m.method == k) for k in ("isin", "symbol", "ticker code", "name")}

    # deck group per stock
    valid_groups = set(groups_cfg["group_order"])

    def deck_group_of(u):
        if u.is_etf:
            return groups_cfg["etf_group"]
        g = str(u.values.get("deck_group") or "").strip()
        if g:
            return g if g in valid_groups else groups_cfg["unmapped_group"]
        so = groups_cfg.get("symbol_overrides", {})
        if u.key in so:
            return so[u.key]
        return groups_cfg["sector_to_group"].get(str(u.values.get("sector") or "").strip(), groups_cfg["unmapped_group"])

    # result period must match the run period (results for another quarter would be mislabelled)
    wrong_period = [f"{m.holding.raw_symbol}: '{m.uni.values.get('result_period')}'" for m in matched
                    if str(m.uni.values.get("result_period") or "").strip()
                    and str(m.uni.values["result_period"]).replace(" ", "").upper() != per.compact.upper()]
    res.checks.append(V.Check(f"Universal 'Q Result Period' matches --period ({per.label})",
                              "FAIL" if wrong_period else "PASS",
                              f"{len(wrong_period)} row(s) hold results for another quarter" if wrong_period
                              else "All filled rows match (or column absent)", wrong_period))

    # Rating / Sentiment: Universal Sheet value > Gemini (when .env has a key) > approved rule > blank.
    # Views are never set here: they are CREST's own call, from the Universal Sheet only.
    # Saved Gemini ratings are applied here; stocks still unrated are rated later in the run's single Gemini
    # request (step 7), together with the narrative.
    env = load_env()
    rp = settings.get("rating_provider", "auto")
    ai_on = rp == "gemini" or (rp == "auto" and resolve_ai(env))
    rating_cache, ai_need, ai_to_rate = None, [], []
    if ai_on:
        _use_sheet_ai_ratings(res, matched)
        rating_cache, ai_need, ai_to_rate = _apply_saved_ai_ratings(res, matched)
    pending_ai = {t["symbol"] for t in ai_to_rate}
    rr = settings.get('rating_rules_file')
    rules = RatingRules(load_json(Path(rr) if Path(rr).is_absolute() else ROOT / rr) if rr else None)
    preview = []
    for m in matched:
        if ai_on:  # Gemini rated what it could (outputs/ai_ratings.md); the draft rules are not used
            m.uni.rating_source = "Universal Sheet" if m.uni.rating and getattr(m.uni, "rating_source", "") != "AI (Gemini)" \
                else getattr(m.uni, "rating_source", "")
            continue
        u = m.uni
        u.rating_source = "Universal Sheet" if u.rating else getattr(u, "rating_source", "")
        if u.is_etf:
            continue
        rt, se, why = rules.evaluate(u.values.get("result"))
        if not u.rating or not u.sentiment:
            preview.append((m.holding.raw_symbol, u.values.get("result") or "", rt or "-", se or "-", why))
        if rules.approved:
            if not u.rating and rt:
                u.rating, u.rating_source = rt, "rule"
            if not u.sentiment and u.rating:
                u.sentiment = rules.sent.get(u.rating, "")
    if preview:
        _write_rating_preview(out / "rating_preview.md", preview, rules.approved)
        res.outputs["rating_preview"] = out / "rating_preview.md"
        if not rules.approved:
            res.notes.append(f"WARN: {len(preview)} stock(s) have no Rating/Sentiment in the Universal Sheet. Draft rules in "
                             "config/rating_rules.json are NOT approved, so these show n/a. Proposed values: "
                             "outputs/rating_preview.md")
    no_view = [m.holding.raw_symbol for m in matched if not str(m.uni.values.get("views") or "").strip()]
    res.checks.append(V.simple("Views filled in the Universal Sheet", no_view))

    # 3. per-stock data issues
    priced = []
    for m in matched:
        u, h = m.uni, m.holding
        if not u.values.get("short_name") and h.symbol in overrides:
            u.values["short_name"] = overrides[h.symbol]
        for n in u.notes:
            res.exceptions.append(Exc("WARN" if "defaulted" in n or "not one of" in n else "INFO",
                                      "Ticker/data note", h.raw_symbol, n))
        if u.etf_inferred:
            res.exceptions.append(Exc("INFO", "ETF inferred", h.raw_symbol,
                                      f"'{u.values['name']}' treated as ETF (name/symbol); set Instrument Type=ETF to make it explicit. "
                                      "Market cap shows 'ETF'; excluded from rating counts and results tables."))
        if not u.is_etf:
            miss = [f for f, k in (("Result Line", "result"), ("Management Commentary", "commentary"))
                    if not (u.values.get(k) or "")]
            if u.key not in pending_ai:  # those are rated in step 7
                miss += [f for f, v in (("Rating", u.rating), ("Sentiment", u.sentiment)) if not v]
            if miss:
                res.exceptions.append(Exc("WARN", "Missing quarterly data", h.raw_symbol, f"Universal row {u.row}: no {', '.join(miss)}"))
            if value_number(u, "mcap") is None:
                res.exceptions.append(Exc("WARN", "Missing market cap", h.raw_symbol, f"Universal row {u.row}"))
            g = deck_group_of(u)
            if not str(u.values.get("sector") or "").strip() and not u.values.get("deck_group"):
                res.exceptions.append(Exc("WARN", "Missing sector", h.raw_symbol,
                                          f"{u.values['name']}: Sector is blank in the Universal Sheet (row {u.row}) - "
                                          "shown under 'Other Holdings' until it is filled"))
            elif g == groups_cfg["unmapped_group"]:
                res.exceptions.append(Exc("WARN", "Unmapped deck group", h.raw_symbol,
                                          f"Sector '{u.values.get('sector')}' / Deck Group '{u.values.get('deck_group') or ''}' "
                                          "not in config/sector_groups.json - shown under 'Other Holdings'"))
        price = value_number(u, "cmp")
        if price is None or price <= 0:
            res.exceptions.append(Exc("WARN", "Missing price", h.raw_symbol,
                                      f"No Universal CMP (cell value: {u.values.get('cmp')!r}); "
                                      "excluded from Portfolio table and deck - fix the price and re-run", h))
            continue
        priced.append((h.quantity * price, m))
    priced = [m for _, m in sorted(priced, key=lambda x: -x[0])]

    sector_chk = V.sector_sanity(matched, settings["sector_check"])
    for f in sector_chk.items:
        res.exceptions.append(Exc("WARN", "Implausible sector", f["symbol"],
                                  f"{f['name']}: Sector '{f['sector']}' (Universal row {f['universal_row']}); "
                                  f"name suggests {', '.join(f['expected'])}"))

    # 4. Excel
    xlsx = out / f"{client} Portfolio.xlsx"
    sev_order = {"FAIL": 0, "WARN": 1, "INFO": 2}
    exc_rows = sorted(res.exceptions, key=lambda e: sev_order.get(e.severity, 3))
    layout = build_workbook(xlsx, client=client, asof=asof, period_label=per.label, parse=bp,
                            priced=priced, snapshot_rows=matched, exception_rows=exc_rows, deck_group_of=deck_group_of)
    res.outputs["excel"] = xlsx

    soffice = find_soffice(settings.get("libreoffice_path", ""))
    res.info["libreoffice"] = soffice or "not found"
    py = dd.python_values(priced)
    if soffice:
        # The delivered workbook carries Google-Sheets GOOGLEFINANCE formulas (Market Cap, CMP) that LibreOffice
        # cannot evaluate (and would rewrite). So a check copy with this run's values in those two columns is
        # recalculated instead; every other formula is identical. The delivered file is left untouched.
        check = out / "_recalc" / f"check_{xlsx.name}"
        check.parent.mkdir(parents=True, exist_ok=True)
        build_workbook(check, client=client, asof=asof, period_label=per.label, parse=bp,
                       priced=priced, snapshot_rows=matched, exception_rows=exc_rows, deck_group_of=deck_group_of,
                       google_finance=False)
        rc = recalculate_copy(check, check, soffice)
        errors = dd.scan_errors(rc)
        excel_vals = dd.read_portfolio(rc, layout)
        res.checks.append(V.excel_errors(errors))
        res.checks.append(V.python_vs_excel(dd.compare(excel_vals, py)))
        res.info["excel_cached_values"] = ("no - Excel / Google Sheets calculate on open; Market Cap and CMP are "
                                           "GOOGLEFINANCE formulas (live in Google Sheets, this run's values elsewhere)")
    else:
        res.checks.append(V.Check("No Excel error values after recalculation", "WARN",
                                  "LibreOffice not found - workbook not recalculated; values computed in Python "
                                  "from the same inputs. Excel will calculate on open (fullCalcOnLoad)."))
        excel_vals = _py_as_excel(py, matched, priced, deck_group_of)

    # 5. validation gate
    # every template holding exactly once: in the Portfolio table or in Exceptions (a matched stock with no price
    # is also in Exceptions)
    all_h = bp.holdings
    accounted = {id(m.holding) for m in matched} | {id(e.holding) for e in res.exceptions if e.holding is not None}
    lost = [h.raw_symbol for h in all_h if id(h) not in accounted]
    res.checks.insert(0, V.Check("Every template holding is in the Portfolio or in Exceptions",
                                 "FAIL" if lost else "PASS",
                                 f"Not accounted for: {', '.join(lost)}" if lost else
                                 f"All {len(all_h)} holding(s) from {sum(len(h.source_rows) for h in all_h)} "
                                 "template row(s) accounted for"))
    in_table = {m.holding.symbol for m in priced}
    exc_inv = sum(h.quantity * h.avg_price for h in all_h if h.symbol not in in_table)
    res.checks.append(V.excel_ties(excel_vals["summary"]["invested"] or 0, exc_inv, all_h))
    res.checks.append(V.allocation(sum(float(r["alloc"] or 0) for r in excel_vals["rows"]), settings["allocation_tolerance_pct"]))
    unm = [f"{e.symbol}: {e.detail}" for e in res.exceptions if e.kind in ("Unmatched", "Ambiguous")]
    res.checks.append(V.simple("Unmatched / ambiguous symbols (listed in Exceptions)", unm,
                               fail=settings["unmatched_policy"] == "fail"))
    res.checks.append(V.simple("Missing price (listed in Exceptions)", [e.symbol for e in res.exceptions if e.kind == "Missing price"]))
    res.checks.append(sector_chk)
    res.checks.append(V.simple("Sector filled in the Universal Sheet",
                               [f"{e.symbol}: {e.detail.split(':')[0]}" for e in res.exceptions if e.kind == "Missing sector"]))
    etf_rows = [r for r in excel_vals["rows"] if excel_vals["snapshot"].get(r["key"], {}).get("is_etf") == "Yes"]
    bad_etf = [r["key"] for r in etf_rows if r["mcap"] != "ETF"]
    res.checks.append(V.Check("ETF market cap shown as 'ETF' (not #N/A)", "FAIL" if bad_etf else "PASS",
                              f"ETFs: {[r['key'] for r in etf_rows] or 'none'}", bad_etf))
    tick = [f"{e.symbol}: {e.detail}" for e in res.exceptions if e.kind == "Ticker/data note"]
    res.checks.append(V.simple("Ticker format NSE:/BSE: (BOM: converted, unknown exchange flagged)", tick, warn_if_any=any(
        "defaulted" in t or "not one of" in t for t in tick)))
    if any(c.status == "FAIL" for c in res.checks):
        raise PipelineError("Validation gate failed - see the checks table. Excel was written for inspection; deck not built.")

    # 6. deck data + narrative_input.json
    def deck_data_and_input():
        d = dd.build_deck_data(excel_vals, client=client, period=per, asof=asof,
                               group_order=groups_cfg["group_order"], unmapped_group=groups_cfg["unmapped_group"],
                               tail_pct=settings["deck"]["tail_weight_threshold_pct"], top_n=settings["deck"]["top_n"])
        n = {k: v for k, v in d.items() if not k.startswith("_")}
        n["narrative_schema"] = {k: {"description": ds, "max_chars": mx} for k, (ds, mx) in SCHEMA.items()}
        n["rules"] = ["Use only figures that appear in this file; every number in narrative.json is checked against it.",
                      "Write the brand as CREST. Do not start a heading with a number.",
                      "Keep each field within max_chars so it fits its box."]
        (out / "narrative_input.json").write_text(json.dumps(n, indent=2, ensure_ascii=False), encoding="utf-8")
        return d, n

    data, ni = deck_data_and_input()
    (out / "narrative_template.json").write_text(json.dumps(blank_template(), indent=2), encoding="utf-8")
    res.outputs["narrative_input"] = out / "narrative_input.json"

    # 7. narrative + Gemini. A saved narrative.json is always used as-is. Missing ratings and a missing narrative
    # are requested from Gemini together, in ONE request (none at all when nothing is missing).
    nar = ManualProvider(Path(narrative_path)).get(ni)
    want_nar = nar is None and resolve_provider_name(settings["narrative_provider"], env) == "gemini"
    res.info["narrative_provider"] = "saved narrative.json" if nar is not None else ("gemini" if want_nar else "none")
    res.info["gemini_requests"] = 0
    if ai_to_rate or want_nar:
        gc = GeminiClient(api_key=env.get("NARRATIVE_API_KEY", ""), model=env.get("NARRATIVE_MODEL") or None)
        what = " and ".join(x for x in ((f"rating {len(ai_to_rate)} stock(s)" if ai_to_rate else ""),
                                       ("writing the narrative" if want_nar else "")) if x)
        print(f"  Gemini: one request - {what} (models in order: {', '.join(gc.models)}) ...", flush=True)
        try:
            got, gnar, gnotes = request_review(gc, narrative_input=ni if want_nar else None, to_rate=ai_to_rate)
        except PipelineError as e:
            # Gemini down / overloaded / bad reply: still deliver the Excel and deck. Unrated stocks show n/a and the
            # narrative slides stay PENDING; the next run asks Gemini again for just what is still missing.
            got, gnar, gnotes = {}, None, [f"WARN: Gemini request failed, so the deck was built without it ({e}). "
                                           "Run again later; it will ask Gemini again for: " + what + "."]
            res.checks.append(V.Check("Gemini request (ratings / narrative)", "WARN", "failed - deck built without it",
                                      [str(e)]))
        if gc.last_bad_reply:
            (out / "gemini_unreadable_reply.txt").write_text(gc.last_bad_reply, encoding="utf-8")
        res.info["gemini_requests"] = len(gc.tried) + (0 if not (got or gnar) else 1)
        res.info["gemini_model"] = gc.model if (got or gnar) else "none answered"
        if gc.tried:
            res.notes.append("Gemini models that could not answer (switched): " + "; ".join(gc.tried))
        res.notes += gnotes
        stored = rating_cache.store(ai_to_rate, got, gc.model) if ai_to_rate else {}
        if stored:
            for m in ai_need:
                g = stored.get(m.uni.key)
                if g:
                    _set_ai_rating(m.uni, g)
                    snap = excel_vals["snapshot"].get(m.uni.key)
                    if snap is not None:
                        snap["rating"], snap["sentiment"] = m.uni.rating, m.uni.sentiment
            # ratings appear in the workbook (Universal_Snapshot) and in the deck data: rebuild both
            build_workbook(xlsx, client=client, asof=asof, period_label=per.label, parse=bp,
                           priced=priced, snapshot_rows=matched, exception_rows=exc_rows, deck_group_of=deck_group_of)
            data, ni = deck_data_and_input()
            res.notes.append(f"WARN: Rating/Sentiment for {len(stored)} stock(s) written by Gemini ({gc.model}) and "
                             f"saved to {rating_cache.path} - review before sending")
        if gnar:
            nar = {"_note": "Written automatically by Gemini. Review and edit before sending; delete this file to "
                            "have it written again.", **gnar}
            Path(narrative_path).parent.mkdir(parents=True, exist_ok=True)
            Path(narrative_path).write_text(json.dumps(nar, indent=2, ensure_ascii=False), encoding="utf-8")
            res.notes.append(f"WARN: narrative written by Gemini ({gc.model}) and saved to {narrative_path} - review it "
                             "before sending the deck")
    elif nar is not None and resolve_provider_name(settings["narrative_provider"], env) == "gemini":
        res.notes.append(f"Narrative: used the saved {Path(narrative_path).name} (delete it to have Gemini write a new one)")
    if ai_on:
        _write_back_to_sheet(res, uni, matched, rating_cache, per)
        shown = [m for m in matched if m in ai_need or getattr(m.uni, "rating_source", "") == "AI (Gemini)"]
        if shown:
            _write_ai_ratings_md(res, shown, rating_cache)
    res.narrative["present"] = nar is not None
    if nar is not None:
        bad_nums = check_numbers(nar, ni, settings.get("narrative_allowed_numbers", []))
        st = style_issues(nar)
        res.narrative["unverified_numbers"] = bad_nums
        res.narrative["style"] = st
        if bad_nums:
            res.checks.append(V.Check("Narrative figures exist in narrative_input.json",
                                      "FAIL" if settings["narrative_number_check"] == "fail" else "WARN",
                                      f"{len(bad_nums)} figure(s) not found in the data",
                                      [f"{p}: {t}" for p, t in bad_nums]))
        else:
            res.checks.append(V.Check("Narrative figures exist in narrative_input.json", "PASS", "All figures found"))
        hard = [s for s in st if "number" in s or "CREST" in s]
        res.checks.append(V.Check("Narrative house style (CREST caps, headings, length)",
                                  "FAIL" if hard else ("WARN" if st else "PASS"), f"{len(st)} issue(s)", st))
        if any(c.status == "FAIL" for c in res.checks):
            raise PipelineError(f"Narrative checks failed - edit {narrative_path} (see the checks table for the "
                                "exact fields) and re-run; a saved narrative.json is not sent to Gemini again.")

    # 8. deck
    deck = out / f"{safe_filename(client).replace(' ', '_')}_CREST_Portfolio_{per.compact}_Review.pptx"
    fill = fill_deck(template, deck, data=data, narrative=nar, period=per, client=client, asof=asof,
                     dcfg=settings["deck"], ratings_by_ai=ai_on)
    res.outputs["deck"] = deck
    res.narrative["pending_fields"] = fill["pending"]
    res.narrative["problems"] = fill["problems"]
    res.info["table_slides"] = fill["table_slides"]
    if fill["problems"]:
        res.checks.append(V.Check("Narrative structure", "FAIL", f"{len(fill['problems'])} problem(s)", fill["problems"]))

    # 9. QA
    q = run_qa(deck, settings["deck"]["pending_marker"], client, per.label)
    res.qa = q
    res.checks.append(V.simple("No leftover template placeholders", q["placeholders"], fail=True))
    hard = [s for s in q["style"] if not s.startswith("fonts")]
    res.checks.append(V.Check("Deck house style", "FAIL" if hard else ("WARN" if q["style"] else "PASS"),
                              "; ".join(q["style"]) or "CREST in caps, no heading starts with a number, template fonts only",
                              q["style"]))
    res.checks.append(V.simple("Text fits its box (estimate)", q["overflow"]))
    if soffice and render_png:
        pdf, pngs, offpage = render(deck, out / "render", soffice, settings.get("render_png_dpi", 80))
        res.outputs["pdf"] = pdf
        res.outputs["png_dir"] = out / "render" / "png"
        res.checks.append(V.simple("Rendered text stays on the page", offpage))
    if any(c.status == "FAIL" for c in res.checks):
        raise PipelineError("Deck QA failed - see the checks table.")
    res.status = "PASSED WITH WARNINGS" if any(c.status == "WARN" for c in res.checks) or fill["pending"] else "PASSED"


def _ai_items(ms):
    return [{"symbol": m.uni.key, "name": m.uni.values.get("name"), "sector": m.uni.values.get("sector"),
             "result": m.uni.values.get("result"), "commentary": m.uni.values.get("commentary")} for m in ms]


def _set_ai_rating(u, g):
    """Fill only what the Universal Sheet left blank; a typed value always wins."""
    typed = bool(u.rating)
    if not u.rating:
        u.rating = g["rating"]
    if not u.sentiment:
        u.sentiment = g["sentiment"]
    u.rating_source = "Universal Sheet" if typed else "AI (Gemini)"


def _use_sheet_ai_ratings(res, matched):
    """Ratings the tool wrote into the Universal Sheet earlier ('AI ...' in Rating Source) are used as-is while
    the stock's result line and commentary are unchanged; if those changed, the stock is rated again."""
    used = stale = 0
    for m in matched:
        u = m.uni
        b = tag_basis(u.values.get("rating_source"))
        if b is None or u.is_etf:
            continue
        if b == rating_basis(_ai_items([m])[0]) and u.rating and u.sentiment:
            u.rating_source, u.sheet_ai = "AI (Gemini)", True
            used += 1
        else:
            u.rating, u.sentiment = "", ""
            stale += 1
    if used:
        res.notes.append(f"Rating/Sentiment for {used} stock(s) from earlier Gemini answers saved in the Universal "
                         "Sheet (no request)")
    if stale:
        res.notes.append(f"{stale} stock(s) had an AI rating in the Universal Sheet but their result line or "
                         "commentary changed since - rated again")


def _write_back_to_sheet(res, uni, matched, cache, per):
    """Save Gemini's ratings into the Universal Sheet (Rating, Sentiment, Rating Source, Rating Reason) so the next
    run reads them from the sheet. Only stocks whose Rating AND Sentiment were blank in the sheet are written."""
    todo = [m for m in matched if getattr(m.uni, "rating_source", "") == "AI (Gemini)"
            and not getattr(m.uni, "sheet_ai", False) and m.uni.key in cache.data
            and (tag_basis(m.uni.values.get("rating_source")) is not None      # earlier AI value, now refreshed
                 or (not str(m.uni.values.get("rating") or "").strip()         # or both cells blank in the sheet
                     and not str(m.uni.values.get("sentiment") or "").strip()))]
    if not todo:
        return
    ci = uni.col_index
    last = max((i for i, h in enumerate(uni.headers) if str(h or "").strip()), default=len(uni.headers) - 1)
    src_i = ci.get("rating_source", last + 1)
    rsn_i = ci.get("rating_reason", max(last, src_i) + 1)
    L = lambda i: get_column_letter(i + 1)  # noqa: E731
    updates = {}
    if "rating_source" not in ci or "rating_reason" not in ci:
        updates[uni.header_row] = {L(src_i): "Rating Source", L(rsn_i): "Rating Reason"}
    for m in todo:
        e = cache.data[m.uni.key]
        updates[m.uni.row] = {L(ci["rating"]): m.uni.rating, L(ci["sentiment"]): m.uni.sentiment,
                              L(src_i): ai_tag(e.get("model", "gemini"), per.compact, e.get("basis", "")),
                              L(rsn_i): e.get("reason", "")}
    try:
        n = write_cells(uni.path, uni.sheet, updates)
        res.notes.append(f"Saved Gemini's Rating/Sentiment for {len(todo)} stock(s) into the Universal Sheet "
                         f"({uni.path.name}: Rating, Sentiment, Rating Source, Rating Reason; {n} row(s); backup in "
                         f"{uni.path.parent.name}/backups) - next runs read them from the sheet")
    except PermissionError:
        res.notes.append(f"WARN: the Universal Sheet is open in another program, so Gemini's ratings were not written "
                         f"into it; they are kept in {cache.path.name} (no new request) and will be written on a run "
                         "when the sheet is closed")
    except Exception as e:  # noqa: BLE001 - never lose a run over the write-back
        res.notes.append(f"WARN: could not write Gemini's ratings into the Universal Sheet ({e}); they are kept in "
                         f"{cache.path.name}")


def _apply_saved_ai_ratings(res, matched):
    """Apply ratings saved in runs/<period>/ai_ratings.json. -> (cache, stocks needing a rating, items to request)."""
    need = [m for m in matched if not m.uni.is_etf and (not m.uni.rating or not m.uni.sentiment)]
    cache = RatingCache(res.run_dir.parent / "ai_ratings.json")
    cached, todo, _nodata = cache.split(_ai_items(need))
    for m in need:
        if m.uni.key in cached:
            _set_ai_rating(m.uni, cached[m.uni.key])
    if cached:
        res.notes.append(f"Rating/Sentiment for {len(cached)} stock(s) from the saved {cache.path.name} (no request)")
    return cache, need, todo


def _write_ai_ratings_md(res, need, cache):
    rows = []
    for m in need:
        u, g = m.uni, dict(cache.data.get(m.uni.key) or {})
        if getattr(u, "sheet_ai", False):
            g.setdefault("reason", u.values.get("rating_reason") or "")
        has_data = u.values.get("result") or u.values.get("commentary")
        why = g.get("reason") if has_data and g else (
            "not rated: Gemini gave no valid answer" if has_data else
            "no result line or commentary in the Universal Sheet - not rated")
        rows.append((m.holding.raw_symbol, u.rating or "n/a", u.sentiment or "n/a", why))
    out = res.run_dir / "outputs" / "ai_ratings.md"
    out.parent.mkdir(parents=True, exist_ok=True)
    L = ["# Rating / Sentiment by Gemini", "",
         "Only stocks with a blank Rating or Sentiment in the Universal Sheet are listed; a value typed in the sheet "
         "always wins. Shared for this period in `../ai_ratings.json` (edit it to correct a rating; delete an entry "
         "to have it rated again).", "",
         "| Symbol | Rating | Sentiment | Reason |", "|---|---|---|---|"]
    L += [f"| {s} | {r} | {se} | {str(why).replace('|', '/')} |" for s, r, se, why in rows]
    out.write_text("\n".join(L) + "\n", encoding="utf-8")
    res.outputs["ai_ratings"] = out


def _write_rating_preview(path, rows, approved):
    L = ["# Rating / Sentiment preview", "",
         f"Rules in `config/rating_rules.json` are **{'APPROVED - applied' if approved else 'NOT approved - not applied (cells show n/a)'}**. "
         "Only stocks with a blank Rating or Sentiment in the Universal Sheet are listed; a value typed in the sheet always wins.", "",
         "| Symbol | Result line | Proposed rating | Proposed sentiment | Rule |", "|---|---|---|---|---|"]
    for sym, line, rt, se, why in rows:
        L.append(f"| {sym} | {str(line).replace('|', '/')} | {rt} | {se} | {why} |")
    from collections import Counter
    c = Counter(r[2] for r in rows)
    L += ["", "Totals: " + ", ".join(f"{k} {v}" for k, v in sorted(c.items()))]
    path.write_text(chr(10).join(L), encoding="utf-8")


def _py_as_excel(py, matched, priced, deck_group_of):
    """Fallback when LibreOffice is unavailable: same shape as dd.read_portfolio, from Python values."""
    snap = {}
    for m in matched:
        u = m.uni
        snap[m.holding.symbol] = {"key": m.holding.symbol, "short_name": u.values.get("short_name"), "deck_group": deck_group_of(u),
                       "result": u.values.get("result"), "rating": u.rating, "sentiment": u.sentiment,
                       "commentary": u.values.get("commentary"), "is_etf": "Yes" if u.is_etf else "No"}
    rows = []
    for m in priced:
        u, p = m.uni, py[m.holding.symbol]
        mc = value_number(u, "mcap")
        rows.append({"key": m.holding.symbol, "name": u.values["name"],
                     "ticker": u.ticker_crest, "sector": u.values["sector"],
                     "mcap": "ETF" if u.is_etf else (mc if mc is not None else ""), "qty": p["qty"], "avg": p["avg"],
                     "invested": p["invested"], "cmp": p["cmp"], "current": p["current"], "pnl": p["pnl"],
                     "pnl_pct": p["pnl_pct"], "alloc": p["alloc"], "views": u.values["views"]})
    inv = sum(r["invested"] for r in rows)
    cur = sum(r["current"] for r in rows)
    return {"rows": rows, "snapshot": snap,
            "summary": {"invested": inv, "present": cur, "pnl": cur - inv, "pnl_pct": (cur - inv) / inv if inv else 0}}

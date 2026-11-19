"""End-to-end run: holdings + Universal Sheet + template -> Excel, narrative_input.json, deck, report.md."""
from __future__ import annotations

import json
import shutil
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from . import deck_data as dd
from . import validate as V
from .brokers import parse_holdings
from .deck_fill import fill_deck
from .excel_build import build_workbook
from .match import MatchException, match_holdings
from .narrative import SCHEMA, blank_template, check_numbers, get_provider, style_issues
from .office import find_soffice, recalculate_copy
from .qa import render, run_qa
from .ratings import RatingRules
from .universal import load_universe, value_number
from .util import CONFIG_DIR, ROOT, PipelineError, load_json, parse_asof, parse_period, safe_filename


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


def run(*, client, period, asof, universal, holdings, template, narrative=None, broker=None,
        price_source=None, runs_dir=None, settings_path=None, render_png=True, publish_dir=None) -> RunResult:
    settings = load_json(Path(settings_path) if settings_path else CONFIG_DIR / "settings.json")
    if price_source:
        settings["price_source"] = price_source
    ps = settings["price_source"]
    if ps not in ("universal_cmp", "broker_ltp"):
        raise PipelineError(f"price_source must be universal_cmp or broker_ltp, got '{ps}'")
    per = parse_period(period)
    universal, holdings, template = Path(universal), Path(holdings), Path(template)
    for p in (universal, holdings, template):
        if not p.exists():
            raise PipelineError(f"Input file not found: {p}")
    asof_note = None
    if not asof:  # take the as-on date from the holdings file's title rows, else the file's saved date
        asof = parse_holdings(holdings, broker).asof_in_file
        if not asof:
            asof = datetime.fromtimestamp(holdings.stat().st_mtime).strftime("%d/%m/%Y")
            asof_note = (f"WARN: the holdings file has no 'as on' date in its title; used the file's saved date "
                         f"{asof}. Pass --asof dd/mm/yyyy if that is wrong.")
    parse_asof(asof)
    res = RunResult()
    if asof_note:
        res.notes.append(asof_note)
    run_dir = _archive_dir(Path(runs_dir) if runs_dir else ROOT / "runs", per.compact, client)
    res.run_dir = run_dir
    out = run_dir / "outputs"
    for p in (universal, holdings, template):
        shutil.copy2(p, run_dir / "inputs" / p.name)
    narrative_path = Path(narrative) if narrative else run_dir / "narrative.json"
    res.info.update(client=client, period=per.label, asof=asof, price_source=ps, universal=universal.name,
                    holdings=holdings.name, template=template.name, narrative_path=str(narrative_path),
                    started=datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
    try:
        _run_inner(res, settings, per, client, asof, universal, holdings, template, narrative_path, broker, out,
                   render_png)
    except PipelineError as e:
        res.status = "FAILED"
        res.notes.append(f"STOPPED: {e}")
    finally:
        from .report import write_report
        write_report(res, run_dir / "report.md")
        if publish_dir:
            _publish(res, Path(publish_dir) / safe_filename(client) / per.compact)
    return res


def _publish(res, dest: Path):
    """Copy the results (Excel, deck, PDF, report, rating preview) to output/<client>/<period>/."""
    if dest.exists():
        shutil.rmtree(dest)
    dest.mkdir(parents=True)
    for key in ("excel", "deck", "pdf", "rating_preview"):
        p = res.outputs.get(key)
        if p and Path(p).exists():
            shutil.copy2(p, dest / Path(p).name)
    shutil.copy2(res.run_dir / "report.md", dest / "report.md")
    res.outputs["published_to"] = dest


def _run_inner(res, settings, per, client, asof, universal, holdings, template, narrative_path, broker, out, render_png):
    ps = settings["price_source"]
    groups_cfg = load_json(CONFIG_DIR / "sector_groups.json")
    overrides = load_json(CONFIG_DIR / "name_overrides.json").get("overrides", {})

    # 1. broker adapter
    bp = parse_holdings(holdings, broker)
    res.info.update(broker=bp.broker, header_row=bp.header_row, column_map=bp.column_map,
                    holdings_parsed=len(bp.holdings), asof_in_file=bp.asof_in_file)
    if Path(narrative_path).exists():
        shutil.copy2(narrative_path, res.run_dir / "inputs" / "narrative.json")
    if bp.asof_in_file and bp.asof_in_file != asof:
        res.notes.append(f"WARN: holdings file says 'as on {bp.asof_in_file}' but --asof is {asof}")
    for d in bp.dropped:
        res.exceptions.append(Exc("INFO", "Dropped row", str(d["values"][0] if d["values"] else ""),
                                  f"Holdings row {d['row']}: {d['reason']}"))
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

    # Rating / Sentiment: Universal Sheet value > approved rule > blank
    rr = settings.get('rating_rules_file')
    rules = RatingRules(load_json(Path(rr) if Path(rr).is_absolute() else ROOT / rr) if rr else None)
    preview = []
    for m in matched:
        u = m.uni
        u.rating_source = "Universal Sheet" if u.rating else ""
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
        price = h.ltp if ps == "broker_ltp" else value_number(u, "cmp")
        if price is None or price <= 0:
            res.exceptions.append(Exc("WARN", "Missing price", h.raw_symbol,
                                      f"No {'broker LTP' if ps == 'broker_ltp' else 'Universal CMP'} "
                                      f"(cell value: {h.ltp if ps == 'broker_ltp' else u.values.get('cmp')!r}); "
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
    layout = build_workbook(xlsx, client=client, asof=asof, period_label=per.label, price_source=ps, parse=bp,
                            priced=priced, snapshot_rows=matched, exception_rows=exc_rows, deck_group_of=deck_group_of)
    res.outputs["excel"] = xlsx

    soffice = find_soffice(settings.get("libreoffice_path", ""))
    res.info["libreoffice"] = soffice or "not found"
    py = dd.python_values(priced, ps)
    if soffice:
        # The delivered workbook carries Google-Sheets GOOGLEFINANCE formulas (Market Cap, CMP) that LibreOffice
        # cannot evaluate (and would rewrite). So a check copy with this run's values in those two columns is
        # recalculated instead; every other formula is identical. The delivered file is left untouched.
        check = out / "_recalc" / f"check_{xlsx.name}"
        check.parent.mkdir(parents=True, exist_ok=True)
        build_workbook(check, client=client, asof=asof, period_label=per.label, price_source=ps, parse=bp,
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
    # every broker holding exactly once (a matched stock with no price is also in Exceptions)
    all_h = list({h.symbol: h for h in [m.holding for m in matched] +
                  [e.holding for e in res.exceptions if e.holding is not None]}.values())
    res.checks.insert(0, V.tieout(all_h, settings["tieout_extra_tolerance_inr"]))
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
    data = dd.build_deck_data(excel_vals, client=client, period=per, asof=asof, price_source=ps,
                              group_order=groups_cfg["group_order"], unmapped_group=groups_cfg["unmapped_group"],
                              tail_pct=settings["deck"]["tail_weight_threshold_pct"], top_n=settings["deck"]["top_n"])
    ni = {k: v for k, v in data.items() if not k.startswith("_")}
    ni["narrative_schema"] = {k: {"description": d, "max_chars": n} for k, (d, n) in SCHEMA.items()}
    ni["rules"] = ["Use only figures that appear in this file; every number in narrative.json is checked against it.",
                   "Write the brand as CREST. Do not start a heading with a number.",
                   "Keep each field within max_chars so it fits its box."]
    (out / "narrative_input.json").write_text(json.dumps(ni, indent=2, ensure_ascii=False), encoding="utf-8")
    (out / "narrative_template.json").write_text(json.dumps(blank_template(), indent=2), encoding="utf-8")
    res.outputs["narrative_input"] = out / "narrative_input.json"

    # 7. narrative
    prov = get_provider(settings["narrative_provider"], narrative_path)
    nar = prov.get(ni)
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
            raise PipelineError("Narrative checks failed - fix narrative.json (see checks table) and re-run.")

    # 8. deck
    deck = out / f"{safe_filename(client).replace(' ', '_')}_CREST_Portfolio_{per.compact}_Review.pptx"
    fill = fill_deck(template, deck, data=data, narrative=nar, period=per, client=client, asof=asof,
                     dcfg=settings["deck"])
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

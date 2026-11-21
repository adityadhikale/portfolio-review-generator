"""Shared setup for Step A (Portfolio Excel) and Step B (deck): Universal Sheet, matching, ratings, deck groups,
per-stock data checks and the exceptions list. Also the validation gate both steps use."""
from __future__ import annotations

import shutil
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl.utils import get_column_letter

from . import deck_data as dd
from . import validate as V
from . import ai_research
from .ai_ratings import RatingCache, ai_tag, rating_basis, tag_basis
from .match import match_holdings
from .narrative_api import resolve_ai
from .ratings import RatingRules
from .sheet_source import ReadOnlySheet
from .universal import load_universe, value_number
from .util import CONFIG_DIR, ROOT, PipelineError, load_env, load_json, normalize_name


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


@dataclass
class Ctx:
    """Everything one run works on. prepare() fills the second half; Step A and Step B read it."""
    res: RunResult
    step: str                    # portfolio | deck | both
    settings: dict
    per: object
    client: str
    asof: str
    universal: object            # UniversalSheet (crest/sheet_source.py)
    template: Path | None
    narrative_path: Path
    hf: object                   # HoldingsFile (from a template or a Portfolio Excel)
    out: Path
    render_png: bool = True
    groups_cfg: dict = field(default_factory=dict)
    deck_group_of: object = None
    uni: object = None
    matched: list = field(default_factory=list)
    priced: list = field(default_factory=list)
    env: dict = field(default_factory=dict)
    ai_on: bool = False
    rating_cache: object = None
    ai_need: list = field(default_factory=list)
    ai_to_rate: list = field(default_factory=list)
    soffice: str | None = None
    xlsx: Path | None = None     # set by Step A
    layout: object = None
    exc_rows: list = field(default_factory=list)
    progress: object = None      # progress(stage key) -> None; the website shows the stage live

    def stage(self, key: str):
        if self.progress:
            self.progress(key)


def prepare(ctx: Ctx):
    res, settings, per, bp, out = ctx.res, ctx.settings, ctx.per, ctx.hf, ctx.out
    narrative_path, universal = ctx.narrative_path, ctx.universal
    groups_cfg = load_json(CONFIG_DIR / "sector_groups.json")
    overrides = load_json(CONFIG_DIR / "name_overrides.json").get("overrides", {})

    # 1. holdings (already read and checked by run())
    res.info.update(holdings_sheet=bp.sheet, holdings_parsed=len(bp.holdings))
    if Path(narrative_path).exists():
        shutil.copy2(narrative_path, res.run_dir / "inputs" / "narrative.json")
    res.notes += bp.notes

    # 2. universal + match
    ctx.stage("match")
    uni = load_universe(universal)
    res.info.update(universal_rows=len(uni.rows), universal_header_row=uni.header_row, universal_sheet=uni.sheet)
    for w in uni.warnings:
        res.exceptions.append(Exc("WARN", "Universal duplicate", "", w))
    matched, mexc = match_holdings(bp.holdings, uni)
    for e in mexc:
        res.exceptions.append(Exc("WARN", e.kind, e.symbol, e.detail, e.holding))
    res.info["matched_by"] = {k: sum(1 for m in matched if m.method == k) for k in ("isin", "symbol", "ticker code", "name")}
    _note_dropped_since_step_a(res, bp, matched)

    ctx.stage("research")
    # 2b. AI research: blank / out-of-date Sector, Result Line and Commentary for the held stocks (before ratings,
    # which read the result line). One request is kept back for the ratings + narrative request.
    env = load_env()
    _research(res, settings, per, env, uni, matched)

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
    rp = settings.get("rating_provider", "auto")
    ai_on = rp == "gemini" or (rp == "auto" and resolve_ai(env))
    rating_cache, ai_need, ai_to_rate = None, [], []
    if ai_on:
        _use_sheet_ai_ratings(res, matched)
        rating_cache, ai_need, ai_to_rate = _apply_saved_ai_ratings(res, matched)
    # those are rated by the deck step's request; a Portfolio-only run flags them as missing
    pending_ai = {t["symbol"] for t in ai_to_rate} if ctx.step != "portfolio" else set()
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
            if u.key not in pending_ai:
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

    # data checks shared by both steps (the FAIL stop is in gate(), after the Excel is written for inspection)
    all_h = bp.holdings
    accounted = {id(m.holding) for m in matched} | {id(e.holding) for e in res.exceptions if e.holding is not None}
    lost = [h.raw_symbol for h in all_h if id(h) not in accounted]
    res.checks.insert(0, V.Check("Every holding is in the Portfolio or in Exceptions",
                                 "FAIL" if lost else "PASS",
                                 f"Not accounted for: {', '.join(lost)}" if lost else
                                 f"All {len(all_h)} holding(s) from {sum(len(h.source_rows) for h in all_h)} "
                                 "row(s) accounted for"))
    unm = [f"{e.symbol}: {e.detail}" for e in res.exceptions if e.kind in ("Unmatched", "Ambiguous")]
    res.checks.append(V.simple("Unmatched / ambiguous symbols (listed in Exceptions)", unm,
                               fail=settings["unmatched_policy"] == "fail"))
    res.checks.append(V.simple("Missing price (listed in Exceptions)",
                               [e.symbol for e in res.exceptions if e.kind == "Missing price"]))
    res.checks.append(sector_chk)
    res.checks.append(V.simple("Sector filled in the Universal Sheet",
                               [f"{e.symbol}: {e.detail.split(':')[0]}" for e in res.exceptions if e.kind == "Missing sector"]))
    tick = [f"{e.symbol}: {e.detail}" for e in res.exceptions if e.kind == "Ticker/data note"]
    res.checks.append(V.simple("Ticker format NSE:/BSE: (BOM: converted, unknown exchange flagged)", tick, warn_if_any=any(
        "defaulted" in t or "not one of" in t for t in tick)))

    sev_order = {"FAIL": 0, "WARN": 1, "INFO": 2}
    ctx.exc_rows = sorted(res.exceptions, key=lambda e: sev_order.get(e.severity, 3))
    ctx.groups_cfg, ctx.deck_group_of, ctx.uni, ctx.matched, ctx.priced = groups_cfg, deck_group_of, uni, matched, priced
    ctx.env, ctx.ai_on, ctx.rating_cache, ctx.ai_need, ctx.ai_to_rate = env, ai_on, rating_cache, ai_need, ai_to_rate


def _research(res, settings, per, env, uni, matched):
    budget = max(0, int(settings.get("ai_max_requests_per_run", 3)) - 1)
    out = ai_research.research(matched=matched, per=per, settings=settings, env=env,
                               cache_path=res.run_dir.parent / "ai_research.json", budget=budget,
                               say=lambda t: print(t, flush=True))
    res.notes += out.notes
    res.info["ai_research"] = {"requests": out.requests, "searches": out.searches, "filled": len(out.filled),
                               "skipped": out.skipped}
    if out.filled:
        ai_research.write_back(res.notes, uni, matched, per)
    res.info["ai_sourced"] = ai_research.ai_sourced(matched)
    if res.info["ai_sourced"]:
        res.notes.append(f"WARN: {len(res.info['ai_sourced'])} Sector / Result Line / Commentary value(s) in this "
                         "portfolio were written by the AI - review them before sending (report: 'AI-sourced')")


def _note_dropped_since_step_a(res, bp, matched):
    """Deck from a Portfolio Excel: warn about holdings Step A had that the Portfolio sheet no longer has (deleted by
    hand, or left out at Step A for a missing price). A row whose Key was cleared but which matched again by Ticker or
    Name still counts as present."""
    if not getattr(bp, "step_a_holdings", None):
        return
    have = {h.symbol for h in bp.holdings} | {m.uni.key for m in matched}
    names = {normalize_name(m.uni.values.get("name") or "") for m in matched} | \
            {normalize_name(h.name) for h in bp.holdings if h.name}
    gone = [k for k, n in bp.step_a_holdings.items() if k not in have and (not n or normalize_name(n) not in names)]
    if gone:
        res.notes.append(f"WARN: {len(gone)} holding(s) from Step A (sheet Holdings_Normalized) are not in the "
                         f"Portfolio sheet, so they are not in this deck: {', '.join(gone)}. Add them to the Portfolio "
                         "sheet if they should be.")


def python_excel_values(ctx: Ctx) -> dict:
    """The Portfolio numbers computed in Python, in the shape dd.read_portfolio returns (Step B without Step A)."""
    return _py_as_excel(dd.python_values(ctx.priced), ctx.matched, ctx.priced, ctx.deck_group_of)


def gate(ctx: Ctx, excel_vals: dict, stop_message: str):
    """Totals, allocation and ETF checks on the Portfolio numbers; any FAIL so far stops the run."""
    res, settings = ctx.res, ctx.settings
    all_h = ctx.hf.holdings
    in_table = {m.holding.symbol for m in ctx.priced}
    exc_inv = sum(h.quantity * h.avg_price for h in all_h if h.symbol not in in_table)
    res.checks.append(V.excel_ties(excel_vals["summary"]["invested"] or 0, exc_inv, all_h))
    res.checks.append(V.allocation(sum(float(r["alloc"] or 0) for r in excel_vals["rows"]),
                                   settings["allocation_tolerance_pct"]))
    etf_rows = [r for r in excel_vals["rows"] if excel_vals["snapshot"].get(r["key"], {}).get("is_etf") == "Yes"]
    bad_etf = [r["key"] for r in etf_rows if r["mcap"] != "ETF"]
    res.checks.append(V.Check("ETF market cap shown as 'ETF' (not #N/A)", "FAIL" if bad_etf else "PASS",
                              f"ETFs: {[r['key'] for r in etf_rows] or 'none'}", bad_etf))
    if any(c.status == "FAIL" for c in res.checks):
        raise PipelineError(stop_message)


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
    updates, expect = {}, {}
    if "rating_source" not in ci or "rating_reason" not in ci:
        updates[uni.header_row] = {L(src_i): "Rating Source", L(rsn_i): "Rating Reason"}
    for m in todo:
        expect[m.uni.row] = (L(ci["name"]), m.uni.values.get("name"))   # skip the row if the sheet moved since
        e = cache.data[m.uni.key]
        updates[m.uni.row] = {L(ci["rating"]): m.uni.rating, L(ci["sentiment"]): m.uni.sentiment,
                              L(src_i): ai_tag(e.get("model", "gemini"), per.compact, e.get("basis", "")),
                              L(rsn_i): e.get("reason", "")}
    if uni.origin is None:
        return
    try:
        n = len([r for r in uni.origin.write(updates, expect) if r != uni.header_row])
        res.notes += [n_ for n_ in uni.origin.notes() if n_ not in res.notes]
        res.notes.append(f"Saved Gemini's Rating/Sentiment for {n} stock(s) into the Universal Sheet ({uni.source}: "
                         "Rating, Sentiment, Rating Source, Rating Reason) - next runs read them from the sheet")
    except ReadOnlySheet as e:
        res.notes.append(f"Gemini's ratings for {len(todo)} stock(s) were not written into the Universal Sheet ({e}); "
                         f"they are kept in {cache.path.name} (no new request) and are written on the next run that "
                         "reads the Google Sheet")
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


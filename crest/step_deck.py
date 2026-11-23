"""Step B: deck data, narrative (and the narrative-API request), the deck, its QA and the PDF.

Holdings come from a Portfolio Excel or a holdings template; price, market cap, sector, views, results, commentary,
rating and sentiment always come fresh from the Universal Sheet."""
from __future__ import annotations

import json
from pathlib import Path

from . import deck_data as dd
from . import validate as V
from .ai_review import request_review
from .deck_fill import fill_deck
from .excel_build import build_workbook
from .narrative import SCHEMA, ManualProvider, blank_template, check_numbers, resolve_provider_name, style_issues
from .narrative_api import GeminiClient
from .office import find_soffice
from .qa import render, run_qa
from .shared import Ctx, _set_ai_rating, _write_ai_ratings_md, _write_back_to_sheet, gate, python_excel_values
from .util import PipelineError, hosted, safe_filename


def build_deck(ctx: Ctx, excel_vals: dict | None = None):
    """excel_vals: Step A's recalculated numbers in a 'both' run; None for a deck-only run (computed here)."""
    res, settings, per, out = ctx.res, ctx.settings, ctx.per, ctx.out
    client, asof, bp, template = ctx.client, ctx.asof, ctx.hf, ctx.template
    priced, matched, deck_group_of, uni = ctx.priced, ctx.matched, ctx.deck_group_of, ctx.uni
    groups_cfg, env, ai_on = ctx.groups_cfg, ctx.env, ctx.ai_on
    rating_cache, ai_need, ai_to_rate = ctx.rating_cache, ctx.ai_need, ctx.ai_to_rate
    narrative_path, render_png = ctx.narrative_path, ctx.render_png
    soffice = ctx.soffice if ctx.xlsx else find_soffice(settings.get("libreoffice_path", ""))
    if excel_vals is None:
        res.info.setdefault("libreoffice", soffice or "not found")
        excel_vals = python_excel_values(ctx)
        gate(ctx, excel_vals, "Validation gate failed - see the checks table. Deck not built.")

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

    ctx.stage("ai")
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
                    snap = excel_vals["snapshot"].get(m.holding.symbol)
                    if snap is not None:
                        snap["rating"], snap["sentiment"] = m.uni.rating, m.uni.sentiment
            # ratings appear in the workbook (Universal_Snapshot, when Step A ran) and in the deck data
            if ctx.xlsx:
                build_workbook(ctx.xlsx, client=client, asof=asof, period_label=per.label, parse=bp, priced=priced,
                               snapshot_rows=matched, exception_rows=ctx.exc_rows, deck_group_of=deck_group_of)
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
    ctx.stage("deck")
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
    ctx.stage("qa")
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
    elif hosted():
        res.notes.append("PDF: open the PowerPoint and use File > Save as PDF (the hosted website makes the PPT only).")
    if any(c.status == "FAIL" for c in res.checks):
        raise PipelineError("Deck QA failed - see the checks table.")
    res.status = "PASSED WITH WARNINGS" if any(c.status == "WARN" for c in res.checks) or fill["pending"] else "PASSED"

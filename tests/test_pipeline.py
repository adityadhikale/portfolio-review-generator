"""Tests.  Run:  python -m unittest discover -s tests -v

Unit tests always run. End-to-end tests need the private sample files, which are NOT in git (client data):
samples/holdings.xlsx (a broker export, copied into the holdings template at test start), samples/reference_portfolio.xlsx, samples/reference_deck_Q1FY27.pptx,
templates/Deck_-_template.pptx and the fixtures built from them by tools/build_fixture.py and
tools/extract_narrative.py. Without them those tests are skipped."""
from __future__ import annotations

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import openpyxl
from pptx import Presentation

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from crest.holdings import read_template  # noqa: E402
from crest.deck_fill import chunk  # noqa: E402
from crest.narrative import check_numbers, style_issues  # noqa: E402
from crest.pipeline import run  # noqa: E402
from crest.universal import load_universe  # noqa: E402
from crest.util import PipelineError, normalize_symbol, parse_period, parse_ticker  # noqa: E402
from tools.build_holdings_template import build as build_blank_template  # noqa: E402

S = ROOT / "samples"
FX = ROOT / "tests" / "fixtures"
BROKER, TPL, UNI = S / "holdings.xlsx", ROOT / "templates" / "Deck_-_template.pptx", FX / "universal_sample.xlsx"
HAVE_SAMPLES = all(p.exists() for p in (BROKER, TPL, UNI, FX / "narrative_reference.json"))
needs_samples = unittest.skipUnless(HAVE_SAMPLES, "private sample files not present (not in git)")
_TMP = Path(tempfile.mkdtemp(prefix="crest-hold-"))
HOLD = _TMP / "Sample Client.xlsx"   # the private broker sample copied into the holdings template (setUpModule)
ARGS = dict(period="Q1FY27", universal=UNI, holdings=HOLD, template=TPL)


def make_template(path: Path, rows, client="Sample Client", asof="08/10/2026") -> Path:
    """A filled holdings template: rows are (ISIN, Symbol, Company Name, Quantity, Average Price)."""
    build_blank_template(path)
    wb = openpyxl.load_workbook(path)
    ws = wb["Holdings"]
    ws["C2"], ws["C3"] = client, asof
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            ws.cell(6 + i, 2 + j, v)
    wb.save(path)
    return path


def sample_to_template():
    """samples/holdings.xlsx (broker export: header on row 3, SYMBOL / QTY / AVG PRICE) -> holdings template."""
    ws = openpyxl.load_workbook(BROKER, data_only=True).active
    rows = [(None, r[0], None, r[1], r[3]) for r in ws.iter_rows(min_row=4, values_only=True)
            if isinstance(r[0], str) and r[0].strip() and r[1]]
    make_template(HOLD, rows)

# Tests never call the Gemini API: ignore any real key in .env (narrative_provider 'auto' -> manual).
_offline = mock.patch("crest.pipeline.load_env", return_value={})


def setUpModule():
    _offline.start()
    if HAVE_SAMPLES:
        sample_to_template()


def tearDownModule():
    _offline.stop()
    shutil.rmtree(_TMP, ignore_errors=True)


class Units(unittest.TestCase):
    def test_normalize_symbol(self):
        self.assertEqual(normalize_symbol(" hblengine-eq "), "HBLENGINE")
        self.assertEqual(normalize_symbol("NSE:TIIL"), "TIIL")
        self.assertEqual(normalize_symbol("RELIANCE.NS"), "RELIANCE")
        self.assertEqual(normalize_symbol(540565.0), "540565")

    def test_ticker(self):
        self.assertEqual(parse_ticker("BOM:540565"), ("BSE", "540565", False))
        self.assertEqual(parse_ticker("NSE:SJS"), ("NSE", "SJS", False))
        self.assertEqual(parse_ticker("SJS"), ("NSE", "SJS", True))      # unknown exchange -> NSE, flagged
        self.assertEqual(parse_ticker("SJS", "BSE"), ("BSE", "SJS", False))

    def test_period(self):
        p = parse_period("Q1FY27")
        self.assertEqual((p.label, p.compact, p.quarter_end), ("Q1 FY27", "Q1FY27", "30 Jun 2026"))
        self.assertEqual(parse_period("Q4FY27").quarter_end, "31 Mar 2027")

    def test_chunk(self):
        self.assertEqual(chunk(8, 8, 6), [8])
        self.assertEqual(chunk(11, 8, 6), [6, 5])
        self.assertEqual(chunk(10, 8, 6), [5, 5])
        self.assertEqual(chunk(13, 8, 6), [5, 4, 4])


class HoldingsTemplate(unittest.TestCase):
    def setUp(self):
        self.td = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.td, True)

    def _err(self, f) -> str:
        with self.assertRaises(PipelineError) as cm:
            read_template(f)
        return str(cm.exception)

    def test_blank_template_layout(self):
        f = build_blank_template(self.td / "blank.xlsx")
        wb = openpyxl.load_workbook(f)
        self.assertEqual(wb.sheetnames, ["Holdings", "How to fill"])
        ws = wb["Holdings"]
        self.assertEqual((ws["B2"].value, ws["B3"].value), ("Client Name", "As On Date"))
        self.assertEqual([ws.cell(5, c).value for c in range(2, 7)],
                         ["ISIN", "Symbol", "Company Name", "Quantity", "Average Price"])
        self.assertIn("C2", self._err(f))          # blank client name and date stop the run
        self.assertIn("C3", self._err(f))

    def test_good_file(self):
        from datetime import datetime
        f = make_template(self.td / "h.xlsx", [
            ("INE775A01035", "MOTHERSON", "Samvardhana Motherson", 6000, 95.95),
            (None, "540565", None, "1,200", "101.5"),       # BSE code, numbers typed as text
            ("INE048G01026", None, "Navin Fluorine", 22, 4543.51),
        ], client="  Sharma Family ", asof=datetime(2026, 10, 7))
        h = read_template(f)
        self.assertEqual((h.client, h.asof, h.sheet), ("Sharma Family", "07/10/2026", "Holdings"))
        self.assertEqual([x.symbol for x in h.holdings], ["MOTHERSON", "540565", "INE048G01026"])
        self.assertEqual(h.holdings[1].quantity, 1200)
        self.assertEqual(h.holdings[0].isin, "INE775A01035")
        self.assertEqual(h.holdings[2].name, "Navin Fluorine")
        self.assertEqual(h.raw_rows[1][2], "  Sharma Family ")    # Raw_Holdings keeps the sheet unchanged

    def test_merged_duplicates(self):
        f = make_template(self.td / "h.xlsx", [
            (None, "ABC-EQ", None, 10, 100), ("INE000A01011", "abc", None, 30, 200),
            ("INE000A01011", None, None, 10, 300), (None, "XYZ", None, 5, 50)])
        h = read_template(f)
        self.assertEqual([x.symbol for x in h.holdings], ["ABC", "XYZ"])
        abc = h.holdings[0]
        self.assertAlmostEqual(abc.quantity, 50)
        self.assertAlmostEqual(abc.avg_price, (10 * 100 + 30 * 200 + 10 * 300) / 50)
        self.assertEqual((abc.source_rows, abc.isin), ([6, 7, 8], "INE000A01011"))
        self.assertTrue(any("merged" in n for n in h.notes))

    def test_bad_numbers_and_missing_ids(self):
        f = make_template(self.td / "h.xlsx", [
            (None, "ABC", None, "ten", 100), (None, "DEF", None, 5, 0), (None, None, "Some Co", 5, 10),
            ("INE12", "GHI", None, 1, 1), (None, "JKL", None, -3, None)])
        e = self._err(f)
        for cell in ("E6", "F7", "Row 8", "B9", "E10", "F10"):
            self.assertIn(cell, e)

    def test_missing_column(self):
        f = make_template(self.td / "h.xlsx", [(None, "ABC", None, 1, 1)])
        wb = openpyxl.load_workbook(f)
        wb["Holdings"]["F5"] = "Avg Cost"
        wb.save(f)
        self.assertIn("'Average Price' is missing", self._err(f))

    def test_bad_date_and_not_a_template(self):
        self.assertIn("C3", self._err(make_template(self.td / "a.xlsx", [(None, "A", None, 1, 1)], asof="2026-10-07")))
        f = self.td / "broker.xlsx"
        wb = openpyxl.Workbook()
        wb.active.append(["SYMBOL", "QTY", "AVG PRICE"])
        wb.save(f)
        self.assertIn("not a CREST holdings template", self._err(f))
        f.with_suffix(".csv").write_text("SYMBOL,QTY\n", encoding="utf-8")
        self.assertIn(".xlsx", self._err(f.with_suffix(".csv")))

    def test_rows_after_blank_row_warned(self):
        f = make_template(self.td / "h.xlsx", [(None, "ABC", None, 1, 1), (None, None, None, None, None),
                                                (None, "DEF", None, 2, 2)])
        h = read_template(f)
        self.assertEqual([x.symbol for x in h.holdings], ["ABC"])
        self.assertTrue(any(n.startswith("WARN") and "8" in n for n in h.notes))

    @needs_samples
    def test_sample_converted(self):
        h = read_template(HOLD)
        self.assertEqual(len(h.holdings), 54)
        self.assertEqual((h.client, h.asof), ("Sample Client", "08/10/2026"))


class GeminiSingleRequest(unittest.TestCase):
    """The one combined Gemini request (ratings + narrative), with the HTTP call replaced: no network."""
    NI = {"run": {"client": "Secret Client", "period": "Q2FY27"},
          "portfolio": {"unrealized_pnl_pct": 37.28, "holdings_count": 39},
          "top_holdings": [{"symbol": "ABC", "weight_pct": 21.0}]}
    ITEMS = [{"symbol": "AAA", "name": "Aaa Ltd", "sector": "Pharma", "result": "Rev +30% YoY; PAT +50% YoY", "commentary": "Strong"},
             {"symbol": "BBB", "name": "Bbb Ltd", "sector": "Bank", "result": "PAT -40% YoY", "commentary": ""},
             {"symbol": "CCC", "name": "Ccc Ltd", "sector": "IT", "result": "", "commentary": ""}]

    @staticmethod
    def _client(replies):
        from crest.narrative_api import GeminiClient
        sent = []

        def post(url, body):
            sent.append((url, body))
            return {"candidates": [{"content": {"parts": [{"text": json.dumps(replies.pop(0))}]}}]}
        return GeminiClient(api_key="k", post=post), sent

    def test_one_request_for_ratings_and_narrative(self):
        from crest.ai_review import request_review
        nar = {"slide2": {"headline": "Steady compounding"}}
        reply = {"ratings": [{"symbol": "AAA", "rating": "excellent", "sentiment": "Bullish", "reason": "Strong"},
                             {"symbol": "BBB", "rating": "Terrible", "sentiment": "Bearish"}],
                 "narrative": nar}
        gc, sent = self._client([reply])
        ratings, narrative, notes = request_review(gc, narrative_input=self.NI, to_rate=self.ITEMS[:2])
        self.assertEqual(len(sent), 1)                                   # exactly one request
        self.assertEqual(ratings, {"AAA": {"rating": "Excellent", "sentiment": "Bullish", "reason": "Strong"}})
        self.assertEqual(narrative, nar)
        self.assertTrue(any("BBB" in n for n in notes))                  # invalid -> n/a, no retry request
        body = json.dumps(sent[0][1])
        self.assertNotIn("Secret Client", body)
        self.assertIn("gemini-pro-latest", sent[0][0])   # first model of the built-in list

    def test_switches_model_when_busy(self):
        from crest.narrative_api import GeminiClient, ModelUnavailable
        urls, waits = [], []

        def post(url, body):
            urls.append(url)
            if "busy-1" in url or "busy-2" in url:
                raise ModelUnavailable("503 busy")
            return {"candidates": [{"content": {"parts": [{"text": "{}"}]}}]}
        gc = GeminiClient(api_key="k", model="busy-1, busy-2, ok-3", post=post, sleep=waits.append)
        self.assertEqual(gc.generate("s", []), "{}")
        self.assertEqual([u.split("/models/")[1].split(":")[0] for u in urls], ["busy-1", "busy-2", "ok-3"])
        self.assertEqual(gc.model, "ok-3")
        self.assertEqual(waits, [])                                # switching does not wait
        self.assertEqual(len(gc.tried), 2)

    def test_all_models_busy_two_passes_then_stop(self):
        from crest.narrative_api import GeminiClient, ModelUnavailable
        urls, waits = [], []

        def post(url, body):
            urls.append(url)
            raise ModelUnavailable("503 busy")
        gc = GeminiClient(api_key="k", model="a,b", post=post, sleep=waits.append)
        with self.assertRaises(PipelineError):
            gc.generate("s", [])
        self.assertEqual(len(urls), 4)                             # each model twice
        self.assertEqual(waits, [30])

    def test_reply_parsing_and_repair(self):
        from crest.narrative_api import parse_json_object as p
        self.assertEqual(p('```json\n{"a": 1}\n```'), {"a": 1})
        self.assertEqual(p('Here you go: {"a": 1} hope it helps'), {"a": 1})
        self.assertEqual(p('{"a": {"b": [1, "x}"'), {"a": {"b": [1, "x}"]}})   # missing closers added
        for bad in ("no json at all", "[1, 2]", '{"a": 1,, "b": 2}'):
            with self.assertRaises(ValueError):
                p(bad)

    def test_unusable_reply_switches_model(self):
        from crest.narrative_api import GeminiClient
        urls = []

        def post(url, body):
            urls.append(url)
            text = "not json" if "bad" in url else '{"ok": true}'
            return {"candidates": [{"content": {"parts": [{"text": text}]}}]}
        gc = GeminiClient(api_key="k", model="bad,good", post=post)
        self.assertEqual(gc.generate_json("s", []), {"ok": True})
        self.assertEqual(gc.model, "good")
        self.assertEqual(gc.last_bad_reply, "not json")

    def test_fatal_error_does_not_switch(self):
        from crest.narrative_api import GeminiClient
        urls = []

        def post(url, body):
            urls.append(url)
            raise PipelineError("Gemini API error 400")
        gc = GeminiClient(api_key="k", model="a,b", post=post)
        with self.assertRaises(PipelineError):
            gc.generate("s", [])
        self.assertEqual(len(urls), 1)

    def test_ratings_only_and_nothing_to_do(self):
        from crest.ai_review import request_review
        gc, sent = self._client([{"ratings": [{"symbol": "AAA", "rating": "Good", "sentiment": "Neutral"}]}])
        ratings, narrative, _ = request_review(gc, narrative_input=None, to_rate=self.ITEMS[:1])
        self.assertEqual(ratings["AAA"]["rating"], "Good")
        self.assertIsNone(narrative)
        self.assertNotIn("narrative_input", sent[0][1]["contents"][0]["parts"][0]["text"])
        self.assertEqual(request_review(gc, narrative_input=None, to_rate=[]), ({}, None, []))
        self.assertEqual(len(sent), 1)                                   # second call made no request

    def test_rating_cache(self):
        from crest.ai_ratings import RatingCache
        with tempfile.TemporaryDirectory() as td:
            c = RatingCache(Path(td) / "ai_ratings.json")
            cached, todo, nodata = c.split(self.ITEMS)
            self.assertEqual((cached, [t["symbol"] for t in todo], [t["symbol"] for t in nodata]), ({}, ["AAA", "BBB"], ["CCC"]))
            c.store(todo, {"AAA": {"rating": "Excellent", "sentiment": "Bullish", "reason": "x"}}, "m")
            c2 = RatingCache(Path(td) / "ai_ratings.json")
            cached, todo, _ = c2.split(self.ITEMS)
            self.assertEqual(list(cached), ["AAA"])
            self.assertEqual([t["symbol"] for t in todo], ["BBB"])
            changed = [dict(self.ITEMS[0], result="Rev +5% YoY; PAT +2% YoY")]   # new result line -> rated again
            self.assertEqual([t["symbol"] for t in c2.split(changed)[1]], ["AAA"])

    def test_auto_picks_provider_from_env(self):
        from crest.narrative import resolve_provider_name
        self.assertEqual(resolve_provider_name("auto", {}), "manual")
        self.assertEqual(resolve_provider_name("auto", {"NARRATIVE_PROVIDER": "gemini"}), "manual")  # no key
        self.assertEqual(resolve_provider_name("auto", {"NARRATIVE_PROVIDER": "Gemini", "NARRATIVE_API_KEY": "k"}), "gemini")
        self.assertEqual(resolve_provider_name("manual", {"NARRATIVE_PROVIDER": "gemini", "NARRATIVE_API_KEY": "k"}), "manual")

    def test_load_env(self):
        from crest.util import load_env
        with tempfile.TemporaryDirectory() as td:
            f = Path(td) / ".env"
            f.write_text('# c\nNARRATIVE_PROVIDER=gemini\nNARRATIVE_MODEL="m-1"\n\nBROKEN LINE\n', encoding="utf-8")
            with mock.patch.dict("os.environ", {"NARRATIVE_MODEL": ""}, clear=False):
                env = load_env(f)
            self.assertEqual(env["NARRATIVE_PROVIDER"], "gemini")
            self.assertEqual(env["NARRATIVE_MODEL"], "m-1")


class SheetWriteBack(unittest.TestCase):
    def test_writes_only_given_cells(self):
        from crest.sheet_writeback import tag_basis, write_cells
        with tempfile.TemporaryDirectory() as td:
            f = Path(td) / "u.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Universal"
            ws.append(["Name", "Ticker", "Rating", "Sentiment"])
            ws.append(["Aaa Ltd", "NSE:AAA", None, None])
            ws.append(["Bbb Ltd", "NSE:BBB", "Good", "Bullish"])
            ws["E2"] = "=1+1"
            wb.save(f)
            n = write_cells(f, "Universal", {1: {"F": "Rating Source"}, 2: {"C": "Weak", "D": "Bearish", "F": "AI · m · Q2FY27 · ab12"}})
            self.assertEqual(n, 2)
            ws2 = openpyxl.load_workbook(f).active
            self.assertEqual([c.value for c in ws2[2]], ["Aaa Ltd", "NSE:AAA", "Weak", "Bearish", "=1+1", "AI · m · Q2FY27 · ab12"])
            self.assertEqual([c.value for c in ws2[3]][:4], ["Bbb Ltd", "NSE:BBB", "Good", "Bullish"])
            self.assertEqual(ws2["F1"].value, "Rating Source")
            self.assertTrue(list((Path(td) / "backups").glob("u-*.xlsx")))
            self.assertEqual(tag_basis(ws2["F2"].value), "ab12")
            self.assertIsNone(tag_basis("CREST"))
            self.assertIsNone(tag_basis(None))


@needs_samples
class GeminiOutage(unittest.TestCase):
    def test_deck_built_when_gemini_fails(self):
        from crest import narrative_api as gemini
        env = {"NARRATIVE_PROVIDER": "gemini", "NARRATIVE_API_KEY": "fake"}
        calls = []

        def down(self_, url, body):
            calls.append(1)
            raise PipelineError("Gemini API error 503 (Gemini is overloaded - try again later)")
        with tempfile.TemporaryDirectory() as td, mock.patch("crest.pipeline.load_env", return_value=env), \
                mock.patch.object(gemini.GeminiClient, "_http_post", down):
            res = run(**ARGS, runs_dir=td, render_png=False)
            self.assertNotEqual(res.status, "FAILED")
            self.assertTrue(Path(res.outputs["deck"]).exists())
            self.assertEqual(len(calls), 1)
            self.assertTrue(any(c.name.startswith("Gemini request") and c.status == "WARN" for c in res.checks))


class UniversalSchema(unittest.TestCase):
    @needs_samples
    def test_fixture_loads(self):
        u = load_universe(UNI)
        self.assertEqual(len(u.rows), 54)
        self.assertIn("INDIGRID", u.by_key)
        self.assertEqual(u.by_key["INDIGRID"].ticker_crest, "BSE:540565")

    def test_missing_quarterly_columns_stop(self):
        with tempfile.TemporaryDirectory() as td:
            f = Path(td) / "u.xlsx"
            wb = openpyxl.Workbook()
            wb.active.append(["Name", "Ticker", "Sector", "Market Cap (Cr)", "CMP", "Views"])
            wb.active.append(["X Ltd", "NSE:X", "Banking", 1, 1, "Hold"])
            wb.save(f)
            with self.assertRaises(PipelineError) as cm:
                load_universe(f)
            self.assertIn("Rating", str(cm.exception))


@needs_samples
class EndToEnd(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = Path(tempfile.mkdtemp(prefix="crest-test-"))
        cls.res = run(**ARGS, runs_dir=cls.tmp, render_png=False)
        cls.out = cls.res.run_dir / "outputs"

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def test_status_and_checks(self):
        self.assertTrue(self.res.status.startswith("PASSED"), self.res.notes)
        st = {c.name: c.status for c in self.res.checks}
        self.assertEqual(st["Every template holding is in the Portfolio or in Exceptions"], "PASS")
        self.assertEqual(st["Excel invested + Exceptions = template holdings"], "PASS")
        self.assertEqual(st["% Allocation sums to 100%"], "PASS")
        self.assertEqual(st["ETF market cap shown as 'ETF' (not #N/A)"], "PASS")
        self.assertEqual(st["No leftover template placeholders"], "PASS")
        self.assertEqual(st["Sector sanity"], "WARN")
        flagged = {f["symbol"] for c in self.res.checks if c.name == "Sector sanity" for f in c.items}
        self.assertTrue({"CEMPRO", "AXISCADES", "BALRAMCHIN", "BETA"} <= flagged)

    def test_excel_formulas(self):
        wb = openpyxl.load_workbook(self.out / "Sample Client Portfolio.xlsx")
        self.assertEqual(wb.sheetnames, ["Portfolio", "Raw_Holdings", "Holdings_Normalized", "Universal_Snapshot", "Exceptions"])
        ws = wb["Portfolio"]
        self.assertEqual(ws["B8"].value, "Name")
        values = {2, 4, 6, 7}            # Name, Sector, Quantity, Average Price: plain values
        for r in range(9, 9 + 54):
            for c in range(2, 15):
                v = ws.cell(r, c).value
                is_f = isinstance(v, str) and v.startswith("=")
                self.assertEqual(is_f, c not in values, f"{ws.cell(r, c).coordinate}: {v}")
        for cell in ("C3", "C4", "C5", "C6"):
            self.assertTrue(str(ws[cell].value).startswith("="))
        self.assertEqual(ws["B9"].value, "HBL Power Systems Limited")
        self.assertEqual(ws["E9"].value, '=IFERROR(__xludf.DUMMYFUNCTION("GOOGLEFINANCE(C9,""marketcap"")/10^7"),21449.11)')
        self.assertEqual(ws["I9"].value, '=IFERROR(__xludf.DUMMYFUNCTION("GOOGLEFINANCE(C9,""price"")"),773.0)')
        self.assertEqual(ws["H9"].value, "=F9*G9")
        self.assertEqual(ws["L9"].value, "=K9/H9")
        self.assertEqual(ws["M9"].value, f"=J9/SUM($J$9:$J${9 + 53})")
        bse = [ws.cell(r, 5).value for r in range(9, 63) if "SUBSTITUTE" in str(ws.cell(r, 5).value)]
        self.assertEqual(len(bse), 1)  # IndiGrid (BSE:540565) -> BOM: for Google Finance
        self.assertIn('SUBSTITUTE(C', bse[0])
        raw = wb["Raw_Holdings"]
        self.assertEqual((raw["B2"].value, raw["C2"].value), ("Client Name", "Sample Client"))
        self.assertEqual(ws["F6"].value, "Universal Sheet CMP")

    def test_deck(self):
        prs = Presentation(next(self.out.glob("*_Review.pptx")))
        self.assertEqual(len(prs.slides), 15)
        titles = [next(s.text_frame.text for s in sl.shapes if s.has_text_frame and s.name == "Text 1") for sl in list(prs.slides)[5:14]]
        self.assertEqual(titles[0], "Results Analysis Q1 FY27 - Banking & Financial Services")
        self.assertEqual(titles[1], "Results Analysis Q1 FY27 - Healthcare & Pharma (1/2)")
        s5 = {s.name: s.text_frame.text for s in prs.slides[4].shapes if s.has_text_frame}
        self.assertEqual((s5["Text 3"], s5["Text 5"], s5["Text 7"]), ("42 POSITIVE", "10 AVERAGE", "1 WEAK"))
        self.assertIn("PENDING NARRATIVE", s5["Text 9"])
        self.assertTrue(self.res.narrative["pending_fields"])
        ni = json.loads((self.out / "narrative_input.json").read_text(encoding="utf-8"))
        self.assertEqual(ni["rating_counts"]["rated"], 53)

    def test_weights_sorted_within_group(self):
        ni = json.loads((self.out / "narrative_input.json").read_text(encoding="utf-8"))
        w = {h["symbol"]: h["weight_pct"] for h in ni["holdings"]}
        for g in ni["groups"]:
            ws_ = [w[s] for s in g["symbols"]]
            self.assertEqual(ws_, sorted(ws_, reverse=True))


class Narrative(unittest.TestCase):
    def test_number_checker(self):
        ni = {"holdings": [{"weight_pct": 9.16, "result_line": "PAT +66% YoY (₹902cr)"}], "rating_counts": {"rated": 53}}
        nar = {"slide2": {"read_through": ["HBL 9.16% weight, Divi's PAT +66% (₹902cr), 53 rated, HBL 8.78%"]}}
        bad = [t for _, t in check_numbers(nar, ni, [])]
        self.assertEqual(bad, ["8.78"])

    def test_style(self):
        issues = style_issues({"slide3": {"cards": [{"heading": "3 themes", "body": "Crest view"}]}})
        self.assertTrue(any("number" in i for i in issues))
        self.assertTrue(any("CREST" in i for i in issues))

    @needs_samples
    def test_maneesh_narrative_flags_foreign_figures(self):
        with tempfile.TemporaryDirectory() as td:
            res = run(**ARGS, runs_dir=td, narrative=FX / "narrative_reference.json", render_png=False)
            self.assertEqual(res.status, "FAILED")
            nums = {t for _, t in res.narrative["unverified_numbers"]}
            self.assertIn("8.78", nums)   # reference-deck weight, not this run's price set

    @needs_samples
    def test_narrative_inserted_in_warn_mode(self):
        with tempfile.TemporaryDirectory() as td:
            res = run(**ARGS, runs_dir=td, narrative=FX / "narrative_reference.json",
                      settings_path=FX / "settings_warn.json", render_png=False)
            self.assertTrue(res.status.startswith("PASSED"), res.notes)
            self.assertFalse(res.narrative["pending_fields"])
            prs = Presentation(res.outputs["deck"])
            s4 = {s.name: s.text_frame.text for s in prs.slides[3].shapes if s.has_text_frame}
            self.assertEqual(s4["Text 4"], "HBL Engineering - 9.16% weight · +127.08%")


REAL_UNI = ROOT / "templates" / "universal-sheet.xlsx"


class RatingRulesTests(unittest.TestCase):
    def test_parse_and_rules(self):
        from crest.ratings import RatingRules, parse_result
        self.assertEqual(parse_result("Rev +4.1% YoY (Rs 3,301cr); PAT -18.5% YoY (Rs 507cr)"),
                         {"rev_yoy": 4.1, "pat_yoy": -18.5, "turnaround": False, "loss": False})
        f = parse_result("Rev +54.7% YoY (Rs 1,405cr); PAT Rs 71cr vs loss Rs 10cr in Q2 FY26")
        self.assertTrue(f["turnaround"] and not f["loss"])
        self.assertTrue(parse_result("Rev +25.9% YoY; Net loss Rs 1,879cr vs loss Rs 2,582cr")["loss"])
        r = RatingRules()
        self.assertEqual(r.evaluate("Rev +45.9% YoY; PAT +153.6% YoY")[:2], ("Excellent", "Bullish"))
        self.assertEqual(r.evaluate("Rev +39.6% YoY; PAT +0.0% YoY")[:2], ("Average", "Neutral"))
        self.assertEqual(r.evaluate("Rev -28.7% YoY; PAT -50.1% YoY")[:2], ("Weak", "Bearish"))
        self.assertEqual(r.evaluate("Rev +6.2% YoY; PAT Rs 746cr vs loss Rs 92cr")[:2], ("Good", "Bullish"))
        self.assertEqual(r.evaluate("")[0], None)

    def test_rules_reject_code(self):
        from crest.ratings import RatingRules
        with self.assertRaises(PipelineError):
            RatingRules({"rules": [{"rating": "Good", "when": "__import__('os')"}]})

    def test_alt_symbols(self):
        from crest.universal import alt_symbols
        self.assertEqual(alt_symbols("509496, ITDCEM"), ["509496", "ITDCEM"])
        self.assertEqual(alt_symbols(543748.0), ["543748"])


@unittest.skipUnless(REAL_UNI.exists() and HAVE_SAMPLES, "real Universal Sheet / samples not present")
class RealUniversalSheet(unittest.TestCase):
    ARGS = dict(universal=REAL_UNI, holdings=HOLD, template=TPL)

    def test_period_mismatch_stops(self):
        with tempfile.TemporaryDirectory() as td:
            res = run(**self.ARGS, period="Q1FY27", runs_dir=td, render_png=False)
            self.assertEqual(res.status, "FAILED")
            self.assertTrue(any("Q Result Period" in c.name and c.status == "FAIL" for c in res.checks))

    def test_q2_unapproved_rules_na(self):
        with tempfile.TemporaryDirectory() as td:
            res = run(**self.ARGS, period="Q2FY27", runs_dir=td, render_png=False)
            self.assertTrue(res.status.startswith("PASSED"), res.notes)
            st = {c.name: c.status for c in res.checks}
            self.assertEqual(st["Sector sanity"], "PASS")
            self.assertEqual(st["Views filled in the Universal Sheet"], "WARN")
            ni = json.loads((res.run_dir / "outputs" / "narrative_input.json").read_text(encoding="utf-8"))
            # unapproved rules rate nothing; any rating shown comes from the sheet itself (typed, or saved there
            # earlier by the Gemini step) - the real sheet is edited over time, so no fixed count
            snap = openpyxl.load_workbook(res.outputs["excel"], data_only=True)["Universal_Snapshot"]
            sources = {snap.cell(r, 17).value for r in range(2, snap.max_row + 1) if snap.cell(r, 11).value}
            self.assertNotIn("rule", sources)
            self.assertEqual([g["group"] for g in ni["groups"]][0], "Banking & Financial Services")
            self.assertTrue((res.run_dir / "outputs" / "rating_preview.md").exists())
            ws = openpyxl.load_workbook(res.outputs["excel"], data_only=True)["Portfolio"]
            self.assertIn(ws["N9"].value, ("", None))  # blank View shows blank, not 0

    def test_q2_approved_rules(self):
        with tempfile.TemporaryDirectory() as td:
            res = run(**self.ARGS, period="Q2FY27", runs_dir=td, render_png=False,
                      settings_path=FX / "settings_rules_approved.json")
            ni = json.loads((res.run_dir / "outputs" / "narrative_input.json").read_text(encoding="utf-8"))
            rc = ni["rating_counts"]
            # the real sheet is edited over time, so check consistency rather than fixed counts
            self.assertGreater(rc["rated"], 0)
            self.assertEqual(rc["excellent"] + rc["good"] + rc["average"] + rc["weak"], rc["rated"])
            self.assertEqual(rc["positive"], rc["excellent"] + rc["good"])


if __name__ == "__main__":
    unittest.main()

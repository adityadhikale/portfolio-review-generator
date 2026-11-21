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
from crest.util import PipelineError, normalize_symbol, parse_period, parse_ticker, to_number  # noqa: E402
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
_offline = mock.patch("crest.shared.load_env", return_value={})


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


class FakeGoogle:
    """Stands in for the Google Sheets API (AuthorizedSession): no network. Rows are served for any values read;
    every request is recorded."""

    class Resp:
        def __init__(self, status, payload):
            self.status_code, self._p, self.text = status, payload, json.dumps(payload)

        def json(self):
            return self._p

    def __init__(self, reads, status=200, columns=26, cells=None):
        self.reads = list(reads)          # successive answers to the full read (lists of rows)
        self.status = status
        self.columns = columns
        self.cells = cells or {}          # A1 -> value, for the check before writing
        self.calls = []

    def get(self, url, timeout=None):
        self.calls.append(("get", url, None))
        if isinstance(self.status, Exception):
            raise self.status
        if self.status != 200:
            return self.Resp(self.status, {"error": {"message": "nope"}})
        if "values:batchGet" in url:
            from urllib.parse import parse_qs, urlparse
            rngs = parse_qs(urlparse(url).query)["ranges"]
            return self.Resp(200, {"valueRanges": [{"values": [[self.cells.get(r.split("!")[1])]]} for r in rngs]})
        if "fields=sheets.properties" in url:
            return self.Resp(200, {"sheets": [{"properties": {"sheetId": 7, "title": "Universal",
                                                              "gridProperties": {"columnCount": self.columns}}}]})
        rows = self.reads.pop(0) if len(self.reads) > 1 else self.reads[0]
        return self.Resp(200, {"values": rows})

    def post(self, url, timeout=None, json=None):
        self.calls.append(("post", url, json))
        return self.Resp(200, {})


class GoogleSheetSource(unittest.TestCase):
    HDR = ["Name", "Ticker", "Sector", "Market Cap (Cr)", "CMP", "Views", "Result Line", "Rating", "Sentiment",
           "Management Commentary", "Q Result Period"]
    ROWS = [HDR, ["Aaa Ltd", "NSE:AAA", "Pharmaceuticals", 1200.5, 101.25, "Hold", "Rev +10% YoY", "", "", "", "Q2 FY27"],
            ["Bbb Ltd", "BOM:540565", "Banks", "#N/A", "#N/A", "", "", "Good", "Bullish", "", "Q2 FY27"]]

    def setUp(self):
        self.td = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.td, True)

    def _sheet(self, fake, **kw):
        from crest.sheet_source import GoogleSheet, UniversalSheet
        waits = []
        g = GoogleSheet("sheet-id-123456", "Universal", session=fake, sleep=waits.append, **kw)
        return UniversalSheet(g, self.td), waits

    def test_read_and_parse(self):
        from crest.universal import load_universe, result_periods
        us, waits = self._sheet(FakeGoogle([self.ROWS]))
        u = load_universe(us)
        self.assertEqual([r.key for r in u.rows], ["AAA", "540565"])
        self.assertEqual(u.by_key["AAA"].values["cmp"], 101.25)
        self.assertIsNone(to_number(u.by_key["540565"].values["cmp"]))      # #N/A -> missing price
        self.assertIs(u.origin, us)
        self.assertEqual(u.by_key["AAA"].notes, [])          # a blank Rating / Sentiment from Google is not "invalid"
        self.assertIsNone(us.data().rows[1][7])
        self.assertEqual(result_periods(us), ["Q2 FY27", "Q2 FY27"])
        self.assertEqual(len([c for c in us.source._session.calls if c[0] == "get"]), 1)   # read once
        self.assertEqual(waits, [])
        self.assertTrue(us.data().live)

    def test_loading_read_again_once(self):
        loading = [self.ROWS[0], self.ROWS[1][:4] + ["Loading..."] + self.ROWS[1][5:], self.ROWS[2]]
        us, waits = self._sheet(FakeGoogle([loading, self.ROWS]))
        self.assertEqual(us.data().rows[1][4], 101.25)
        self.assertEqual(waits, [15.0])
        self.assertTrue(any("Loading" in n for n in us.notes()))
        us2, waits2 = self._sheet(FakeGoogle([loading]))               # still loading: price counts as missing
        us2.data()
        self.assertEqual(len(waits2), 1)
        self.assertTrue(any(n.startswith("WARN") and "Loading" in n for n in us2.notes()))

    def test_snapshot_saved_once_and_used_when_google_is_down(self):
        us, _ = self._sheet(FakeGoogle([self.ROWS]))
        snap = us.save_snapshot("Q2FY27")
        self.assertEqual(snap.parent, self.td / "Q2FY27" / "snapshots")
        us_again, _ = self._sheet(FakeGoogle([self.ROWS]))
        self.assertEqual(us_again.save_snapshot("Q2FY27"), snap)        # same data -> same file
        down, _ = self._sheet(FakeGoogle([], status=ConnectionError("no network")))
        d = down.data()
        self.assertFalse(d.live)
        self.assertEqual((d.snapshot, d.rows[1][0]), (snap, "Aaa Ltd"))
        self.assertIn("snapshot", d.warning)
        from crest.sheet_source import ReadOnlySheet
        with self.assertRaises(ReadOnlySheet):                           # a snapshot is never written back
            down.write({2: {"H": "Good"}})
        busy, _ = self._sheet(FakeGoogle([], status=503))
        self.assertFalse(busy.data().live)

    def test_down_without_snapshot_and_access_refused(self):
        down, _ = self._sheet(FakeGoogle([], status=ConnectionError("no network")))
        with self.assertRaises(PipelineError) as cm:
            down.data()
        self.assertIn("no saved snapshot", str(cm.exception))
        refused, _ = self._sheet(FakeGoogle([], status=403))
        with self.assertRaises(PipelineError) as cm:
            refused.data()
        self.assertIn("shared with", str(cm.exception))

    def test_write_one_request_skips_moved_rows_and_widens_grid(self):
        fake = FakeGoogle([self.ROWS], columns=11, cells={"A2": "Aaa Ltd", "A3": "Someone Else Ltd"})
        us, _ = self._sheet(fake)
        us.data()
        n = us.write({1: {"L": "Rating Source"}, 2: {"H": "Good", "L": "AI · m · Q2FY27 · ab"}, 3: {"H": "Weak"}},
                     expect={2: ("A", "Aaa Ltd"), 3: ("A", "Bbb Ltd")})
        self.assertEqual(n, [1, 2])                                      # header + row 2; row 3 moved
        posts = [c for c in fake.calls if c[0] == "post"]
        self.assertEqual(posts[0][2]["requests"][0]["appendDimension"]["length"], 1)    # column L is new
        body = posts[1][2]
        self.assertEqual(body["valueInputOption"], "RAW")
        self.assertEqual({d["range"]: d["values"][0][0] for d in body["data"]},
                         {"'Universal'!L1": "Rating Source", "'Universal'!H2": "Good",
                          "'Universal'!L2": "AI · m · Q2FY27 · ab"})
        self.assertTrue(any("[3]" in n for n in us.notes()))

    def test_local_file_is_read_only_and_source_choice(self):
        from crest.sheet_source import GoogleSheet, LocalSheet, ReadOnlySheet, open_universal
        self.assertIsInstance(open_universal({"UNIVERSAL_SHEET_ID": "abc"}).source, GoogleSheet)
        self.assertIsInstance(open_universal({}).source, LocalSheet)
        self.assertIsInstance(open_universal({"UNIVERSAL_SHEET_ID": "abc"}, local_path="x.xlsx").source, LocalSheet)
        with self.assertRaises(ReadOnlySheet):
            LocalSheet("x.xlsx").write({2: {"A": 1}})
        with self.assertRaises(PipelineError) as cm:                     # no key file yet
            open_universal({"UNIVERSAL_SHEET_ID": "abc", "GOOGLE_CREDENTIALS_FILE": str(self.td / "k.json")}).data()
        self.assertIn("credentials file not found", str(cm.exception))

    def test_pipeline_write_back(self):
        """Gemini's ratings for blank cells go to the Google Sheet in one request; a typed rating is not touched."""
        from types import SimpleNamespace
        from crest.match import match_holdings
        from crest.holdings import Holding
        from crest.shared import RunResult, _write_back_to_sheet
        from crest.universal import load_universe
        fake = FakeGoogle([self.ROWS], cells={"A2": "Aaa Ltd"})
        us, _ = self._sheet(fake)
        uni = load_universe(us)
        matched, _ = match_holdings([Holding("AAA", "AAA", 1, 1), Holding("540565", "540565", 1, 1)], uni)
        a, b = matched[0].uni, matched[1].uni
        a.rating, a.sentiment, a.rating_source = "Excellent", "Bullish", "AI (Gemini)"
        b.rating_source = "Universal Sheet"                                   # typed in the sheet
        cache = SimpleNamespace(data={"AAA": {"model": "m", "basis": "ab", "reason": "Strong"}},
                                path=self.td / "ai_ratings.json")
        res = RunResult()
        _write_back_to_sheet(res, uni, matched, cache, parse_period("Q2FY27"))
        body = [c for c in fake.calls if c[0] == "post" and "values:batchUpdate" in c[1]][0][2]
        got = {d["range"]: d["values"][0][0] for d in body["data"]}
        self.assertEqual(got["'Universal'!L1"], "Rating Source")
        self.assertEqual((got["'Universal'!H2"], got["'Universal'!I2"], got["'Universal'!M2"]),
                         ("Excellent", "Bullish", "Strong"))
        self.assertEqual(got["'Universal'!L2"], "AI · m · Q2FY27 · ab")
        self.assertFalse(any(k.endswith("3") for k in got))
        self.assertTrue(any("Saved Gemini's Rating/Sentiment for 1" in n for n in res.notes))

    def test_ai_tag(self):
        from crest.ai_ratings import ai_tag, tag_basis
        self.assertEqual(tag_basis(ai_tag("m", "Q2FY27", "ab12")), "ab12")
        self.assertIsNone(tag_basis("CREST"))
        self.assertIsNone(tag_basis(None))

    @needs_samples
    def test_run_against_google(self):
        """A full run reading the Universal Sheet through the (fake) Google API: same result as the local file."""
        from crest.sheet_source import GoogleSheet, UniversalSheet
        from crest.universal import read_xlsx_rows
        _, rows = read_xlsx_rows(UNI)
        fake = FakeGoogle([[[c for c in r] for r in rows]])
        us = UniversalSheet(GoogleSheet("sheet-id-123456", session=fake, sleep=lambda s: None), self.td)
        res = run(**dict(ARGS, universal=us), runs_dir=self.td, render_png=False)
        self.assertTrue(res.status.startswith("PASSED"), res.notes)
        st = {c.name: c.status for c in res.checks}
        self.assertEqual(st["Universal Sheet read live from Google"], "PASS")
        self.assertTrue(Path(res.info["universal_snapshot"]).exists())
        self.assertIn("Data snapshot used", (res.run_dir / "report.md").read_text(encoding="utf-8"))
        ni = json.loads((res.run_dir / "outputs" / "narrative_input.json").read_text(encoding="utf-8"))
        self.assertEqual(ni["rating_counts"]["rated"], 53)
        res2 = run(**ARGS, step="portfolio", runs_dir=self.td, render_png=False)     # local file: no Google check
        self.assertNotIn("Universal Sheet read live from Google", {c.name for c in res2.checks})


@needs_samples
class GeminiOutage(unittest.TestCase):
    def test_deck_built_when_gemini_fails(self):
        from crest import narrative_api as gemini
        env = {"NARRATIVE_PROVIDER": "gemini", "NARRATIVE_API_KEY": "fake"}
        calls = []

        def down(self_, url, body):
            calls.append(1)
            raise PipelineError("Gemini API error 503 (Gemini is overloaded - try again later)")
        with tempfile.TemporaryDirectory() as td, mock.patch("crest.shared.load_env", return_value=env), \
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
        self.assertEqual(st["Every holding is in the Portfolio or in Exceptions"], "PASS")
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


def make_portfolio(path: Path, rows, headers=None, client="Sample Client", asof="08/10/2026", marker=True,
                   normalized=None) -> Path:
    """A Portfolio Excel shaped like Step A's: marker in A1, summary labels, headers on row 8, Total row.
    rows: dicts keyed by header name."""
    from crest.excel_build import MARKER, PORTFOLIO_COLS
    headers = headers or PORTFOLIO_COLS
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Portfolio"
    if marker:
        ws["A1"] = MARKER
    ws["E3"], ws["F3"], ws["E4"], ws["F4"] = "Client", client, "Data as of", asof
    for j, h in enumerate(headers, start=1):
        ws.cell(8, j, h)
    for i, r in enumerate(rows):
        for j, h in enumerate(headers, start=1):
            ws.cell(9 + i, j, r.get(h))
    ws.cell(9 + len(rows), headers.index("Name") + 1, "Total")
    if normalized is not None:
        hn = wb.create_sheet("Holdings_Normalized")
        hn.append(["Key", "Symbol / ISIN (as typed)"])
        for k in normalized:
            hn.append([k, k])
        hn.append([None, None, None, "Qty x Avg Price total"])
    wb.save(path)
    return path


class PortfolioReader(unittest.TestCase):
    """Step B reading holdings back from a Portfolio Excel (no Universal Sheet or samples needed)."""
    ROWS = [{"Key": "ABC", "Name": "Abc Ltd", "Ticker": "NSE:ABC", "Quantity Available": 10, "Average Price": 100},
            {"Key": "540565", "Name": "IndiGrid", "Ticker": "BSE:540565", "Quantity Available": 5, "Average Price": 120.5}]

    def setUp(self):
        self.td = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.td, True)

    def test_kind_detection(self):
        from crest.portfolio_reader import file_kind
        self.assertEqual(file_kind(make_portfolio(self.td / "p.xlsx", self.ROWS)), "portfolio")
        self.assertEqual(file_kind(make_portfolio(self.td / "old.xlsx", self.ROWS, marker=False)), "portfolio")
        self.assertEqual(file_kind(make_template(self.td / "t.xlsx", [(None, "A", None, 1, 1)])), "template")

    def test_as_built(self):
        from crest.portfolio_reader import read_portfolio_excel
        h = read_portfolio_excel(make_portfolio(self.td / "p.xlsx", self.ROWS, normalized=["ABC", "540565"]))
        self.assertEqual((h.client, h.asof, h.sheet), ("Sample Client", "08/10/2026", "Portfolio"))
        self.assertEqual([(x.symbol, x.quantity, x.avg_price) for x in h.holdings],
                         [("ABC", 10, 100), ("540565", 5, 120.5)])
        self.assertEqual(h.notes, [])

    def test_columns_moved_key_blank_row_deleted(self):
        from crest.portfolio_reader import read_portfolio_excel
        headers = ["Average Price", "Quantity Available", "Views", "Name", "Ticker", "Key"]   # re-ordered
        rows = [{"Key": "ABC", "Name": "Abc Ltd", "Quantity Available": 10, "Average Price": 100},
                {"Key": None, "Name": "New Co Ltd", "Ticker": "NSE:NEWCO", "Quantity Available": 3, "Average Price": 9},
                {"Key": "", "Name": "Typed Name Only", "Quantity Available": 1, "Average Price": 2}]
        h = read_portfolio_excel(make_portfolio(self.td / "p.xlsx", rows, headers=headers,
                                                normalized=["ABC", "540565"]))   # 540565 deleted by hand
        self.assertEqual([x.symbol for x in h.holdings], ["ABC", "NEWCO", "TYPEDNAMEONLY"])
        self.assertEqual(h.holdings[2].name, "Typed Name Only")
        self.assertTrue(any("Ticker (1) or Name (1)" in n for n in h.notes))
        self.assertEqual(set(h.step_a_holdings), {"ABC", "540565"})   # the deck run warns 540565 is gone

    def test_google_drive_round_trip(self):
        """Downloaded from Google Sheets: the date became a real date, formulas carry their values, Key blank."""
        from datetime import datetime
        from crest.portfolio_reader import read_portfolio_excel
        rows = [dict(self.ROWS[0], Key=None, Ticker="NSE:ABC"), self.ROWS[1]]
        h = read_portfolio_excel(make_portfolio(self.td / "p.xlsx", rows, asof=datetime(2026, 10, 8)))
        self.assertEqual(h.asof, "08/10/2026")
        self.assertEqual([x.symbol for x in h.holdings], ["ABC", "540565"])

    def test_duplicate_rows_merged(self):
        from crest.portfolio_reader import read_portfolio_excel
        rows = self.ROWS + [{"Key": "ABC", "Name": "Abc Ltd", "Quantity Available": 30, "Average Price": 200}]
        h = read_portfolio_excel(make_portfolio(self.td / "p.xlsx", rows))
        self.assertAlmostEqual(h.holdings[0].quantity, 40)
        self.assertAlmostEqual(h.holdings[0].avg_price, (10 * 100 + 30 * 200) / 40)

    def test_errors_name_the_cell(self):
        from crest.portfolio_reader import read_portfolio_excel
        rows = [dict(self.ROWS[0], **{"Quantity Available": "ten"}), dict(self.ROWS[1], **{"Average Price": 0})]
        with self.assertRaises(PipelineError) as cm:
            read_portfolio_excel(make_portfolio(self.td / "p.xlsx", rows, client=" ", asof="soon"))
        for cell in ("F9", "G10", "F3", "F4"):
            self.assertIn(cell, str(cm.exception))
        headers = ["Key", "Name", "Ticker", "Qty"]
        with self.assertRaises(PipelineError) as cm:
            read_portfolio_excel(make_portfolio(self.td / "q.xlsx", [], headers=headers))
        self.assertIn("Quantity Available", str(cm.exception))

    def test_portfolio_excel_cannot_make_a_portfolio(self):
        from crest.pipeline import read_holdings
        f = make_portfolio(self.td / "p.xlsx", self.ROWS)
        with self.assertRaises(PipelineError) as cm:
            read_holdings(f, "portfolio")
        self.assertIn("Portfolio Excel", str(cm.exception))
        self.assertEqual(read_holdings(f, "deck")[1], "portfolio")
        with self.assertRaises(PipelineError):
            run(period="Q1FY27", universal=f, holdings=f, step="nonsense")


@needs_samples
class Steps(unittest.TestCase):
    """Step A and Step B run separately, and Step B reading the Portfolio Excel (as built, and edited by hand)."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = Path(tempfile.mkdtemp(prefix="crest-steps-"))
        cls.both = run(**ARGS, runs_dir=cls.tmp / "both", render_png=False)
        cls.xlsx = cls.tmp / "Sample Client Portfolio.xlsx"
        shutil.copy2(cls.both.outputs["excel"], cls.xlsx)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    @staticmethod
    def _ni(res):
        return json.loads((res.run_dir / "outputs" / "narrative_input.json").read_text(encoding="utf-8"))

    def test_portfolio_only(self):
        res = run(**ARGS, step="portfolio", runs_dir=self.tmp / "a", render_png=False)
        self.assertTrue(res.status.startswith("PASSED"), res.notes)
        self.assertIn("excel", res.outputs)
        self.assertNotIn("deck", res.outputs)
        self.assertTrue(all(c.step in ("Shared setup", "Step A: Portfolio") for c in res.checks))
        rep = (res.run_dir / "report.md").read_text(encoding="utf-8")
        self.assertIn("## Checks: Step A: Portfolio", rep)
        self.assertNotIn("## Narrative", rep)
        ws = openpyxl.load_workbook(res.outputs["excel"])["Portfolio"]
        self.assertEqual(ws["A1"].value, "CREST Portfolio v1")
        self.assertTrue(ws.column_dimensions["A"].hidden)

    def test_deck_from_portfolio_excel_matches_both(self):
        res = run(**dict(ARGS, holdings=self.xlsx), step="deck", runs_dir=self.tmp / "b", render_png=False)
        self.assertTrue(res.status.startswith("PASSED"), res.notes)
        self.assertNotIn("excel", res.outputs)
        self.assertEqual(res.info["holdings_kind"], "Portfolio Excel")
        a, b = self._ni(self.both), self._ni(res)
        self.assertEqual(a["portfolio"], b["portfolio"])
        self.assertEqual({h["symbol"]: h["weight_pct"] for h in a["holdings"]},
                         {h["symbol"]: h["weight_pct"] for h in b["holdings"]})
        self.assertEqual(a["rating_counts"], b["rating_counts"])
        self.assertTrue(Path(res.outputs["deck"]).exists())
        self.assertIn("## Checks: Step B: Deck", (res.run_dir / "report.md").read_text(encoding="utf-8"))

    def test_deck_from_template(self):
        res = run(**ARGS, step="deck", runs_dir=self.tmp / "c", render_png=False)
        self.assertTrue(res.status.startswith("PASSED"), res.notes)
        self.assertEqual(self._ni(res)["portfolio"], self._ni(self.both)["portfolio"])

    def test_deck_from_edited_excel(self):
        """Staff edits: a row deleted, a quantity changed, a column moved, a row's Key cleared."""
        f = self.tmp / "edited.xlsx"
        wb = openpyxl.load_workbook(self.xlsx)
        ws = wb["Portfolio"]
        deleted = ws["A9"].value
        ws.delete_rows(9)
        changed = ws["A10"].value
        ws["F10"] = ws["F10"].value * 2
        cleared = ws["A11"].value
        ws["A11"] = None                                  # matched again by Name
        ws.move_range("F8:F80", cols=10)                  # Quantity Available moved to column P
        wb.save(f)
        res = run(**dict(ARGS, holdings=f), step="deck", runs_dir=self.tmp / "d", render_png=False)
        self.assertTrue(res.status.startswith("PASSED"), res.notes)
        before = {h["symbol"]: h for h in self._ni(self.both)["holdings"]}
        after = {h["symbol"]: h for h in self._ni(res)["holdings"]}
        self.assertNotIn(deleted, after)
        self.assertEqual(len(after), len(before) - 1)
        self.assertEqual(after[changed]["quantity"], before[changed]["quantity"] * 2)
        self.assertTrue(any("Name (1)" in n for n in res.notes))
        self.assertTrue(any(n.startswith("WARN") and deleted in n for n in res.notes))
        cleared_name = before[cleared]["name"]
        self.assertTrue(any(h["name"] == cleared_name for h in after.values()))

    def test_deck_run_keeps_published_excel(self):
        out = self.tmp / "out"
        run(**ARGS, step="portfolio", runs_dir=self.tmp / "e", render_png=False, publish_dir=out)
        run(**dict(ARGS, holdings=self.xlsx), step="deck", runs_dir=self.tmp / "e", render_png=False, publish_dir=out)
        names = {p.suffix for p in (out / "Sample Client" / "Q1FY27").iterdir()}
        self.assertTrue({".xlsx", ".pptx"} <= names)



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

"""Tests.  Run:  python -m unittest discover -s tests -v

Unit tests always run. End-to-end tests need the private sample files, which are NOT in git (client data):
samples/holdings.xlsx, samples/reference_portfolio.xlsx, samples/reference_deck_Q1FY27.pptx,
templates/Deck_-_template.pptx and the fixtures built from them by tools/build_fixture.py and
tools/extract_narrative.py. Without them those tests are skipped."""
from __future__ import annotations

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl
from pptx import Presentation

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from crest.brokers import parse_holdings  # noqa: E402
from crest.deck_fill import chunk  # noqa: E402
from crest.narrative import check_numbers, style_issues  # noqa: E402
from crest.pipeline import run  # noqa: E402
from crest.universal import load_universe  # noqa: E402
from crest.util import PipelineError, normalize_symbol, parse_period, parse_ticker  # noqa: E402

S = ROOT / "samples"
FX = ROOT / "tests" / "fixtures"
HOLD, TPL, UNI = S / "holdings.xlsx", ROOT / "templates" / "Deck_-_template.pptx", FX / "universal_sample.xlsx"
ARGS = dict(client="Sample Client", period="Q1FY27", asof="08/10/2026", universal=UNI, holdings=HOLD, template=TPL)
HAVE_SAMPLES = all(p.exists() for p in (HOLD, TPL, UNI, FX / "narrative_reference.json"))
needs_samples = unittest.skipUnless(HAVE_SAMPLES, "private sample files not present (not in git)")


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


class BrokerAdapter(unittest.TestCase):
    @needs_samples
    def test_sample_holdings(self):
        p = parse_holdings(HOLD)
        self.assertEqual(p.header_row, 3)          # title row, blank row, header
        self.assertEqual(len(p.holdings), 54)
        self.assertEqual(p.asof_in_file, "08/10/2026")
        self.assertEqual(len(p.dropped), 2)        # total row + count row

    def test_csv_duplicates_suffix_zero_qty(self):
        with tempfile.TemporaryDirectory() as td:
            f = Path(td) / "h.csv"
            f.write_text("Client report\nSYMBOL,QTY,AVG PRICE,LTP\nABC-EQ,10,100,110\nABC,30,200,110\nXYZ,0,50,40\n,,,\nTotal,,,\n",
                         encoding="utf-8")
            p = parse_holdings(f)
            self.assertEqual([h.symbol for h in p.holdings], ["ABC"])
            self.assertAlmostEqual(p.holdings[0].quantity, 40)
            self.assertAlmostEqual(p.holdings[0].avg_price, (10 * 100 + 30 * 200) / 40)
            self.assertTrue(any(d["reason"] == "zero quantity" for d in p.dropped))

    def test_other_layout_isin_and_broker_codes(self):
        """ICICI-style export: broker codes instead of NSE symbols, ISIN column, header on row 1, stray cells."""
        with tempfile.TemporaryDirectory() as td:
            f = Path(td) / "h.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["Stock Symbol", "Company Name", "ISIN Code", "Qty", "Average Cost Price", "Current Market Price",
                       "% Change over prev close", "Value At Cost", "Value At Market Price", "Unrealized Profit/Loss %"])
            ws.append(["MOTSUM", "SAMVARDHANA MOTHERSON INT LTD", "INE775A01035", 6000, 95.95, 154.1, -3.2, 575700, 924600, 60.6])
            ws.append(["NAVFLU", "NAVIN FLUORINE INTERNATIONAL L", "INE048G01026", 22, 4543.51, 8425, -0.77, 99957.22, 185350, 85.4])
            ws.append([None, None, None, None, None, None, None, 675657.22, 1109950, None])
            wb.save(f)
            p = parse_holdings(f)
            self.assertEqual(p.header_row, 1)
            self.assertEqual(p.column_map["isin"], "ISIN Code")
            self.assertEqual(p.column_map["avg_price"], "Average Cost Price")
            self.assertEqual(p.column_map["ltp"], "Current Market Price")
            self.assertEqual(p.column_map["invested_value"], "Value At Cost")
            self.assertEqual([h.isin for h in p.holdings], ["INE775A01035", "INE048G01026"])

    def test_keyword_headers_and_exclusions(self):
        with tempfile.TemporaryDirectory() as td:
            f = Path(td) / "h.csv"
            f.write_text("Scrip Name,ISIN No.,Long Term Qty,Free Qty,Avg Buy Rate,Closing Rate\n"
                         "ABC Ltd,INE000A01011,5,10,100,120\n", encoding="utf-8")
            p = parse_holdings(f)
            self.assertEqual(p.column_map["quantity"], "Free Qty")       # not 'Long Term Qty'
            self.assertEqual(p.column_map["avg_price"], "Avg Buy Rate")
            self.assertEqual(p.holdings[0].quantity, 10)

    def test_unknown_layout(self):
        with tempfile.TemporaryDirectory() as td:
            f = Path(td) / "h.csv"
            f.write_text("foo,bar\n1,2\n", encoding="utf-8")
            with self.assertRaises(PipelineError):
                parse_holdings(f)


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
        self.assertEqual(st["Invested value ties to broker file"], "PASS")
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
        self.assertEqual(raw["A1"].value, "Trading account: " + raw["A1"].value.split(": ", 1)[-1])

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


@needs_samples
class BrokerLtp(unittest.TestCase):
    def test_broker_ltp_weights(self):
        with tempfile.TemporaryDirectory() as td:
            res = run(**ARGS, runs_dir=td, price_source="broker_ltp", render_png=False)
            ni = json.loads((res.run_dir / "outputs" / "narrative_input.json").read_text(encoding="utf-8"))
            w = {h["symbol"]: h["weight_pct"] for h in ni["holdings"]}
            self.assertAlmostEqual(w["HBLENGINE"], 8.88, places=2)
            self.assertAlmostEqual(ni["portfolio"]["present_value"], 24574529.2, places=1)


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
    ARGS = dict(client="Sample Client", asof="08/10/2026", universal=REAL_UNI, holdings=HOLD, template=TPL)

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
            self.assertEqual(ni["rating_counts"]["rated"], 0)
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

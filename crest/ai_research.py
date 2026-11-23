"""AI research (Phase 4): fill blank or out-of-date Sector, Result Line and Management Commentary in the
Universal Sheet for the stocks held in the portfolio being run. Views are never touched.

Which cells
  * Sector: when blank.
  * Result Line and Management Commentary: when blank, or written by the AI earlier for an older quarter.
  * A cell CREST typed (non-blank, not listed in the row's Data Source) is never changed. A row whose typed
    result is for another quarter is left alone; the run's period check flags it.

How (config/settings.json "ai_research")
  tavily      a Tavily web search per stock (.env TAVILY_API_KEY), then one Gemini request per batch of
              ~10 stocks reads the search extracts. Works with a free Gemini key.
  api_search  one request per batch with the narrative API's own web search turned on (needs a paid key;
              "gemini" is accepted as the same).
  sector_only no web search: Gemini fills blank Sectors only; results stay "Results awaited".
  auto        tavily when TAVILY_API_KEY is set, else sector_only.   off: nothing.

Checks before anything is used
  * The quarter must be the run's period, else "results awaited".
  * Sector must be one the deck knows (config/sector_groups.json).
  * The Result Line is built here, in the house format, from the figures returned, and must be readable by the
    rating rules ("Rev +45.9% YoY (Rs 610cr); PAT +153.6% YoY (Rs 71cr)").
  * Tavily mode: the revenue / profit figures must appear in the cited search extracts, and every figure in the
    commentary too; otherwise the answer (or the commentary) is rejected.
  * Every result and commentary carries its source links.

Answers are saved in runs/<period>/ai_research.json: a stock's result for a quarter is never asked for twice;
"not published yet" is asked again after ai_research_retry_days. New answers are written to the Universal Sheet
(Google) with Data Source and Data Updated."""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from openpyxl.utils import get_column_letter

from .ai_ratings import RatingCache, canon
from .narrative_api import GeminiClient, resolve_ai
from .ratings import parse_result
from .sheet_source import ReadOnlySheet
from .universal import RATINGS, SENTIMENTS, value_number
from .util import CONFIG_DIR, PipelineError, load_json

FIELDS = ("sector", "result", "commentary")
FIELD_NAMES = {"sector": "Sector", "result": "Result Line", "commentary": "Management Commentary"}
COMMENTARY_MAX = 280
TAVILY_URL = "https://api.tavily.com/search"
EXTRACT_CHARS = 600       # per search result sent to Gemini
SECTOR_BATCH = 100        # sector-only stocks per request (tiny items, no search)
RESULTS_PER_SEARCH = 5


# --------------------------------------------------------------------------- Data Source tag
def research_tag(model: str, period: str, fields, sources) -> str:
    return " · ".join(["AI", model, period, ",".join(sorted(fields)), " | ".join(sources[:3])])


def tagged_fields(tag) -> set:
    """The fields an 'AI · model · period · fields · sources' Data Source says the AI wrote ({} when typed)."""
    s = str(tag or "").strip()
    if not s.upper().startswith("AI"):
        return set()
    parts = [p.strip() for p in s.split("·")]
    return {f for f in (parts[3].split(",") if len(parts) > 3 else []) if f in FIELDS}


def tagged_period(tag) -> str:
    parts = [p.strip() for p in str(tag or "").split("·")]
    return parts[2].replace(" ", "").upper() if len(parts) > 2 else ""


# --------------------------------------------------------------------------- what a stock needs
def _blank(v) -> bool:
    return v is None or not str(v).strip()


def needs(u, per) -> set:
    """Fields to research for one Universal row this period (empty for ETFs)."""
    if u.is_etf:
        return set()
    v = u.values
    owned = tagged_fields(v.get("data_source"))
    out = set()
    if _blank(v.get("sector")):
        out.add("sector")
    row_period = str(v.get("result_period") or "").replace(" ", "").upper()
    ai_old = "result" in owned and tagged_period(v.get("data_source")) != per.compact.upper()
    if row_period and row_period != per.compact.upper() and not ai_old:
        return out                     # CREST's result for another quarter: not touched (period check flags it)
    for f in ("result", "commentary"):
        if _blank(v.get(f)) or (f in owned and ai_old):
            out.add(f)
    return out


# --------------------------------------------------------------------------- saved answers
class ResearchCache:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.data = {}
        if self.path.exists():
            try:
                self.data = json.loads(self.path.read_text(encoding="utf-8"))
            except json.JSONDecodeError as e:
                raise PipelineError(f"{self.path} is not valid JSON: {e}") from e

    def get(self, key):
        return self.data.get(key)

    def put(self, key, entry):
        self.data[key] = {**self.data.get(key, {}), **entry}

    def save(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(dict(sorted(self.data.items())), indent=2, ensure_ascii=False),
                             encoding="utf-8")


def searched_note(per, day: str, why: str) -> str:
    """The AI Searched cell: 'Q2FY27 · 23/11/2026 · Q2 FY27 results not found yet'."""
    return " · ".join([per.compact, day, why])


def _searched_recently(cell, per, retry_days: float):
    """(day, why) when the AI Searched cell says this period was searched less than retry_days ago, else None."""
    parts = [p.strip() for p in str(cell or "").split("·")]
    if len(parts) < 2 or parts[0].replace(" ", "").upper() != per.compact.upper():
        return None
    return (parts[1], parts[2] if len(parts) > 2 else "") if _days_since(parts[1]) < retry_days else None


def _days_since(ddmmyyyy: str) -> float:
    try:
        return (datetime.now() - datetime.strptime(ddmmyyyy, "%d/%m/%Y")).days
    except (TypeError, ValueError):
        return 1e9


# --------------------------------------------------------------------------- Tavily
class TavilySearch:
    def __init__(self, api_key: str, post=None, timeout: int = 60):
        self.api_key, self.timeout = api_key, timeout
        self._post = post or self._http_post
        self.used = 0

    def search(self, query: str) -> list[dict]:
        r = self._post(TAVILY_URL, {"query": query, "search_depth": "basic", "max_results": RESULTS_PER_SEARCH,
                                    "topic": "general"})
        self.used += 1
        return [{"title": x.get("title") or "", "url": x.get("url") or "",
                 "published": x.get("published_date") or "", "text": (x.get("content") or "")[:EXTRACT_CHARS]}
                for x in (r.get("results") or []) if x.get("url")]

    def _http_post(self, url, body):
        import urllib.error
        import urllib.request
        req = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"), method="POST",
                                     headers={"Content-Type": "application/json",
                                              "Authorization": f"Bearer {self.api_key}"})
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            why = {401: "the TAVILY_API_KEY is not valid", 429: "too many searches - try again later",
                   432: "this month's free Tavily searches are used up", 433: "Tavily pay-as-you-go limit reached"}
            raise SearchStopped(f"Tavily {e.code}: {why.get(e.code, 'search failed')}") from e
        except (TimeoutError, OSError) as e:
            raise SearchStopped(f"Tavily could not be reached ({getattr(e, 'reason', e)})") from e


class SearchStopped(Exception):
    """The search service cannot be used for the rest of this run."""


def search_query(u, per) -> str:
    sym = u.ticker_code
    return (f"{u.values.get('name')} ({u.ticker_exchange}: {sym}) {per.label} quarterly results quarter ended "
            f"{per.quarter_end} revenue net profit year on year")


# --------------------------------------------------------------------------- the Gemini request
def _system(per, sectors, web: str) -> str:
    src = ("Use ONLY the numbered search extracts given with each company for figures, the quarter and the "
           "commentary - never outside knowledge or memory." if web == "tavily" else
           "Search the web for each company's results; use only what the search finds, never memory, for "
           "figures, the quarter and the commentary." if web == "gemini" else
           "There is no web search in this request: fill only the sector, and set status 'not_searched'.")
    return f"""You research Indian listed companies for CREST, an investment advisory, for a quarterly portfolio
review. Indian fiscal year: Q1 = Apr-Jun, Q2 = Jul-Sep, Q3 = Oct-Dec, Q4 = Jan-Mar; FY27 = April 2026 - March 2027.
The quarter wanted is {per.label}, the quarter ended {per.quarter_end}.

{src}

For every company in `stocks` fill only the fields listed in its `need`, and return ONE JSON object:
{{"stocks": [{{"symbol": "...", "status": "found|not_published|unclear|not_searched", "quarter": "{per.label}",
  "sector": "...", "revenue_cr": 0, "revenue_yoy_pct": 0, "pat_cr": 0, "pat_yoy_pct": 0,
  "pat_same_quarter_last_year_cr": 0, "commentary": "...", "sources": [1, 2]}}]}}

- sector: exactly one name from `allowed_sectors` (outside knowledge is fine for the sector only).
- status "found" only when the sources show results for {per.label} itself; if only an older quarter is out,
  use "not_published" and leave the figures null. "unclear" when the sources disagree or are not specific.
- Figures: consolidated if available, else standalone. revenue_cr = revenue from operations, pat_cr = net profit
  after tax, both in Rs crore (convert lakh / million if needed); *_yoy_pct = change against the same quarter last
  year, in percent. A loss is a negative pat_cr. Leave a figure null rather than guess.
- commentary: one or two factual sentences, at most {COMMENTARY_MAX - 30} characters, on management commentary or
  business highlights for the quarter: orders and order book, guidance, capacity, margins, asset quality, launches.
  Do not just repeat revenue and profit (the result line already has them) unless the sources hold nothing else.
  British spelling, no hype, no recommendation. Every figure in it must come from the sources. Write every amount
  as Rs ... cr in crore, the house style (e.g. "order book of Rs 2,100 cr"): convert lakh (/100), million (/10) and
  billion (x100); never use the rupee sign, lakh, million or billion. Example of the style: "Order inflow rose 70%
  YoY to Rs 824 cr; FY27 revenue guidance raised to Rs 2,900-3,000 cr."
- sources: the numbers of the extracts you used{'' if web == 'tavily' else ' (use [] when not given numbered extracts)'}.
- rating, sentiment, rating_reason: only for companies marked "rate": true whose status is "found"; judge from the
  figures and commentary you report, nothing else. rating = quality of THIS quarter's results, one of {RATINGS}:
  Excellent = strong growth in both revenue and profit, or a clear turnaround; Good = solid growth, minor soft spots;
  Average = flat or mixed, or growth offset by a meaningful concern; Weak = profit decline, a loss, margin collapse or
  a serious problem. sentiment = the forward-looking tone, one of {SENTIMENTS}. rating_reason = one plain sentence,
  at most 160 characters.
allowed_sectors: {json.dumps(sectors)}"""


def _indian(n: float) -> str:
    """12345678 -> '1,23,45,678' (whole rupees crore, Indian grouping, as in the sheet)."""
    s = str(int(round(abs(n))))
    head, tail = s[:-3], s[-3:]
    while len(head) > 2:
        tail = head[-2:] + "," + tail
        head = head[:-2]
    return (head + "," + tail) if head else tail


def _num(v):
    try:
        return None if v is None or v == "" else float(v)
    except (TypeError, ValueError):
        return None


def result_line(a: dict, per_prev_label: str) -> str | None:
    """House-format Result Line from the figures in one answer, or None when there is not enough."""
    rev, rev_y = _num(a.get("revenue_cr")), _num(a.get("revenue_yoy_pct"))
    pat, pat_y, prev = _num(a.get("pat_cr")), _num(a.get("pat_yoy_pct")), _num(a.get("pat_same_quarter_last_year_cr"))
    parts = []
    if rev_y is not None:
        parts.append(f"Rev {rev_y:+.1f}% YoY" + (f" (Rs {_indian(rev)}cr)" if rev else ""))
    if pat is not None and pat < 0:
        parts.append(f"Net loss Rs {_indian(pat)}cr" + (f" vs loss Rs {_indian(prev)}cr" if prev is not None and prev < 0
                                                         else (f" vs profit Rs {_indian(prev)}cr" if prev else "")))
    elif pat is not None and prev is not None and prev < 0:
        parts.append(f"PAT Rs {_indian(pat)}cr vs loss Rs {_indian(prev)}cr in {per_prev_label}")
    elif pat_y is not None:
        parts.append(f"PAT {pat_y:+.1f}% YoY" + (f" (Rs {_indian(pat)}cr)" if pat is not None else ""))
    line = "; ".join(parts)
    return line if line and parse_result(line) else None


_NUM = re.compile(r"\d[\d,]*(?:\.\d+)?")
# period labels and years are not figures to check: Q2 FY27, H1 FY27, 9M FY26, FY'27, 2026
_PERIODS = re.compile(r"\b(?:Q[1-4]|H[12]|9M)?\s*FY\s*'?\d{2,4}\b|\bQ[1-4]\b|\b(?:19|20)\d{2}\b", re.I)


def _numbers(text: str) -> list[float]:
    out = []
    for t in _NUM.findall(text or ""):
        try:
            out.append(float(t.replace(",", "")))
        except ValueError:
            pass
    return out


def _crore_pool(nums: list[float]) -> list[float]:
    """Source figures plus their crore equivalents, so a figure the AI converted to Rs crore (as the house style
    asks) still counts as found: lakh / 100, million / 10, billion x 100."""
    return nums + [n / 100 for n in nums] + [n / 10 for n in nums] + [n * 100 for n in nums]


def _found_in(x: float, pool: list[float]) -> bool:
    x = abs(x)
    return any(abs(p - x) <= max(0.6, 0.01 * x) for p in pool)


def _prev_label(per) -> str:
    m = re.match(r"Q([1-4]) FY(\d{2})", per.label)
    return f"Q{m.group(1)} FY{int(m.group(2)) - 1:02d}" if m else "the same quarter last year"


# --------------------------------------------------------------------------- one run
@dataclass
class Outcome:
    requests: int = 0
    searches: int = 0
    filled: list = field(default_factory=list)       # (key, field, value) applied this run (new or saved)
    notes: list = field(default_factory=list)
    skipped: list = field(default_factory=list)      # (symbol, why) not researched / rejected
    rated: int = 0                                     # stocks rated in the research request itself


def research(*, matched, per, settings, env, cache_path: Path, budget: int, post=None, tavily_post=None,
             say=print, step: str = "both", rate: bool = False) -> Outcome:
    """Fill what the held stocks need (in memory, on the UniRow values), from saved answers first and then from at
    most `budget` Gemini requests. Returns the Outcome; write_back() then saves the new values to the sheet.

    step: a Portfolio-only run fills Sector only (the Excel shows no result line or commentary), with no web search.
    rate: also ask Rating / Sentiment for researched stocks that have none, in the same request; the answers go to
    runs/<period>/ai_ratings.json, so the later ratings + narrative request skips them."""
    out = Outcome()
    mode = str(settings.get("ai_research", "auto")).lower()
    if mode == "off" or not resolve_ai(env):
        return out
    if mode == "auto":
        mode = "tavily" if env.get("TAVILY_API_KEY") else "sector_only"
    if mode == "api_search":
        mode = "gemini"
    if mode == "tavily" and not env.get("TAVILY_API_KEY"):
        out.notes.append("WARN: ai_research is 'tavily' but .env has no TAVILY_API_KEY - filling Sector only")
        mode = "sector_only"
    if step == "portfolio":
        mode = "sector_only"           # results and commentary are only needed for the deck
    groups = load_json(CONFIG_DIR / "sector_groups.json")
    sectors = sorted(groups["sector_to_group"])
    canon_sector = {s.lower(): s for s in sectors}
    cache = ResearchCache(cache_path)
    retry_days = float(settings.get("ai_research_retry_days", 3))
    today = datetime.now().strftime("%d/%m/%Y")

    # by weight, largest first, one row per Universal stock
    seen, todo = set(), []
    for m in sorted(matched, key=lambda m: -(m.holding.quantity * (value_number(m.uni, "cmp") or 0))):
        if m.uni.key in seen:
            continue
        seen.add(m.uni.key)
        need = needs(m.uni, per)
        if not need:
            continue
        if mode == "sector_only":
            need &= {"sector"}
        c = cache.get(m.uni.key) or {}
        if "sector" in need and c.get("sector"):
            _apply(m.uni, "sector", c["sector"], c, per, out)
            need.discard("sector")
        if need & {"result", "commentary"} and c.get("period") == per.compact:
            if c.get("status") == "found":
                for f in ("result", "commentary"):
                    if f in need and c.get(f):
                        _apply(m.uni, f, c[f], c, per, out)
                need -= {"result", "commentary"}
            elif c.get("status") in ("awaited", "rejected") and _days_since(c.get("searched_at")) < retry_days:
                why = c.get("reason") or c.get("status")
                out.skipped.append((m.holding.raw_symbol, f"asked on {c.get('searched_at')}: {why}"
                                    f"; asked again after {retry_days:g} day(s)"))
                need -= {"result", "commentary"}
                note = searched_note(per, c.get("searched_at"), why)
                if str(m.uni.values.get("ai_searched") or "").strip() != note:
                    m.uni.ai_searched_new = note            # the sheet keeps it too (a host forgets this file)
        if need & {"result", "commentary"}:
            seen_at = _searched_recently(m.uni.values.get("ai_searched"), per, retry_days)
            if seen_at:                                     # remembered in the Universal Sheet
                out.skipped.append((m.holding.raw_symbol, f"asked on {seen_at[0]}: {seen_at[1] or 'not usable'}; "
                                    f"asked again after {retry_days:g} day(s)"))
                need -= {"result", "commentary"}
        if need:
            todo.append((m, need))
    if not todo:
        return out

    batch = max(1, int(settings.get("ai_research_batch_size", 20)))
    gc = GeminiClient(api_key=env.get("NARRATIVE_API_KEY", ""), model=env.get("NARRATIVE_MODEL") or None,
                      post=post)
    searcher = TavilySearch(env["TAVILY_API_KEY"], post=tavily_post) if mode == "tavily" else None
    tav = searcher
    # stocks needing a web search go in batches (largest holdings first); sector-only stocks need no search and
    # ride along in the first request, so they rarely cost a request of their own
    searching = [t for t in todo if t[1] & {"result", "commentary"}]
    sector_only = [t for t in todo if not t[1] & {"result", "commentary"}]
    batches = [searching[i:i + batch] for i in range(0, len(searching), batch)]
    if batches and sector_only:
        batches[0] = batches[0] + sector_only[:SECTOR_BATCH]
        sector_only = sector_only[SECTOR_BATCH:]
    batches += [sector_only[i:i + SECTOR_BATCH] for i in range(0, len(sector_only), SECTOR_BATCH)]
    rated, rate_items = {}, []
    for b in batches[budget:]:
        for m, _ in b:
            out.skipped.append((m.holding.raw_symbol, "not researched this run (request limit "
                                "ai_max_requests_per_run); the next run continues"))
    for b in batches[:budget]:
        items, extracts = [], {}
        for m, need in b:
            it = {"symbol": m.uni.key, "name": m.uni.values.get("name"), "ticker": m.uni.ticker_crest,
                  "need": sorted(need)}
            if rate and need & {"result", "commentary"} and not (m.uni.rating and m.uni.sentiment):
                it["rate"] = True
            if tav is not None and need & {"result", "commentary"}:
                try:
                    found = tav.search(search_query(m.uni, per))
                except SearchStopped as e:
                    out.notes.append(f"WARN: {e}; results not researched for the rest of this run")
                    tav = None
                    found = []
                extracts[m.uni.key] = found
                it["search_extracts"] = [{"n": i + 1, **x} for i, x in enumerate(found)]
            items.append(it)
        web = "tavily" if mode == "tavily" else ("gemini" if mode == "gemini" else "none")
        say(f"  AI research: one request for {len(items)} stock(s) "
            f"({'with web search' if web != 'none' else 'sector only'}) ...")
        try:
            reply = gc.generate_json(_system(per, sectors, web),
                                     [{"role": "user", "parts": [{"text": json.dumps({"stocks": items}, ensure_ascii=False)}]}],
                                     temperature=0.1, tools=[{"google_search": {}}] if web == "gemini" else None)
        except PipelineError as e:
            out.notes.append(f"WARN: AI research request failed ({e}); these stocks are researched on a later run: "
                             + ", ".join(m.holding.raw_symbol for m, _ in b))
            out.requests += 1
            break
        out.requests += 1
        grounding = [c.get("web", {}).get("uri") for c in (gc.last_grounding.get("groundingChunks") or [])
                     if c.get("web", {}).get("uri")]
        answers = {str(a.get("symbol") or "").upper(): a for a in (reply.get("stocks") or []) if isinstance(a, dict)}
        for m, need in b:
            a = answers.get(m.uni.key.upper()) or {}
            _take(m, need, a, extracts.get(m.uni.key), grounding, web, gc.model, per, canon_sector, cache, today, out)
            v = m.uni.values
            rt, se = canon(a.get("rating"), RATINGS), canon(a.get("sentiment"), SENTIMENTS)
            if rate and rt and se and (v.get("result") or v.get("commentary")) and not (m.uni.rating and m.uni.sentiment):
                rated[m.uni.key] = {"rating": rt, "sentiment": se, "reason": str(a.get("rating_reason") or "").strip()}
                rate_items.append({"symbol": m.uni.key, "name": v.get("name"), "sector": v.get("sector"),
                                   "result": v.get("result"), "commentary": v.get("commentary")})
    out.searches = searcher.used if searcher is not None else 0
    cache.save()
    if rated:
        RatingCache(Path(cache_path).parent / "ai_ratings.json").store(rate_items, rated, gc.model)
        out.rated = len(rated)
    return out


def _apply(u, f, value, entry, per, out):
    u.values[f] = value
    old = u.values.get("data_source")
    owned = tagged_fields(old) | {f}
    old_src = [x.strip() for x in str(old).split("·")[4].split("|")] if len(str(old or "").split("·")) > 4 else []
    u.values["data_source"] = research_tag(entry.get("model", "ai"), per.compact, owned,
                                           entry.get("sources") or [x for x in old_src if x])
    if f == "result":
        u.values["result_period"] = per.label
    u.ai_written = getattr(u, "ai_written", set()) | {f}
    out.filled.append((u.key, f, value))


def _take(m, need, a, extracts, grounding, web, model, per, canon_sector, cache, today, out):
    u, sym = m.uni, m.holding.raw_symbol
    entry = {"period": per.compact, "searched_at": today, "model": model}
    if "sector" in need:
        sec = canon_sector.get(str(a.get("sector") or "").strip().lower())
        if sec:
            entry["sector"] = sec
            _apply(u, "sector", sec, entry, per, out)
        else:
            out.skipped.append((sym, f"sector '{a.get('sector')}' is not a sector the deck knows - left blank"))
    if not need & {"result", "commentary"}:
        cache.put(u.key, entry)
        return
    if web == "tavily":
        ext = {i + 1: x for i, x in enumerate(extracts or [])}
        used = [ext[n] for n in (a.get("sources") or []) if isinstance(n, int) and n in ext]
        sources = [x["url"] for x in used]
        pool = _crore_pool(_numbers(" ".join(x["title"] + " " + x["text"] for x in used)))
    else:
        sources, pool = list(dict.fromkeys(grounding))[:5], None
    status = str(a.get("status") or "unclear").lower()
    quarter = str(a.get("quarter") or "").replace(" ", "").upper()
    why = None
    if status != "found":
        why = {"not_published": f"{per.label} results not found yet", "unclear": "sources unclear"}.get(
            status, f"no answer ({status})")
    elif quarter != per.compact.upper():
        why = f"answer was for {a.get('quarter')!r}, not {per.label}"
    elif not sources:
        why = "no source links"
    if why:
        cache.put(u.key, {**entry, "status": "awaited" if status == "not_published" else "rejected", "reason": why})
        u.ai_searched_new = searched_note(per, today, why)
        out.skipped.append((sym, why))
        return
    entry.update(status="found", sources=sources)
    if "result" in need:
        line = result_line(a, _prev_label(per))
        key_figs = [x for x in (_num(a.get("revenue_cr")), _num(a.get("pat_cr"))) if x]
        if line is None:
            out.skipped.append((sym, "figures returned do not make a readable Result Line"))
        elif pool is not None and not any(_found_in(x, pool) for x in key_figs):
            out.skipped.append((sym, "revenue / profit figures not found in the cited sources - not used"))
            line = None
        if line:
            entry["result"] = line
            _apply(u, "result", line, entry, per, out)
    if "commentary" in need:
        cm = re.sub(r"\s+", " ", str(a.get("commentary") or "")).strip()
        cm = re.sub(r"₹\s*", "Rs ", cm)                                  # house style: Rs ... cr
        cm = re.sub(r"(\d)\s*(?:crores?|crs?\.?)\b", r"\1 cr", cm, flags=re.I)
        if len(cm) > COMMENTARY_MAX:
            cm = ""
            out.skipped.append((sym, "commentary too long - not used"))
        elif re.search(r"\b(?:lakhs?|lacs?|millions?|mn|billions?|bn)\b", cm, re.I):
            cm = ""
            out.skipped.append((sym, "commentary not in Rs crore (house style) - not used"))
        elif cm and pool is not None and not all(_found_in(x, pool) for x in _numbers(_PERIODS.sub(" ", cm)) if x >= 10):
            cm = ""
            out.skipped.append((sym, "commentary has figures not in the cited sources - not used"))
        if cm:
            entry["commentary"] = cm
            _apply(u, "commentary", cm, entry, per, out)
    if not entry.get("result") and not entry.get("commentary"):
        entry.update(status="rejected", reason="nothing usable in the answer")
        u.ai_searched_new = searched_note(per, today, entry["reason"])
    elif str(u.values.get("ai_searched") or "").strip():
        u.ai_searched_new = ""                              # found now: clear the earlier "searched" note
    cache.put(u.key, entry)


# --------------------------------------------------------------------------- Universal Sheet
def write_back(res_notes: list, uni, matched, per) -> int:
    """Write AI-filled cells (this run's answers and saved ones the sheet does not have yet) to the Google Sheet,
    with Data Source and Data Updated, and the AI Searched note of stocks whose results were searched but are not
    out yet. Returns the number of stock rows written."""
    rows = {}
    for m in matched:
        u = m.uni
        if (getattr(u, "ai_written", set()) or getattr(u, "ai_searched_new", None) is not None) and u.row not in rows:
            rows[u.row] = u
    if not rows or uni.origin is None:
        return 0
    ci = uni.col_index
    last = max((i for i, h in enumerate(uni.headers) if str(h or "").strip()), default=len(uni.headers) - 1)
    nxt = [last]

    def col(key, header, updates):
        if key in ci:
            return get_column_letter(ci[key] + 1)
        nxt[0] += 1
        ci[key] = nxt[0]
        updates.setdefault(uni.header_row, {})[get_column_letter(nxt[0] + 1)] = header
        return get_column_letter(nxt[0] + 1)

    updates, expect = {}, {}
    today = datetime.now().strftime("%d/%m/%Y")
    filled = 0
    for r, u in rows.items():
        cells = {}
        written = getattr(u, "ai_written", set())
        for f in written:
            cells[col(f, FIELD_NAMES[f], updates)] = u.values[f]
        if "result" in written:
            cells[col("result_period", "Q Result Period", updates)] = per.label
        if written:
            filled += 1
            cells[col("data_source", "Data Source", updates)] = u.values["data_source"]
            cells[col("data_updated", "Data Updated", updates)] = today
        if getattr(u, "ai_searched_new", None) is not None:
            cells[col("ai_searched", "AI Searched", updates)] = u.ai_searched_new
        updates[r] = cells
        expect[r] = (get_column_letter(ci["name"] + 1), u.values.get("name"))
    try:
        done = [r for r in uni.origin.write(updates, expect) if r != uni.header_row]
        res_notes += [n for n in uni.origin.notes() if n not in res_notes]
        if done and filled:
            res_notes.append(f"Saved AI-researched Sector / Result Line / Commentary for {filled} stock(s) into "
                             f"the Universal Sheet ({uni.source}), with Data Source and Data Updated")
        return len(done)
    except ReadOnlySheet as e:
        if filled:
            res_notes.append(f"AI-researched values for {filled} stock(s) were not written into the Universal Sheet "
                             f"({e}); they are kept in ai_research.json and written on the next run that reads the "
                             "Google Sheet")
    except Exception as e:  # noqa: BLE001 - never lose a run over the write-back
        res_notes.append(f"WARN: could not write AI-researched values into the Universal Sheet ({e}); they are kept "
                         "in ai_research.json")
    return 0


def ai_sourced(matched) -> list[dict]:
    """Every AI-written Sector / Result Line / Commentary among this run's holdings (this run's and earlier ones),
    for the report's 'review before sending' list."""
    rows, seen = [], set()
    for m in matched:
        u = m.uni
        if u.key in seen:
            continue
        seen.add(u.key)
        tag = u.values.get("data_source")
        fields = tagged_fields(tag)
        if not fields:
            continue
        parts = [p.strip() for p in str(tag).split("·")]
        for f in sorted(fields):
            rows.append({"symbol": m.holding.raw_symbol, "name": u.values.get("name") or "",
                         "field": FIELD_NAMES[f], "value": u.values.get(f) or "",
                         "new": f in getattr(u, "ai_written", set()), "period": parts[2] if len(parts) > 2 else "",
                         "sources": parts[4] if len(parts) > 4 else ""})
    return rows


def ai_rated(matched) -> list[dict]:
    """Every Rating / Sentiment the AI gave among this run's holdings (this run's and ones saved earlier), with the
    AI's one-line reason, for the Portfolio Excel's References sheet."""
    rows, seen = [], set()
    for m in matched:
        u = m.uni
        if u.key in seen or getattr(u, "rating_source", "") != "AI (Gemini)":
            continue
        seen.add(u.key)
        rows.append({"symbol": m.holding.raw_symbol, "name": u.values.get("name") or "", "rating": u.rating or "",
                     "sentiment": u.sentiment or "",
                     "reason": getattr(u, "ai_rating_reason", "") or u.values.get("rating_reason") or ""})
    return rows

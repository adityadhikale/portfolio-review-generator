"""One Gemini request per client run: rate the stocks that still need a Rating / Sentiment AND write the
narrative, together. Either part is left out when it is not needed (ratings cached, narrative.json already
saved); when neither is needed no request is made at all.

There are no automatic correction requests. The pipeline checks the narrative (figures, house style,
length); if something fails, the run says which field, and narrative.json can be edited by hand and the run
repeated (a saved narrative.json is never sent for again)."""
from __future__ import annotations

import json

from .ai_ratings import parse_ratings
from .narrative_api import GeminiClient
from .narrative import blank_template
from .universal import RATINGS, SENTIMENTS

SYSTEM = f"""You work for CREST, an Indian investment advisory and family office, on a quarterly client portfolio
review deck. Tone: institutional, measured, specific; no hype, no exclamation marks, no emojis, no promises of
returns. British spelling throughout (favour, prioritise, realise, programme, behaviour, centre). Always write the
brand as CREST (capitals).

The request has up to two tasks. Do only the tasks whose input is present, and return ONE JSON object:
{{"ratings": [{{"symbol": "...", "rating": "...", "sentiment": "...", "reason": "..."}}], "narrative": {{...}}}}
(omit "ratings" if there is no `rate` list; omit "narrative" if there is no `narrative_input`).

TASK A - ratings, for every company in `rate` (each has its quarterly result line and CREST's notes on
management commentary). Judge only from that text, not outside knowledge.
  rating - quality of THIS QUARTER's results, one of {RATINGS}:
    Excellent: strong growth in both revenue and profit (or a clear turnaround) with healthy margins.
    Good: solid growth and execution; minor soft spots only.
    Average: flat or mixed numbers, or growth offset by a meaningful concern (e.g. one-off gains inflating profit).
    Weak: profit decline, a loss, margin collapse, or a serious problem flagged in the commentary.
  sentiment - the forward-looking tone from the commentary and trajectory, one of {SENTIMENTS}. It may differ
    from the rating (a narrowing loss with a confident outlook can be Weak / Neutral).
  reason - one plain sentence, at most 160 characters, citing what decided it.
  If the text is too thin to judge, use Average / Neutral and say so in the reason.

TASK B - narrative: fill `template` (same keys, same list lengths) from `narrative_input`, which holds every
computed fact for the portfolio (holdings, weights, gains, sectors, deck groups, results, commentary, ratings
and CREST's views). Holdings listed in `rate` have no rating in narrative_input yet: use your Task A ratings
for them, but never state counts of ratings.
  1. Use ONLY figures that appear in narrative_input. Every number is checked; any other number (years,
     counts, outside knowledge) fails the run. Round only to the decimals shown, or fewer. When unsure, use words.
  2. Do not invent results, guidance or events. Thin data -> keep it short and general.
  3. Keep every field within its limit in `max_chars` (characters). Exact counts: slide2.themes 6, slide3.cards 6,
     slide4.cards 4; slide2.read_through 5 to 8 bullets. Never start a headline, title or heading with a
     number. slide4.cards[*].symbol must be a symbol from top_holdings (its heading is built automatically).
  4. A holding whose result line and commentary are both blank has no results data: do not describe its
     results, profitability, growth or outlook - only its weight, gain/loss and CREST's view.
  5. The Current View (Buy / Add / Hold / Sell) is CREST's own decision: reflect it faithfully, never recommend
     anything that contradicts it, and name Buy / Add holdings where the text talks about adding capital.
     Call it a "view" (e.g. "Sell view"), not a rating.
  6. This text goes to a client: never mention missing data, rating coverage, the Universal Sheet, or how the
     ratings were produced."""


# Per-holding fields Gemini needs to write the text. Quantities, prices, rupee amounts, tickers and market caps
# are left out: they are not needed for the sentences, make the request slower, and are client data.
_HOLDING_FIELDS = ("symbol", "display_name", "sector", "deck_group", "weight_pct", "gain_pct", "view", "rating",
                   "sentiment", "result_line", "management_commentary")
_PORTFOLIO_FIELDS = ("unrealized_pnl_pct", "holdings_count", "equity_count", "etf_count", "etfs")


def _client_safe(ni: dict) -> dict:
    """The part of narrative_input sent to Gemini: no client name, no amounts, only what the text needs.
    (The narrative is still checked against the full narrative_input, so nothing here can loosen the checks.)"""
    safe = {k: v for k, v in json.loads(json.dumps(ni)).items() if k not in ("narrative_schema", "rules")}
    if isinstance(safe.get("run"), dict):
        safe["run"].pop("client", None)
    if isinstance(safe.get("portfolio"), dict):
        safe["portfolio"] = {k: safe["portfolio"][k] for k in _PORTFOLIO_FIELDS if k in safe["portfolio"]}
    safe["holdings"] = [{k: h.get(k) for k in _HOLDING_FIELDS if h.get(k) not in (None, "")}
                        for h in safe.get("holdings", [])]
    safe["max_chars"] = {k: v["max_chars"] for k, v in (ni.get("narrative_schema") or {}).items()}
    return safe


def request_review(client: GeminiClient, *, narrative_input: dict | None, to_rate: list[dict]):
    """-> (ratings {symbol: {rating, sentiment, reason}}, narrative dict or None, notes [str]). One API call."""
    if not narrative_input and not to_rate:
        return {}, None, []
    parts = []
    if to_rate:
        send = [{k: t.get(k) for k in ("symbol", "name", "sector", "result", "commentary")} for t in to_rate]
        parts.append("rate:\n" + json.dumps(send, ensure_ascii=False))
    if narrative_input:
        parts.append("template:\n" + json.dumps(blank_template(), indent=1))
        parts.append("narrative_input:\n" + json.dumps(_client_safe(narrative_input), ensure_ascii=False))
    # generate_json: a reply that is not a readable JSON object moves the request to the next model in the list
    reply = client.generate_json(SYSTEM, [{"role": "user", "parts": [{"text": "\n\n".join(parts)}]}], temperature=0.3)

    notes = []
    want = {t["symbol"] for t in to_rate}
    ratings = parse_ratings(reply.get("ratings"), want)
    if want - set(ratings):
        notes.append(f"WARN: Gemini gave no valid rating for {sorted(want - set(ratings))}; they show n/a")
    narrative = None
    if narrative_input:
        narrative = reply.get("narrative")
        if not isinstance(narrative, dict) or not narrative:
            notes.append("WARN: Gemini returned no narrative; slides 2-5 and 15 stay pending")
            narrative = None
    return ratings, narrative, notes

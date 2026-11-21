"""Cache of Gemini's Rating / Sentiment per stock and quarter (runs/<period>/ai_ratings.json).

Precedence per stock (pipeline): value typed in the Universal Sheet > this AI rating > blank (n/a).
Views are never touched: they are CREST's own call and come only from the Universal Sheet.

Ratings are per stock and quarter, not per client, so every client run for the period shares this file.
A cached entry is reused while the stock's result line and commentary are unchanged; edit the file to
correct a rating, or delete an entry to have it rated again. The ratings themselves are requested together
with the narrative in a single Gemini call (crest/ai_review.py)."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path

from .universal import RATINGS, SENTIMENTS
from .util import PipelineError


def rating_basis(item: dict) -> str:
    """Short hash of a stock's result line + commentary: a saved rating is reused while this is unchanged."""
    return hashlib.sha1(f"{item.get('result') or ''}\n{item.get('commentary') or ''}".encode("utf-8")).hexdigest()[:16]


def ai_tag(model: str, period: str, basis: str) -> str:
    """Rating Source text for a rating the tool wrote: 'AI · model · period · basis'."""
    return f"AI · {model} · {period} · {basis}"


def tag_basis(tag) -> str | None:
    """The basis hash from an 'AI · model · period · basis' tag, or None when the cell is not an AI tag."""
    s = str(tag or "").strip()
    if not s.upper().startswith("AI"):
        return None
    return s.split("·")[-1].strip() if "·" in s else ""


def canon(v, allowed):
    s = str(v or "").strip().lower()
    return next((a for a in allowed if a.lower() == s), None)


class RatingCache:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.data = {}
        if self.path.exists():
            try:
                self.data = json.loads(self.path.read_text(encoding="utf-8"))
            except json.JSONDecodeError as exc:
                raise PipelineError(f"{self.path} is not valid JSON: {exc}") from exc

    def split(self, items: list[dict]) -> tuple[dict, list[dict], list[dict]]:
        """-> (cached {symbol: entry}, to_rate [items], no_data [items]). Items: {symbol, name, sector, result, commentary}."""
        cached, todo, nodata = {}, [], []
        for it in items:
            if not (it.get("result") or it.get("commentary")):
                nodata.append(it)
                continue
            c = self.data.get(it["symbol"])
            if c and c.get("basis") == rating_basis(it) and c.get("rating") in RATINGS and c.get("sentiment") in SENTIMENTS:
                cached[it["symbol"]] = c
            else:
                todo.append(it)
        return cached, todo, nodata

    def store(self, items: list[dict], got: dict, model: str) -> dict:
        """Save Gemini's answers for `items`; returns {symbol: entry} for the valid ones."""
        out = {}
        for it in items:
            r = got.get(it["symbol"])
            if r:
                out[it["symbol"]] = self.data[it["symbol"]] = {
                    "rating": r["rating"], "sentiment": r["sentiment"], "reason": str(r.get("reason") or "")[:200],
                    "basis": rating_basis(it), "model": model}
        if out:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.path.write_text(json.dumps(dict(sorted(self.data.items())), indent=2, ensure_ascii=False),
                                 encoding="utf-8")
        return out


def parse_ratings(rows, want: set) -> dict:
    """Gemini's 'ratings' list -> {symbol: {rating, sentiment, reason}} for valid entries of wanted symbols."""
    good = {}
    for r in rows if isinstance(rows, list) else []:
        if not isinstance(r, dict):
            continue
        s, rt, se = r.get("symbol"), canon(r.get("rating"), RATINGS), canon(r.get("sentiment"), SENTIMENTS)
        if s in want and rt and se:
            good[s] = {"rating": rt, "sentiment": se, "reason": str(r.get("reason") or "").strip()}
    return good

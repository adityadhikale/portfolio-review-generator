"""Narrative I/O.

The pipeline never writes analysis itself. It exports narrative_input.json (every computed fact for
the run plus the narrative.json schema with length limits), and reads narrative.json back.

A saved narrative.json is always used as-is (ManualProvider). When none exists and Gemini is selected
(settings.json -> narrative_provider "auto" with a key in .env, or "gemini"), crest/ai_review.py writes it in
the same single request as the ratings. The checker below runs on every narrative.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from .util import PipelineError

# key -> (description, max_chars). max_chars is sized to the template text box.
SCHEMA = {
    "slide2.headline": ("Title after '<period> Outlook - ' (slide 2). Must not start with a number.", 55),
    "slide2.read_through": ("List of 5-8 bullet strings for 'Our read-through'.", 1400),
    "slide2.themes": ("List of exactly 6 short theme names for the 'Themes we like' tiles.", 28),
    "slide2.portfolio_implication": ("One or two sentences after 'Portfolio implication:'.", 300),
    "slide3.title_suffix": ("Title after 'Market Outlook - ' (slide 3).", 50),
    "slide3.subtitle": ("One sentence under the slide 3 title.", 170),
    "slide3.cards": ("List of exactly 6 objects {heading, body} (sector cards).", 230),
    "slide4.title": ("Optional. Default: 'Industrial Compounders - the Core Return Engine'.", 60),
    "slide4.subtitle": ("One sentence under the slide 4 title.", 170),
    "slide4.cards": ("List of exactly 4 objects {symbol, body}. Heading is built from data: '<name> - <weight>% weight · <gain>%'.", 260),
    "slide4.portfolio_action": ("Numbered actions after 'Portfolio action'.", 420),
    "slide5.best": ("Text after 'Best results of the quarter:'.", 330),
    "slide5.weakest": ("Text after 'Weakest results of the quarter:'.", 330),
    "slide5.broad_theme": ("Text after 'Broad theme:'.", 330),
    "slide15.title_suffix": ("Optional. Title after 'Market Outlook - ' (slide 15). Default: 'Positioning for the "
                             "<panel.time_horizon>'.", 50),
    "slide15.our_view": ("Text after 'Our view:'.", 200),
    "slide15.investment_stance": ("Text after 'Investment stance:'.", 330),
    "slide15.time_horizon": ("Text after 'Time horizon:'.", 90),
    "slide15.key_themes": ("Text after 'Key themes:'.", 170),
    "slide15.screening_discipline": ("Text after 'Screening discipline:'.", 190),
    "slide15.panel.investment_stance": ("Side panel, 1-3 words.", 30),
    "slide15.panel.time_horizon": ("Side panel, 1-3 words.", 30),
    "slide15.panel.approach": ("Side panel, a few words.", 40),
    "slide15.panel.key_themes": ("Side panel, short list separated by ' · '.", 110),
}


def blank_template() -> dict:
    out: dict = {}
    for k in SCHEMA:
        cur = out
        parts = k.split(".")
        for p in parts[:-1]:
            cur = cur.setdefault(p, {})
        leaf = parts[-1]
        if leaf == "themes":
            cur[leaf] = [""] * 6          # the template has exactly 6 theme tiles
        elif leaf == "read_through":
            cur[leaf] = [""] * 5          # 5-8 bullets
        elif leaf == "cards":
            cur[leaf] = [{"heading": "", "body": ""}] * 6 if parts[0] == "slide3" else [{"symbol": "", "body": ""}] * 4
        else:
            cur[leaf] = ""
    return out


def get_path(d: dict, dotted: str):
    cur = d
    for p in dotted.split("."):
        if not isinstance(cur, dict) or p not in cur:
            return None
        cur = cur[p]
    return cur


class ManualProvider:
    """Reads narrative.json that a person drafted (e.g. with an AI chat) and pasted in."""

    def __init__(self, path: Path | None):
        self.path = path

    def get(self, _narrative_input: dict):
        if not self.path or not self.path.exists():
            return None
        try:
            with open(self.path, encoding="utf-8") as fh:
                return json.load(fh)
        except json.JSONDecodeError as exc:
            raise PipelineError(f"narrative.json is not valid JSON: {exc}") from exc


def ai_digest(narrative: dict) -> str:
    """A fingerprint of the narrative's text (keys starting with '_' left out). Saved with a narrative Gemini wrote,
    so a later run can tell whether anyone has edited the file since."""
    import hashlib
    body = {k: v for k, v in narrative.items() if not str(k).startswith("_")}
    return hashlib.sha1(json.dumps(body, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()[:16]


def untouched_ai(narrative: dict | None) -> bool:
    """True when the narrative is exactly as Gemini wrote it (not edited by CREST since)."""
    return bool(narrative) and narrative.get("_ai_digest") == ai_digest(narrative)


def resolve_provider_name(setting: str, env: dict) -> str:
    """'auto' -> 'gemini' when .env names Gemini and has a key, else 'manual'. Gemini's narrative is requested in
    the same single call as the ratings (crest/ai_review.py); ManualProvider always reads a saved narrative.json."""
    if setting != "auto":
        return setting
    name = (env.get("NARRATIVE_PROVIDER") or "").strip().lower()
    return name if name and name != "manual" and env.get("NARRATIVE_API_KEY") else "manual"


# --------------------------------------------------------------------------- number checker
_NUM = re.compile(r"(?<![A-Za-z])\d[\d,]*(?:\.\d+)?")


def _strings_and_numbers(obj, nums: list, strs: list):
    if isinstance(obj, bool) or obj is None:
        return
    if isinstance(obj, (int, float)):
        nums.append(float(obj))
    elif isinstance(obj, str):
        strs.append(obj)
    elif isinstance(obj, dict):
        for k, v in obj.items():
            if not str(k).startswith("_"):
                _strings_and_numbers(v, nums, strs)
    elif isinstance(obj, (list, tuple)):
        for v in obj:
            _strings_and_numbers(v, nums, strs)


def _tokens(text: str):
    # drop list markers like "1)" / "2." at the start or after whitespace
    cleaned = re.sub(r"(?:(?<=\s)|^)\d{1,2}[).](?=\s)", " ", text)
    return [m.group(0) for m in _NUM.finditer(cleaned)]


def allowed_numbers(narrative_input: dict, extra: list) -> list:
    nums, strs = [], []
    _strings_and_numbers(narrative_input, nums, strs)
    for s in strs:
        for t in _tokens(s):
            nums.append(float(t.replace(",", "")))
    nums += [float(x) for x in extra]
    return nums


def check_numbers(narrative: dict, narrative_input: dict, extra_allowed: list) -> list:
    """Return [(field, token)] for every figure in the narrative that is not in narrative_input.json.
    A token with d decimals is accepted if some input value rounds to it at d decimals."""
    pool = allowed_numbers(narrative_input, extra_allowed)
    by_dec: dict[int, set] = {}
    flagged = []

    def ok(tok: str) -> bool:
        val = float(tok.replace(",", ""))
        d = len(tok.split(".")[1]) if "." in tok else 0
        if d not in by_dec:
            by_dec[d] = {round(abs(x), d) for x in pool}
        return round(val, d) in by_dec[d]

    def walk(obj, path):
        if isinstance(obj, str):
            for t in _tokens(obj):
                if not ok(t):
                    flagged.append((path, t))
        elif isinstance(obj, dict):
            for k, v in obj.items():
                if k != "symbol" and not str(k).startswith("_"):      # '_note', '_ai_digest': not narrative text
                    walk(v, f"{path}.{k}" if path else k)
        elif isinstance(obj, list):
            for i, v in enumerate(obj):
                walk(v, f"{path}[{i}]")

    walk(narrative, "")
    return flagged


CARD_HEADING_LIMIT = 34


def _limit(path: str) -> int | None:
    base = re.sub(r"\[\d+\]", "", path)
    if base.endswith(".heading"):
        return CARD_HEADING_LIMIT
    if base.endswith(".body"):
        base = base.rsplit(".", 1)[0]
    if base.endswith(".read_through"):
        return 260  # per bullet; the whole list is limited by the box (checked in QA)
    return SCHEMA.get(base, ("", None))[1]


def style_issues(narrative: dict) -> list:
    """House-style checks on narrative text: CREST in caps; headings must not start with a number."""
    issues = []

    def walk(obj, path):
        if isinstance(obj, str):
            if re.search(r"\b(?!CREST\b)[Cc][Rr][Ee][Ss][Tt]\b", obj):
                issues.append(f"{path}: brand must be written 'CREST'")
            if any(path.endswith(h) for h in ("headline", "title_suffix", "title", "heading")) and re.match(r"^\s*\d", obj):
                issues.append(f"{path}: heading starts with a number")
            lim = _limit(path)
            if lim and len(obj) > lim:
                issues.append(f"{path}: {len(obj)} chars, longer than the box allows (~{lim})")
        elif isinstance(obj, dict):
            for k, v in obj.items():
                if not str(k).startswith("_"):
                    walk(v, f"{path}.{k}" if path else k)
        elif isinstance(obj, list):
            for i, v in enumerate(obj):
                walk(v, f"{path}[{i}]")

    walk(narrative, "")
    return issues

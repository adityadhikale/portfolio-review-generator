"""Minimal Google Gemini client (REST, standard library only).

Key and models come from .env: NARRATIVE_API_KEY, and NARRATIVE_MODEL as one model or a comma-separated list
in order of preference. When a model is busy (overloaded / server error), does not answer within the timeout,
is over its quota, or does not exist for this key, the SAME request goes straight to the next model - no
waiting. If every model in the list fails that way, it waits 30 seconds and tries the list once more, then
stops with a readable message. Errors that another model would not fix (bad key, bad request) stop at once."""
from __future__ import annotations

import json
import time

from .util import PipelineError

# Used when NARRATIVE_MODEL is blank. A Pro model first (works on a paid key; a free key gets 'quota' at once
# and moves on), then fast Flash models.
DEFAULT_MODELS = ("gemini-pro-latest", "gemini-3.5-flash", "gemini-flash-latest", "gemini-flash-lite-latest")
URL = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
SWITCH_CODES = (404, 429, 500, 502, 503, 504)  # another model may work
PASS_WAIT = 30  # seconds before the second pass over the list


class ModelUnavailable(Exception):
    """This model cannot answer right now; another model might."""


def resolve_ai(env: dict) -> bool:
    """True when .env names Gemini and has a key."""
    return (env.get("NARRATIVE_PROVIDER") or "").strip().lower() == "gemini" and bool(env.get("NARRATIVE_API_KEY"))


def parse_json_object(text: str) -> dict:
    """The JSON object in a model reply, tolerating ```json fences or text around it. Raises ValueError."""
    s = text.strip()
    if s.startswith("```"):
        s = s.split("\n", 1)[1] if "\n" in s else ""
        s = s.rsplit("```", 1)[0]
    a = s.find("{")
    if a < 0:
        raise ValueError("no JSON object in the reply")
    s = s[a:]
    obj, err = None, None
    for cand in (s, s[:s.rfind("}") + 1], s + _missing_closers(s)):
        try:
            obj = json.loads(cand)
            break
        except json.JSONDecodeError as e:
            err = err or e
    if obj is None:
        raise ValueError(f"invalid JSON ({err.msg} at character {err.pos})")
    if not isinstance(obj, dict):
        raise ValueError(f"reply is a JSON {type(obj).__name__}, not an object")
    return obj


def _missing_closers(s: str) -> str:
    """The '}' / ']' needed to close brackets left open at the end of s (strings are skipped). Repairs a reply
    whose model forgot its last closing bracket(s); anything else still fails json.loads."""
    stack, in_str, esc = [], False, False
    for ch in s:
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
        elif ch == '"':
            in_str = True
        elif ch in "{[":
            stack.append("}" if ch == "{" else "]")
        elif ch in "}]" and stack:
            stack.pop()
    return "" if in_str else "".join(reversed(stack))


def parse_models(model) -> list[str]:
    if isinstance(model, (list, tuple)):
        names = [str(m).strip() for m in model]
    else:
        names = [m.strip() for m in str(model or "").split(",")]
    names = [m for m in names if m]
    return names or list(DEFAULT_MODELS)


class GeminiClient:
    def __init__(self, *, api_key: str, model=None, timeout: int = 120, post=None, sleep=time.sleep):
        if not api_key:
            raise PipelineError("Gemini is selected but NARRATIVE_API_KEY is blank in .env.")
        self.api_key, self.models, self.timeout = api_key, parse_models(model), timeout
        self.model = self.models[0]       # after generate(): the model that answered
        self.tried: list[str] = []        # "model: reason" for every model that could not answer
        self.last_bad_reply = ""          # raw text of the last unreadable reply (saved by the pipeline)
        self.last_grounding: dict = {}    # groundingMetadata of the last answer (web search sources), if any
        self._post = post or self._http_post  # post(url, body) -> response dict; replaced in tests
        self._sleep = sleep

    def generate_json(self, system: str, contents: list, temperature: float = 0.4, tools=None) -> dict:
        """Like generate(), but the reply must be a JSON object. A model whose reply cannot be read as one is
        treated like a busy model: the same request goes to the next model in the list."""
        return self.generate(system, contents, temperature, parse=parse_json_object, tools=tools)

    def generate(self, system: str, contents: list, temperature: float = 0.4, parse=None, tools=None):
        """One request that must return JSON text, sent to the first model in the list that can answer.
        parse(text) may convert the reply; if it raises ValueError the next model is tried.
        tools: e.g. [{"google_search": {}}] (web search; needs a paid key). With tools the JSON reply mode is off
        (the API does not combine them), so the prompt must ask for JSON and parse() reads it."""
        gen = {"temperature": temperature}
        if not tools:
            gen["responseMimeType"] = "application/json"
        body = {"systemInstruction": {"parts": [{"text": system}]}, "contents": contents, "generationConfig": gen}
        if tools:
            body["tools"] = tools
        for attempt in range(2):
            if attempt:
                print(f"  All Gemini models busy; trying the list again in {PASS_WAIT}s ...", flush=True)
                self._sleep(PASS_WAIT)
            for m in self.models:
                try:
                    resp = self._post(URL.format(model=m), body)
                except ModelUnavailable as e:
                    self.tried.append(f"{m}: {e}")
                    print(f"  Gemini {m} unavailable ({e}); switching model ...", flush=True)
                    continue
                text = self._text(resp)
                self.last_grounding = (resp.get("candidates") or [{}])[0].get("groundingMetadata") or {}
                if parse is not None:
                    try:
                        out = parse(text)
                    except ValueError as e:
                        self.tried.append(f"{m}: unusable reply ({e})")
                        self.last_bad_reply = text
                        print(f"  Gemini {m} gave an unusable reply ({e}); switching model ...", flush=True)
                        continue
                    self.model = m
                    return out
                self.model = m
                return text
        raise PipelineError("No Gemini model could answer (" + "; ".join(self.tried[-len(self.models):])
                            + "). Try again later, or change NARRATIVE_MODEL in .env.")

    def _text(self, resp: dict) -> str:
        fb = resp.get("promptFeedback", {})
        if fb.get("blockReason"):
            raise PipelineError(f"Gemini refused the request: {fb['blockReason']}")
        cands = resp.get("candidates") or []
        if not cands:
            raise PipelineError(f"Gemini returned no answer: {json.dumps(resp)[:300]}")
        parts = cands[0].get("content", {}).get("parts", [])
        text = "".join(p.get("text", "") for p in parts if not p.get("thought"))
        if not text.strip():
            raise PipelineError(f"Gemini returned an empty answer (finishReason {cands[0].get('finishReason')})")
        return text

    def _http_post(self, url: str, body: dict) -> dict:
        import urllib.error
        import urllib.request
        req = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"), method="POST",
                                     headers={"Content-Type": "application/json", "x-goog-api-key": self.api_key})
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            try:
                msg = json.loads(e.read().decode("utf-8")).get("error", {}).get("message", "")
            except Exception:  # noqa: BLE001
                msg = ""
            if e.code in SWITCH_CODES:
                label = {404: "not available for this key", 429: "over quota"}.get(e.code, "busy")
                raise ModelUnavailable(f"{e.code} {label}") from e
            hint = {400: "check the request / API key", 401: "the API key is not valid",
                    403: "the API key is not allowed to use Gemini"}
            raise PipelineError(f"Gemini API error {e.code} ({hint.get(e.code, 'see message')}): {msg}") from e
        except (TimeoutError, OSError) as e:  # read timeout, dropped connection (URLError is an OSError)
            raise ModelUnavailable(f"no answer: {getattr(e, 'reason', e)}") from e

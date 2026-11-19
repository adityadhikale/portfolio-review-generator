"""Deterministic Rating / Sentiment from the Result Line (config/rating_rules.json).

Precedence per stock: value typed in the Universal Sheet > rule (only if approved) > blank (n/a).
Rule conditions are tiny expressions over: rev_yoy, pat_yoy (floats or None), turnaround, loss (bools).
They are evaluated with a restricted evaluator (no builtins, names only from the parsed result)."""
from __future__ import annotations

import ast
import re

from .util import CONFIG_DIR, PipelineError, load_json

_REV = re.compile(r"\bRev(?:enue)?\b[^;%]*?([+-]?\d+(?:\.\d+)?)%", re.I)
_PAT = re.compile(r"\bPAT\b\s*([+-]?\d+(?:\.\d+)?)%", re.I)
_TURN = re.compile(r"\bPAT\b[^;]*\bvs\s+(?:a\s+)?loss", re.I)
_LOSS = re.compile(r"\bnet\s+loss\b|\bPAT\s+loss\b|\bloss\s+of\b", re.I)


def parse_result(line: str) -> dict | None:
    if not line or not str(line).strip():
        return None
    s = str(line)
    rev = _REV.search(s)
    pat = _PAT.search(s)
    turnaround = bool(_TURN.search(s)) and not _LOSS.search(s)
    loss = bool(_LOSS.search(s))
    if not rev and not pat and not turnaround and not loss:
        return None
    return {"rev_yoy": float(rev.group(1)) if rev else None, "pat_yoy": float(pat.group(1)) if pat else None,
            "turnaround": turnaround, "loss": loss}


_ALLOWED = (ast.Expression, ast.BoolOp, ast.And, ast.Or, ast.UnaryOp, ast.Not, ast.USub, ast.Compare, ast.Name,
            ast.Load, ast.Constant, ast.Lt, ast.LtE, ast.Gt, ast.GtE, ast.Eq, ast.NotEq)


def _compile(expr: str):
    tree = ast.parse(expr, mode="eval")
    for node in ast.walk(tree):
        if not isinstance(node, _ALLOWED):
            raise PipelineError(f"rating_rules.json: unsupported syntax in '{expr}'")
    return compile(tree, "<rule>", "eval")


class RatingRules:
    def __init__(self, cfg: dict | None = None):
        self.cfg = cfg or load_json(CONFIG_DIR / "rating_rules.json")
        self.approved = bool(self.cfg.get("approved"))
        self.rules = [(r["rating"], r["when"], _compile(r["when"]), r.get("note", "")) for r in self.cfg["rules"]]
        self.sent = self.cfg.get("sentiment_from_rating", {})

    def evaluate(self, result_line: str):
        """-> (rating, sentiment, explanation) or (None, None, reason)."""
        f = parse_result(result_line)
        if f is None:
            return None, None, "result line not parseable"
        env = dict(f)
        for rating, when, code, note in self.rules:
            try:
                if eval(code, {"__builtins__": {}}, env):  # noqa: S307 - AST-restricted above
                    return rating, self.sent.get(rating), f"{when} ({note})"
            except TypeError:  # comparison with None (field missing) -> rule does not apply
                continue
        return None, None, "no rule matched"

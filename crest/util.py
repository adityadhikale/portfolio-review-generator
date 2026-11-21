"""Shared helpers: config loading, normalisation, period parsing."""
from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = ROOT / "config"


class PipelineError(Exception):
    """A readable, user-facing failure. The message is shown as-is and written to report.md."""


def load_json(path: Path) -> dict:
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def load_env(path: Path | None = None) -> dict:
    """KEY=VALUE pairs from .env (blank lines and # comments ignored, surrounding quotes removed).
    A variable already set in the real environment wins over the file."""
    out: dict = {}
    p = path or ROOT / ".env"
    if p.exists():
        for line in p.read_text(encoding="utf-8-sig").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            v = v.strip()
            if len(v) >= 2 and v[0] == v[-1] and v[0] in "'\"":
                v = v[1:-1]
            out[k.strip()] = v
    for k in list(out):
        if os.environ.get(k):
            out[k] = os.environ[k]
    for k in ("NARRATIVE_PROVIDER", "NARRATIVE_API_KEY", "NARRATIVE_MODEL"):
        if k not in out and os.environ.get(k):
            out[k] = os.environ[k]
    return out


# --------------------------------------------------------------------------- symbols / names
_SYMBOL_SUFFIXES = ("-EQ", "-BE", "-BZ", "-SM", "-ST", ".NS", ".BO", "-E1")
_EXCHANGE_PREFIX = re.compile(r"^\s*(NSE|BSE|BOM)\s*:\s*", re.I)


def normalize_symbol(raw) -> str:
    """Trim, uppercase, drop an exchange prefix and broker series suffixes (-EQ, .NS ...)."""
    if raw is None:
        return ""
    s = str(raw).strip().upper()
    if isinstance(raw, float) and raw.is_integer():  # BSE codes read as floats
        s = str(int(raw))
    s = _EXCHANGE_PREFIX.sub("", s)
    s = re.sub(r"\s+", "", s)
    changed = True
    while changed:
        changed = False
        for suf in _SYMBOL_SUFFIXES:
            if s.endswith(suf) and len(s) > len(suf):
                s = s[: -len(suf)]
                changed = True
    return s


_NAME_NOISE = re.compile(r"\b(limited|ltd|private|pvt|the|inc|co|corporation|corp|india|\(india\))\b", re.I)


def normalize_name(raw) -> str:
    """Lowercase, drop punctuation and corporate suffixes. Used for exact-equality matching only."""
    if not raw:
        return ""
    s = str(raw).lower().replace("&", " and ")
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    s = _NAME_NOISE.sub(" ", s)
    return re.sub(r"\s+", " ", s).strip()


def short_display_name(name: str) -> str:
    """Name minus 'Limited/Ltd/Private Limited'. Nothing else is altered."""
    if not name:
        return ""
    s = re.sub(r"\s+(Private\s+Limited|Pvt\.?\s+Ltd\.?|Limited|Ltd\.?)\s*$", "", str(name).strip(), flags=re.I)
    return s.strip()


def parse_ticker(raw_ticker, exchange_hint=None) -> tuple[str, str, bool]:
    """Return (exchange, code, exchange_was_defaulted).

    'NSE:X' -> NSE; 'BSE:X' and 'BOM:X' -> BSE; bare -> exchange_hint or NSE (flagged)."""
    t = "" if raw_ticker is None else str(raw_ticker).strip()
    if isinstance(raw_ticker, float) and raw_ticker.is_integer():
        t = str(int(raw_ticker))
    m = re.match(r"^\s*(NSE|BSE|BOM)\s*:\s*(.+)$", t, re.I)
    if m:
        ex = m.group(1).upper()
        return ("BSE" if ex in ("BSE", "BOM") else "NSE"), m.group(2).strip().upper(), False
    hint = (str(exchange_hint).strip().upper() if exchange_hint else "")
    if hint in ("NSE", "BSE", "BOM"):
        return ("BSE" if hint in ("BSE", "BOM") else "NSE"), t.upper(), False
    if t.isdigit():  # a bare numeric code can only be a BSE scrip code
        return "BSE", t, True
    return "NSE", t.upper(), True


# --------------------------------------------------------------------------- periods
@dataclass
class Period:
    raw: str
    label: str        # "Q1 FY27"
    compact: str      # "Q1FY27"
    quarter_end: str  # "30 Jun 2026" or ""
    is_quarter: bool


_Q_END = {1: (6, 30), 2: (9, 30), 3: (12, 31), 4: (3, 31)}
_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def parse_period(raw: str) -> Period:
    m = re.match(r"^\s*Q([1-4])\s*FY\s*'?(\d{2}|\d{4})\s*$", raw, re.I)
    if not m:
        return Period(raw, raw.strip(), raw.strip().replace(" ", ""), "", False)
    q = int(m.group(1))
    fy = int(m.group(2)) % 100
    fy_end_year = 2000 + fy
    year = fy_end_year if q == 4 else fy_end_year - 1
    mo, d = _Q_END[q]
    return Period(raw, f"Q{q} FY{fy:02d}", f"Q{q}FY{fy:02d}", f"{d} {_MONTHS[mo - 1]} {year}", True)


def safe_filename(s: str) -> str:
    return re.sub(r'[<>:"/\\|?*]+', "_", s).strip() or "client"


def to_number(v):
    """Excel/CSV cell -> float or None. Accepts '1,234.50', '₹ 12', '(5)'."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("₹", "").replace("Rs.", "").replace("Rs", "").strip()
    if not s or s in ("-", "--", "#N/A", "N/A", "NA"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try:
        f = float(s)
    except ValueError:
        return None
    return -f if neg else f

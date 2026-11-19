"""Match broker holdings to Universal Sheet rows. Never guesses: ISIN (when the broker file has one), then exact
symbol, then ticker code / Alt Symbols, then exact normalised name (only if unique). Anything else goes to Exceptions."""
from __future__ import annotations

from dataclasses import dataclass

from .brokers import Holding
from .universal import UniRow, Universe
from .util import normalize_name


@dataclass
class Matched:
    holding: Holding
    uni: UniRow
    method: str  # "isin" | "symbol" | "ticker code" | "name"


@dataclass
class MatchException:
    kind: str
    symbol: str
    detail: str
    holding: Holding | None = None


def match_holdings(holdings: list[Holding], uni: Universe):
    matched, exceptions = [], []
    for h in holdings:
        if h.isin and h.isin in uni.by_isin:
            cands = uni.by_isin[h.isin]
            if len(cands) == 1:
                matched.append(Matched(h, cands[0], "isin"))
                continue
            exceptions.append(MatchException("Ambiguous", h.raw_symbol,
                                             f"ISIN {h.isin} appears on Universal rows {[c.row for c in cands]}", h))
            continue
        if h.symbol in uni.duplicates:
            exceptions.append(MatchException("Ambiguous", h.raw_symbol,
                                             f"Symbol appears on Universal rows {uni.duplicates[h.symbol]}", h))
            continue
        u, how = uni.by_key.get(h.symbol), "symbol"
        if u is None:
            u, how = uni.by_code.get(h.symbol), "ticker code"
        if u is None and h.name:
            cands = uni.by_name.get(normalize_name(h.name), [])
            if len(cands) == 1:
                u, how = cands[0], "name"
            elif len(cands) > 1:
                exceptions.append(MatchException("Ambiguous", h.raw_symbol,
                                                 f"Name '{h.name}' matches Universal rows {[c.row for c in cands]}", h))
                continue
        if u is None:
            exceptions.append(MatchException("Unmatched", h.raw_symbol,
                                             f"Not found in Universal Sheet by ISIN ({h.isin or 'none in file'}), symbol, ticker code or exact name"
                                             f"{' (' + h.name + ')' if h.name else ''}. "
                                             "Add it to the Universal Sheet (or its 'Symbol' column).", h))
            continue
        matched.append(Matched(h, u, how))
    return matched, exceptions

"""One-call entry point: ticker in, comps analysis out."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from .analysis import CompsResult, run_comps
from .data import DataProvider
from .peers import PeerScore, find_peers, score_manual_peers


@dataclass
class Analysis:
    result: CompsResult
    candidates: list[PeerScore]   # every peer candidate considered, best first


def analyze(ticker: str, provider: Optional[DataProvider] = None, n_peers: int = 5,
            manual_peers: Optional[list[str]] = None) -> Analysis:
    if provider is None:
        from .data import YahooProvider
        provider = YahooProvider()
    target = provider.get_company(ticker)
    if manual_peers:
        chosen = score_manual_peers(provider, target, manual_peers)
        candidates = chosen + [c for c in find_peers(provider, target)
                               if c.ticker not in {s.ticker for s in chosen}]
    else:
        candidates = find_peers(provider, target)
        chosen = candidates[:n_peers]
    if not chosen:
        from .data import DataError
        raise DataError(f"Could not find usable peers for {target.ticker}. Try entering peers manually.")
    return Analysis(result=run_comps(target, [s.company for s in chosen]), candidates=candidates)

"""One-call entry point: ticker in, comps analysis out."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from .analysis import CompsResult, run_comps
from .data import DataProvider, resolve_ticker
from .peers import PeerScore, find_peers, score_manual_peers


@dataclass
class Analysis:
    result: CompsResult
    candidates: list[PeerScore]   # every peer candidate considered, best first


def _try_resolve(provider: DataProvider, query: str) -> Optional[str]:
    """Resolve a custom peer, skipping (rather than failing on) anything unrecognised."""
    try:
        return resolve_ticker(provider, query)
    except Exception:
        return None


def analyze(query: str, provider: Optional[DataProvider] = None, n_peers: int = 5,
            manual_peers: Optional[list[str]] = None) -> Analysis:
    if provider is None:
        from .data import YahooProvider
        provider = YahooProvider()
    target = provider.get_company(resolve_ticker(provider, query))
    if manual_peers:
        chosen = score_manual_peers(provider, target, [t for t in (_try_resolve(provider, q) for q in manual_peers) if t])
        candidates = chosen + [c for c in find_peers(provider, target)
                               if c.ticker not in {s.ticker for s in chosen}]
    else:
        candidates = find_peers(provider, target)
        chosen = candidates[:n_peers]
    if not chosen:
        from .data import DataError
        raise DataError(f"Could not find usable peers for {target.ticker}. Try entering peers manually.")
    return Analysis(result=run_comps(target, [s.company for s in chosen]), candidates=candidates)

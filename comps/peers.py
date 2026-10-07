"""Peer discovery and ranking.

Candidates come from three Yahoo sources:
  1. Top companies in the same Yahoo *industry* (closest business match)
  2. Yahoo's "people also watch" list (how the market groups the stock)
  3. Top companies in the same *sector* (only used to backfill a thin list)

Each candidate is then scored 0-100 on business and financial similarity so the
strongest 3-5 comps float to the top. The score breakdown is kept so the user can
see *why* a peer was picked and override it.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Optional

from .data import CompanyData, DataProvider, fetch_many

WEIGHTS = {
    "industry": 35.0,   # same industry (sector-only match earns partial credit)
    "market": 10.0,     # appears in Yahoo's "people also watch"
    "size": 20.0,       # market cap proximity (log scale)
    "growth": 15.0,     # revenue growth proximity
    "ebitda_margin": 10.0,
    "gross_margin": 10.0,
}


@dataclass
class PeerScore:
    company: CompanyData
    score: float
    components: dict[str, float] = field(default_factory=dict)
    reasons: list[str] = field(default_factory=list)

    @property
    def ticker(self) -> str:
        return self.company.ticker


def _closeness(a: Optional[float], b: Optional[float], window: float) -> Optional[float]:
    """1.0 when equal, falling linearly to 0 when |a - b| >= window."""
    if a is None or b is None:
        return None
    return max(0.0, 1.0 - abs(a - b) / window)


def score_peer(target: CompanyData, cand: CompanyData, market_peer: bool = False) -> PeerScore:
    comp: dict[str, float] = {}
    reasons: list[str] = []

    if target.industry and cand.industry == target.industry:
        comp["industry"] = WEIGHTS["industry"]
        reasons.append(f"Same industry ({cand.industry})")
    elif target.sector and cand.sector == target.sector:
        comp["industry"] = WEIGHTS["industry"] * 0.4
        reasons.append(f"Same sector ({cand.sector})")
    else:
        comp["industry"] = 0.0

    comp["market"] = WEIGHTS["market"] if market_peer else 0.0
    if market_peer:
        reasons.append("Grouped with target by investors (Yahoo 'people also watch')")

    if target.market_cap and cand.market_cap and target.market_cap > 0 and cand.market_cap > 0:
        ratio = cand.market_cap / target.market_cap
        size = max(0.0, 1.0 - abs(math.log10(ratio)) / 2.0)  # 0 at 100x apart
        comp["size"] = WEIGHTS["size"] * size
        if size >= 0.75:
            reasons.append(f"Similar size ({ratio:.1f}x target market cap)")
    else:
        comp["size"] = 0.0

    # Missing data earns half credit so we don't over-penalise sparse records
    for key, a, b, window, label in (
        ("growth", target.revenue_growth, cand.revenue_growth, 0.30, "revenue growth"),
        ("ebitda_margin", target.ebitda_margin, cand.ebitda_margin, 0.30, "EBITDA margin"),
        ("gross_margin", target.gross_margin, cand.gross_margin, 0.40, "gross margin"),
    ):
        c = _closeness(a, b, window)
        comp[key] = WEIGHTS[key] * (0.5 if c is None else c)
        if c is not None and c >= 0.8:
            reasons.append(f"Similar {label} ({b:.0%} vs {a:.0%})")

    return PeerScore(company=cand, score=round(sum(comp.values()), 1), components=comp, reasons=reasons)


def _is_usable(target: CompanyData, c: CompanyData) -> bool:
    if c.ticker == target.ticker:
        return False
    if (c.quote_type or "EQUITY").upper() != "EQUITY":
        return False
    if not c.market_cap or not c.price or not c.revenue or c.revenue <= 0:
        return False
    # Stick to the target's listing market: skip foreign-exchange lines like 'SAP.DE'
    # for a US target (ADRs such as BABA or TSM have no suffix and are kept).
    target_foreign = "." in target.ticker
    if not target_foreign and "." in c.ticker:
        return False
    return True


def _dedupe_share_classes(scores: list[PeerScore]) -> list[PeerScore]:
    """Keep one line per issuer (e.g. GOOGL vs GOOG, BRK-A vs BRK-B)."""
    seen: set[str] = set()
    out = []
    for s in scores:
        key = (s.company.name or s.ticker).lower()
        for suffix in (" class a", " class b", " class c", " inc.", " inc", ","):
            key = key.replace(suffix, "")
        key = key.strip()
        if key in seen:
            continue
        seen.add(key)
        out.append(s)
    return out


def find_peers(
    provider: DataProvider,
    target: CompanyData,
    max_candidates: int = 20,
    extra: Optional[list[str]] = None,
) -> list[PeerScore]:
    """Return every usable candidate, ranked best-first."""
    industry = provider.industry_peers(target.industry_key)
    market = provider.recommended_symbols(target.ticker)
    pool = list(dict.fromkeys(industry + market + (extra or [])))
    if len([p for p in pool if p.upper() != target.ticker]) < 8:
        pool += [s for s in provider.sector_peers(target.sector_key) if s not in pool]
    pool = [p.upper() for p in pool if p.upper() != target.ticker][:max_candidates]

    companies = fetch_many(provider, pool)
    market_set = {m.upper() for m in market}
    scores = [
        score_peer(target, c, market_peer=c.ticker in market_set)
        for c in companies.values()
        if _is_usable(target, c)
    ]
    scores.sort(key=lambda s: s.score, reverse=True)
    return _dedupe_share_classes(scores)


def score_manual_peers(provider: DataProvider, target: CompanyData, tickers: list[str]) -> list[PeerScore]:
    """User-specified peers: fetched and scored, but kept in the order given."""
    companies = fetch_many(provider, tickers)
    out = []
    for t in tickers:
        c = companies.get(t.strip().upper())
        if c is not None and c.ticker != target.ticker:
            s = score_peer(target, c)
            s.reasons.insert(0, "Selected manually")
            out.append(s)
    return out

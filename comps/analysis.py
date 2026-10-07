"""Relative valuation: where does the target trade vs. its peers, and what price
would peer multiples imply?"""
from __future__ import annotations

from dataclasses import dataclass, field
from statistics import mean, median
from typing import Optional

from .data import CompanyData
from .multiples import MULTIPLES, Selection, select_multiples


def percentile(values: list[float], q: float) -> Optional[float]:
    """Linear-interpolated percentile (same as Excel PERCENTILE.INC)."""
    if not values:
        return None
    v = sorted(values)
    pos = (len(v) - 1) * q
    lo, hi = int(pos), min(int(pos) + 1, len(v) - 1)
    return v[lo] + (v[hi] - v[lo]) * (pos - lo)


@dataclass
class MultipleResult:
    key: str
    label: str
    target: Optional[float]
    peers: dict[str, Optional[float]]
    low: Optional[float] = None      # 25th percentile
    median: Optional[float] = None
    mean: Optional[float] = None
    high: Optional[float] = None     # 75th percentile
    min: Optional[float] = None
    max: Optional[float] = None
    premium: Optional[float] = None  # target / peer median - 1
    rank: Optional[str] = None       # e.g. "4th of 6"
    implied_low: Optional[float] = None
    implied_mid: Optional[float] = None
    implied_high: Optional[float] = None
    is_primary: bool = False

    @property
    def peer_values(self) -> list[float]:
        return [v for v in self.peers.values() if v is not None]


@dataclass
class CompsResult:
    target: CompanyData
    peers: list[CompanyData]
    selection: Selection
    multiples: list[MultipleResult]
    blended_price: Optional[float] = None
    blended_upside: Optional[float] = None
    avg_premium: Optional[float] = None
    operating: dict[str, dict] = field(default_factory=dict)
    verdict: str = ""


def implied_price(target: CompanyData, key: str, multiple: Optional[float]) -> Optional[float]:
    m = MULTIPLES[key]
    metric = getattr(target, m.metric)
    shares = target.shares
    if multiple is None or metric is None or metric <= 0 or not shares:
        return None
    value = multiple * metric
    if m.numerator == "ev":
        value -= target.net_debt or 0.0
    price = value / shares
    return price if price > 0 else None


def _ordinal(n: int) -> str:
    suffix = "th" if 11 <= n % 100 <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def analyze_multiple(target: CompanyData, peers: list[CompanyData], key: str) -> MultipleResult:
    m = MULTIPLES[key]
    tv = m.value(target)
    pv = {p.ticker: m.value(p) for p in peers}
    vals = [v for v in pv.values() if v is not None]
    r = MultipleResult(key=key, label=m.label, target=tv, peers=pv)
    if vals:
        r.low, r.median, r.high = percentile(vals, 0.25), median(vals), percentile(vals, 0.75)
        r.mean, r.min, r.max = mean(vals), min(vals), max(vals)
    if tv is not None and r.median:
        r.premium = tv / r.median - 1
        everyone = sorted(vals + [tv], reverse=True)
        r.rank = f"{_ordinal(everyone.index(tv) + 1)} highest of {len(everyone)}"
    r.implied_low = implied_price(target, key, r.low)
    r.implied_mid = implied_price(target, key, r.median)
    r.implied_high = implied_price(target, key, r.high)
    return r


OPERATING_METRICS = [
    ("revenue_growth", "Revenue Growth (YoY)"),
    ("gross_margin", "Gross Margin"),
    ("ebitda_margin", "EBITDA Margin"),
    ("operating_margin", "Operating Margin"),
    ("net_margin", "Net Margin"),
    ("roe", "Return on Equity"),
]


def operating_comparison(target: CompanyData, peers: list[CompanyData]) -> dict[str, dict]:
    out = {}
    for attr, label in OPERATING_METRICS:
        vals = [getattr(p, attr) for p in peers if getattr(p, attr) is not None]
        out[attr] = {
            "label": label,
            "target": getattr(target, attr),
            "peer_median": median(vals) if vals else None,
        }
    return out


def _verdict(res: CompsResult) -> str:
    t = res.target
    if res.avg_premium is None:
        return f"Not enough comparable data to position {t.ticker} against its peers."
    p = res.avg_premium
    if abs(p) < 0.10:
        stance = "roughly in line with"
    elif p > 0:
        stance = f"at a {p:.0%} premium to"
    else:
        stance = f"at a {-p:.0%} discount to"
    names = ", ".join(c.ticker for c in res.peers)
    lines = [f"{t.ticker} trades {stance} the peer median on its key multiples (peers: {names})."]

    op = res.operating
    g, m = op.get("revenue_growth", {}), op.get("ebitda_margin", {})
    support = []
    if g.get("target") is not None and g.get("peer_median") is not None:
        diff = g["target"] - g["peer_median"]
        if abs(diff) >= 0.03:
            support.append(f"{'faster' if diff > 0 else 'slower'} revenue growth "
                           f"({g['target']:.0%} vs {g['peer_median']:.0%})")
    if m.get("target") is not None and m.get("peer_median") is not None:
        diff = m["target"] - m["peer_median"]
        if abs(diff) >= 0.03:
            support.append(f"{'higher' if diff > 0 else 'lower'} EBITDA margins "
                           f"({m['target']:.0%} vs {m['peer_median']:.0%})")
    if support and abs(p) >= 0.10:
        justified = all(("faster" in s or "higher" in s) == (p > 0) for s in support)
        word = "consistent with" if justified else "despite"
        lines.append(f"That {'premium' if p > 0 else 'discount'} is {word} its "
                     + " and ".join(support) + ".")
    if res.blended_price is not None and t.price:
        lines.append(f"Applying peer-median multiples implies ~{t.currency} {res.blended_price:,.2f}/share "
                     f"vs. {t.price:,.2f} today ({res.blended_upside:+.0%}).")
    return " ".join(lines)


def run_comps(target: CompanyData, peers: list[CompanyData], selection: Optional[Selection] = None) -> CompsResult:
    selection = selection or select_multiples(target, peers)
    results = [analyze_multiple(target, peers, k) for k in selection.multiples]
    for r in results:
        r.is_primary = r.key in selection.primary

    res = CompsResult(target=target, peers=peers, selection=selection, multiples=results)
    prim = [r for r in results if r.is_primary]
    prem = [r.premium for r in prim if r.premium is not None]
    res.avg_premium = median(prem) if prem else None
    implied = [r.implied_mid for r in prim if r.implied_mid is not None]
    if implied:
        res.blended_price = mean(implied)
        res.blended_upside = res.blended_price / target.price - 1 if target.price else None
    res.operating = operating_comparison(target, peers)
    res.verdict = _verdict(res)
    return res

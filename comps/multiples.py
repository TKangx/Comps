"""Trading multiple definitions and the logic that decides which ones matter."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from .data import CompanyData


@dataclass(frozen=True)
class Multiple:
    key: str
    label: str
    numerator: str      # "ev" (enterprise value) or "equity" (market cap)
    metric: str         # CompanyData attribute used as denominator
    metric_label: str
    max_meaningful: float  # above this the multiple is treated as NM (not meaningful)

    def value(self, c: CompanyData) -> Optional[float]:
        num = c.enterprise_value if self.numerator == "ev" else c.market_cap
        den = getattr(c, self.metric)
        if num is None or den is None or den <= 0 or num <= 0:
            return None
        v = num / den
        return v if v <= self.max_meaningful else None


MULTIPLES: dict[str, Multiple] = {m.key: m for m in [
    Multiple("ev_revenue", "EV / LTM Revenue", "ev", "revenue", "LTM Revenue", 60),
    Multiple("ev_ntm_revenue", "EV / NTM Revenue", "ev", "ntm_revenue", "NTM Revenue", 60),
    Multiple("ev_gross_profit", "EV / Gross Profit", "ev", "gross_profit", "LTM Gross Profit", 80),
    Multiple("ev_ebitda", "EV / EBITDA", "ev", "ebitda", "LTM EBITDA", 100),
    Multiple("ev_ebit", "EV / EBIT", "ev", "ebit", "LTM EBIT", 150),
    Multiple("pe_ltm", "P / E (LTM)", "equity", "net_income", "LTM Net Income", 150),
    Multiple("pe_fwd", "P / E (Forward)", "equity", "forward_earnings", "Forward Earnings", 150),
    Multiple("p_fcf", "P / FCF", "equity", "free_cash_flow", "LTM Free Cash Flow", 150),
    Multiple("p_ocf", "P / Operating Cash Flow", "equity", "operating_cash_flow", "LTM Operating Cash Flow", 100),
    Multiple("p_b", "P / Book", "equity", "book_value", "Book Value of Equity", 50),
]}


FINANCIAL_INDUSTRY_WORDS = ("bank", "insurance", "capital markets", "asset management",
                            "mortgage", "financial conglomerates", "shell companies")


@dataclass
class Selection:
    profile: str
    explanation: str
    multiples: list[str]          # ordered, most relevant first
    primary: list[str]            # used for the blended implied price
    reasons: dict[str, str]


def classify(c: CompanyData) -> str:
    ind = (c.industry or "").lower()
    if c.sector == "Financial Services" and any(w in ind for w in FINANCIAL_INDUSTRY_WORDS):
        return "financial"
    if c.sector == "Real Estate" and "reit" in ind:
        return "reit"
    if (c.ebitda is not None and c.ebitda <= 0) or (c.net_income is not None and c.net_income <= 0
                                                    and (c.ebitda is None or c.ebitda <= 0)):
        return "unprofitable"
    if c.revenue_growth is not None and c.revenue_growth >= 0.20:
        return "high_growth"
    if c.sector in ("Energy", "Basic Materials", "Utilities"):
        return "capital_intensive"
    return "mature"


PROFILES = {
    "financial": (
        "Financial company: balance-sheet driven, so EV-based multiples are not meaningful "
        "(debt is raw material, not financing). Valued on earnings and book value.",
        ["pe_fwd", "pe_ltm", "p_b"],
        ["pe_fwd", "pe_ltm", "p_b"],
    ),
    "reit": (
        "REIT: depreciation distorts earnings. Operating cash flow is used as a proxy for "
        "FFO; EV/EBITDA and P/Book capture asset value.",
        ["p_ocf", "ev_ebitda", "p_b", "ev_revenue"],
        ["p_ocf", "ev_ebitda"],
    ),
    "unprofitable": (
        "Not yet profitable: earnings-based multiples are negative / NM, so the comparison "
        "leans on revenue and gross profit.",
        ["ev_ntm_revenue", "ev_revenue", "ev_gross_profit"],
        ["ev_ntm_revenue", "ev_revenue", "ev_gross_profit"],
    ),
    "high_growth": (
        "High-growth (20%+ revenue growth): forward-looking multiples matter most, with "
        "revenue multiples alongside since margins are still scaling.",
        ["ev_ntm_revenue", "ev_ebitda", "pe_fwd", "ev_gross_profit", "ev_revenue", "p_fcf"],
        ["ev_ntm_revenue", "ev_ebitda", "pe_fwd"],
    ),
    "capital_intensive": (
        "Capital-intensive / commodity-exposed: EV/EBITDA neutralises differences in "
        "leverage and D&A; cash-flow multiples capture reinvestment needs.",
        ["ev_ebitda", "pe_fwd", "p_ocf", "pe_ltm", "ev_ebit", "p_b"],
        ["ev_ebitda", "pe_fwd", "p_ocf"],
    ),
    "mature": (
        "Profitable, established business: standard enterprise and equity earnings multiples "
        "apply, with P/FCF as a cash sanity check.",
        ["ev_ebitda", "pe_fwd", "ev_ebit", "pe_ltm", "p_fcf", "ev_revenue"],
        ["ev_ebitda", "pe_fwd", "ev_ebit"],
    ),
}

MULTIPLE_REASONS = {
    "ev_revenue": "Capital-structure neutral; works even when earnings are depressed.",
    "ev_ntm_revenue": "Forward revenue captures growth the market is already pricing in.",
    "ev_gross_profit": "Normalises for different revenue mixes (e.g. marketplace vs. 1P sales).",
    "ev_ebitda": "Most common cross-company multiple: ignores leverage, tax and D&A differences.",
    "ev_ebit": "Like EV/EBITDA but charges for capex intensity via D&A.",
    "pe_ltm": "What investors pay per dollar of trailing earnings to equity holders.",
    "pe_fwd": "Market's primary equity multiple; based on consensus next-year EPS.",
    "p_fcf": "Cash actually generated for shareholders after capex.",
    "p_ocf": "Cash earnings before capex; a proxy for FFO in real estate.",
    "p_b": "Key for balance-sheet businesses where assets are marked near fair value.",
}


def select_multiples(target: CompanyData, peers: list[CompanyData], min_peer_values: int = 2,
                     max_multiples: int = 5) -> Selection:
    profile = classify(target)
    explanation, ordered, primary = PROFILES[profile]

    usable = []
    for key in ordered:
        m = MULTIPLES[key]
        if m.value(target) is None:
            continue
        if sum(1 for p in peers if m.value(p) is not None) < min(min_peer_values, len(peers)):
            continue
        usable.append(key)
    usable = usable[:max_multiples]

    # Fall back to anything computable if the profile's preferred set is too thin
    if len(usable) < 2:
        for key, m in MULTIPLES.items():
            if key not in usable and m.value(target) is not None and \
                    sum(1 for p in peers if m.value(p) is not None) >= 1:
                usable.append(key)
            if len(usable) >= 3:
                break

    prim = [k for k in primary if k in usable] or usable[:3]
    return Selection(profile=profile, explanation=explanation, multiples=usable, primary=prim,
                     reasons={k: MULTIPLE_REASONS[k] for k in usable})

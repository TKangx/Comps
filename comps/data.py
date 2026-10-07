"""Market data layer.

Everything downstream works with ``CompanyData`` objects whose dollar figures are
all expressed in the *quote* currency (the currency the stock trades in), so
multiples are apples-to-apples even for ADRs that report in another currency
(e.g. BABA trades in USD but reports in CNY).
"""
from __future__ import annotations

import math
import time
from dataclasses import dataclass, field, asdict
from datetime import datetime, timezone
from typing import Iterable, Optional, Protocol


def _num(value) -> Optional[float]:
    """Coerce a Yahoo value to float, mapping junk (None, 'Infinity', NaN) to None."""
    if value is None or isinstance(value, bool):
        return None
    try:
        f = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(f) or math.isinf(f):
        return None
    return f


@dataclass
class CompanyData:
    ticker: str
    name: str = ""
    sector: str = ""
    industry: str = ""
    sector_key: str = ""
    industry_key: str = ""
    country: str = ""
    exchange: str = ""
    quote_type: str = "EQUITY"
    currency: str = "USD"
    financial_currency: str = "USD"
    summary: str = ""

    # Market data (quote currency)
    price: Optional[float] = None
    market_cap: Optional[float] = None
    total_debt: Optional[float] = None
    cash: Optional[float] = None

    # LTM financials (converted to quote currency)
    revenue: Optional[float] = None
    gross_profit: Optional[float] = None
    ebitda: Optional[float] = None
    ebit: Optional[float] = None
    net_income: Optional[float] = None
    operating_cash_flow: Optional[float] = None
    free_cash_flow: Optional[float] = None
    book_value: Optional[float] = None  # total common equity

    # Forward-looking (consensus)
    ntm_revenue: Optional[float] = None
    forward_earnings: Optional[float] = None  # aggregate next-FY net income implied by forward P/E

    # Operating profile (fractions, e.g. 0.12 = 12%)
    revenue_growth: Optional[float] = None
    earnings_growth: Optional[float] = None
    gross_margin: Optional[float] = None
    ebitda_margin: Optional[float] = None
    operating_margin: Optional[float] = None
    net_margin: Optional[float] = None
    roe: Optional[float] = None
    beta: Optional[float] = None
    dividend_yield: Optional[float] = None
    peg_ratio: Optional[float] = None

    as_of: str = field(default_factory=lambda: datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"))

    @property
    def net_debt(self) -> Optional[float]:
        if self.total_debt is None and self.cash is None:
            return None
        return (self.total_debt or 0.0) - (self.cash or 0.0)

    @property
    def enterprise_value(self) -> Optional[float]:
        """Market cap + total debt - cash. (Excludes minority interest / preferreds,
        which Yahoo does not expose consistently.)"""
        if self.market_cap is None:
            return None
        return self.market_cap + (self.net_debt or 0.0)

    @property
    def shares(self) -> Optional[float]:
        """Implied share count (market cap / price). Using the implied count keeps ADR
        ratios consistent with the quoted price."""
        if not self.market_cap or not self.price:
            return None
        return self.market_cap / self.price

    def to_dict(self) -> dict:
        d = asdict(self)
        d["enterprise_value"] = self.enterprise_value
        d["net_debt"] = self.net_debt
        d["shares"] = self.shares
        return d


class DataProvider(Protocol):
    def get_company(self, ticker: str) -> CompanyData: ...
    def industry_peers(self, industry_key: str) -> list[str]: ...
    def sector_peers(self, sector_key: str) -> list[str]: ...
    def recommended_symbols(self, ticker: str) -> list[str]: ...


class DataError(RuntimeError):
    pass


def _ntm_weight(info: dict) -> float:
    """Fraction of the *current* fiscal year still ahead of us. NTM revenue is
    blended as w * FY0 + (1 - w) * FY1."""
    fy_end = _num(info.get("nextFiscalYearEnd")) or _num(info.get("lastFiscalYearEnd"))
    if not fy_end:
        return 0.5
    now = time.time()
    year = 365.25 * 86400
    # roll forward until the FY end is in the future
    while fy_end < now:
        fy_end += year
    while fy_end - now > year:
        fy_end -= year
    return max(0.0, min(1.0, (fy_end - now) / year))


class YahooProvider:
    """Pulls data from Yahoo Finance through the ``yfinance`` package."""

    RECS_URL = "https://query2.finance.yahoo.com/v6/finance/recommendationsbysymbol/{}"

    def __init__(self):
        import yfinance as yf  # imported lazily so tests don't need network libs

        self.yf = yf
        self._fx_cache: dict[tuple[str, str], float] = {}

    # ---- FX -----------------------------------------------------------------
    def fx_rate(self, from_ccy: str, to_ccy: str) -> float:
        if not from_ccy or not to_ccy or from_ccy == to_ccy:
            return 1.0
        key = (from_ccy, to_ccy)
        if key not in self._fx_cache:
            rate = None
            try:
                rate = _num(self.yf.Ticker(f"{from_ccy}{to_ccy}=X").fast_info["last_price"])
            except Exception:
                rate = None
            if not rate:
                raise DataError(f"Could not get FX rate {from_ccy}->{to_ccy}")
            self._fx_cache[key] = rate
        return self._fx_cache[key]

    # ---- Company ------------------------------------------------------------
    def get_company(self, ticker: str) -> CompanyData:
        ticker = ticker.strip().upper()
        t = self.yf.Ticker(ticker)
        try:
            info = t.info or {}
        except Exception as exc:  # network / bad ticker
            raise DataError(f"Could not load {ticker} from Yahoo Finance: {exc}") from exc
        if not info or (info.get("marketCap") is None and info.get("regularMarketPrice") is None):
            raise DataError(f"No market data found for '{ticker}'. Check the ticker symbol.")

        currency = info.get("currency") or "USD"
        fin_ccy = info.get("financialCurrency") or currency
        # Yahoo quotes some London listings in pence
        price_scale = 0.01 if currency in ("GBp", "GBX", "ZAc", "ILA") else 1.0
        if price_scale != 1.0:
            currency = {"GBp": "GBP", "GBX": "GBP", "ZAc": "ZAR", "ILA": "ILS"}[currency]
        try:
            fx = self.fx_rate(fin_ccy, currency)
        except DataError:
            fx = None  # leave financials unconverted -> treat as missing below

        def fin(key):
            v = _num(info.get(key))
            if v is None or fx is None:
                return None
            return v * fx

        price = _num(info.get("currentPrice")) or _num(info.get("regularMarketPrice"))
        price = price * price_scale if price else None
        market_cap = _num(info.get("marketCap"))
        if market_cap and price_scale != 1.0:
            # Yahoo's marketCap for pence listings is in pence too
            market_cap *= price_scale

        revenue = fin("totalRevenue")
        op_margin = _num(info.get("operatingMargins"))
        shares_implied = market_cap / price if market_cap and price else None
        bvps = _num(info.get("bookValue"))

        forward_pe = _num(info.get("forwardPE"))
        forward_earnings = market_cap / forward_pe if market_cap and forward_pe and forward_pe > 0 else None

        c = CompanyData(
            ticker=ticker,
            name=info.get("longName") or info.get("shortName") or ticker,
            sector=info.get("sector") or "",
            industry=info.get("industry") or "",
            sector_key=info.get("sectorKey") or "",
            industry_key=info.get("industryKey") or "",
            country=info.get("country") or "",
            exchange=info.get("exchange") or "",
            quote_type=info.get("quoteType") or "EQUITY",
            currency=currency,
            financial_currency=fin_ccy,
            summary=info.get("longBusinessSummary") or "",
            price=price,
            market_cap=market_cap,
            total_debt=fin("totalDebt"),
            cash=fin("totalCash"),
            revenue=revenue,
            gross_profit=fin("grossProfits"),
            ebitda=fin("ebitda"),
            ebit=(revenue * op_margin) if revenue and op_margin is not None else None,
            net_income=fin("netIncomeToCommon"),
            operating_cash_flow=fin("operatingCashflow"),
            free_cash_flow=fin("freeCashflow"),
            # bookValue is per share in the financial currency; scale by implied shares
            book_value=(bvps * fx * shares_implied) if bvps and fx and shares_implied else None,
            forward_earnings=forward_earnings,
            revenue_growth=_num(info.get("revenueGrowth")),
            earnings_growth=_num(info.get("earningsGrowth")),
            gross_margin=_num(info.get("grossMargins")),
            ebitda_margin=_num(info.get("ebitdaMargins")),
            operating_margin=op_margin,
            net_margin=_num(info.get("profitMargins")),
            roe=_num(info.get("returnOnEquity")),
            beta=_num(info.get("beta")),
            dividend_yield=_normalize_yield(_num(info.get("dividendYield"))),
            peg_ratio=_num(info.get("trailingPegRatio")) or _num(info.get("pegRatio")),
        )
        c.ntm_revenue = self._ntm_revenue(t, info, fx)
        return c

    def _ntm_revenue(self, t, info: dict, fx: Optional[float]) -> Optional[float]:
        if fx is None:
            return None
        try:
            est = t.revenue_estimate
        except Exception:
            return None
        if est is None or getattr(est, "empty", True) or "avg" not in est.columns:
            return None
        fy0 = _num(est["avg"].get("0y")) if "0y" in est.index else None
        fy1 = _num(est["avg"].get("+1y")) if "+1y" in est.index else None
        if fy0 and fy1:
            w = _ntm_weight(info)
            return (w * fy0 + (1 - w) * fy1) * fx
        if fy1 or fy0:
            return (fy1 or fy0) * fx
        return None

    # ---- Peer candidate sources ---------------------------------------------
    def industry_peers(self, industry_key: str) -> list[str]:
        if not industry_key:
            return []
        try:
            df = self.yf.Industry(industry_key).top_companies
        except Exception:
            return []
        return [] if df is None else [str(s) for s in df.index]

    def sector_peers(self, sector_key: str) -> list[str]:
        if not sector_key:
            return []
        try:
            df = self.yf.Sector(sector_key).top_companies
        except Exception:
            return []
        return [] if df is None else [str(s) for s in df.index]

    def recommended_symbols(self, ticker: str) -> list[str]:
        """Yahoo's 'People also watch' list -- a market-perceived peer signal."""
        try:
            from yfinance.data import YfData

            data = YfData().get_raw_json(self.RECS_URL.format(ticker))
            result = data["finance"]["result"][0]["recommendedSymbols"]
            return [r["symbol"] for r in result if r.get("symbol")]
        except Exception:
            return []


def _normalize_yield(y: Optional[float]) -> Optional[float]:
    # Newer yfinance versions return dividendYield in percent (0.45 == 0.45%)
    if y is None:
        return None
    return y / 100 if y > 0.25 else y


def fetch_many(provider: DataProvider, tickers: Iterable[str], max_workers: int = 8) -> dict[str, CompanyData]:
    """Fetch several companies in parallel, silently skipping ones that fail."""
    from concurrent.futures import ThreadPoolExecutor

    tickers = list(dict.fromkeys(t.upper() for t in tickers))
    out: dict[str, CompanyData] = {}

    def _one(t):
        try:
            return t, provider.get_company(t)
        except Exception:
            return t, None

    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        for t, c in pool.map(_one, tickers):
            if c is not None:
                out[t] = c
    return out

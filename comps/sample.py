"""Offline sample data so the app and tests run without internet access.

Figures are rounded, illustrative approximations for demonstration only -- they
are NOT live market data. Use the Yahoo provider for real analysis.
"""
from __future__ import annotations

from .data import CompanyData, DataError

B = 1e9


def _c(ticker, name, sector, industry, price, mcap, debt, cash, rev, ntm, gp, ebitda, ebit, ni, fwd_pe,
       ocf, fcf, bv, growth, **kw):
    return CompanyData(
        ticker=ticker, name=name, sector=sector, industry=industry,
        sector_key=sector.lower().replace(" ", "-"), industry_key=industry.lower().replace(" ", "-"),
        price=price, market_cap=mcap * B, total_debt=debt * B, cash=cash * B, revenue=rev * B,
        ntm_revenue=ntm * B if ntm else None, gross_profit=gp * B, ebitda=ebitda * B, ebit=ebit * B,
        net_income=ni * B, forward_earnings=(mcap * B / fwd_pe) if fwd_pe else None,
        operating_cash_flow=ocf * B, free_cash_flow=fcf * B if fcf is not None else None,
        book_value=bv * B, revenue_growth=growth, gross_margin=gp / rev, ebitda_margin=ebitda / rev,
        operating_margin=ebit / rev, net_margin=ni / rev, as_of="sample data", **kw)


CR, IR = "Consumer Cyclical", "Internet Retail"
CS, IC = "Communication Services", "Internet Content & Information"
TECH = "Technology"

SAMPLE: dict[str, CompanyData] = {c.ticker: c for c in [
    _c("AMZN", "Amazon.com, Inc.", CR, IR, 225.0, 2400, 160, 95, 670, 735, 330, 145, 76, 70, 33, 125, 25, 330, 0.11),
    _c("BABA", "Alibaba Group Holding Limited", CR, IR, 135.0, 320, 35, 75, 138, 150, 55, 25, 17, 17, 14, 22, 8, 135, 0.06),
    _c("PDD", "PDD Holdings Inc.", CR, IR, 120.0, 170, 2, 50, 56, 63, 33, 16, 15, 15, 10, 16, 15, 45, 0.10),
    _c("MELI", "MercadoLibre, Inc.", CR, IR, 2400.0, 120, 8, 9, 24, 30, 11, 3.6, 3.0, 2.0, 45, 7, 6, 5, 0.37),
    _c("JD", "JD.com, Inc.", CR, IR, 33.0, 50, 15, 30, 160, 172, 25, 6.5, 5.0, 5.5, 8, 7, 4, 35, 0.11),
    _c("EBAY", "eBay Inc.", CR, IR, 90.0, 41, 8, 6, 10.5, 11, 7.4, 3.3, 2.8, 2.1, 16, 2.5, 2.1, 5, 0.04),
    _c("CPNG", "Coupang, Inc.", CR, IR, 30.0, 54, 3, 6, 33, 38, 10, 1.6, 0.6, 0.2, 60, 2.0, 0.9, 4, 0.18),
    _c("CHWY", "Chewy, Inc.", CR, IR, 38.0, 16, 0.4, 0.6, 12, 12.8, 3.5, 0.6, 0.15, 0.2, 30, 0.7, 0.5, 0.3, 0.08),
    _c("WMT", "Walmart Inc.", "Consumer Defensive", "Discount Stores", 100.0, 800, 55, 9, 690, 720, 170, 42, 30, 20, 38, 37, 14, 95, 0.05),
    _c("GOOGL", "Alphabet Inc.", CS, IC, 240.0, 2900, 25, 95, 385, 430, 230, 150, 125, 115, 24, 140, 70, 360, 0.14),
    _c("GOOG", "Alphabet Inc.", CS, IC, 241.0, 2900, 25, 95, 385, 430, 230, 150, 125, 115, 24, 140, 70, 360, 0.14),
    _c("MSFT", "Microsoft Corporation", TECH, "Software - Infrastructure", 510.0, 3800, 60, 95, 290, 330, 200, 160, 130, 105, 34, 140, 72, 360, 0.16),
    _c("META", "Meta Platforms, Inc.", CS, IC, 720.0, 1800, 50, 50, 185, 215, 150, 100, 80, 70, 26, 105, 50, 200, 0.20),
    _c("SAP.DE", "SAP SE", TECH, "Software - Application", 240.0, 280, 10, 10, 37, 40, 27, 11, 9, 6, 35, 7, 6, 45, 0.10),
]}

INDUSTRY_TOP = {
    "internet-retail": ["AMZN", "BABA", "PDD", "MELI", "JD", "CPNG", "EBAY", "CHWY"],
    "internet-content-&-information": ["GOOGL", "GOOG", "META"],
}
SECTOR_TOP = {"consumer-cyclical": ["AMZN", "BABA", "WMT", "PDD", "MELI"]}
RECOMMENDED = {"AMZN": ["GOOGL", "MSFT", "META", "WMT", "SAP.DE"]}


class SampleProvider:
    def get_company(self, ticker: str) -> CompanyData:
        t = ticker.strip().upper()
        if t not in SAMPLE:
            raise DataError(f"'{t}' is not in the offline sample set ({', '.join(SAMPLE)}).")
        return SAMPLE[t]

    def industry_peers(self, industry_key: str) -> list[str]:
        return INDUSTRY_TOP.get(industry_key, [])

    def sector_peers(self, sector_key: str) -> list[str]:
        return SECTOR_TOP.get(sector_key, [])

    def recommended_symbols(self, ticker: str) -> list[str]:
        return RECOMMENDED.get(ticker.upper(), [])

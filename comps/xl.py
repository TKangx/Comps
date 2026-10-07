"""Excel integration via xlwings (Windows).

Two ways to use it from Excel:

1. **Refresh the CompsModel workbook** – type a company in the yellow cell and click
   xlwings ▸ Run main (which calls ``CompsModel.main()`` → ``refresh()``), or wire
   ``RunPython "import comps.xl; comps.xl.refresh()"`` to a button / cell-change macro.

2. **Worksheet functions** (after xlwings ▸ Import Functions, UDF module ``comps.xl``)::

       =COMPS_TICKER("Coca-Cola")            -> KO
       =COMPS_PEERS("AMZN", 5)               -> spills 5 peer tickers down
       =COMPS_FIELD("MSFT", "ebitda_margin") -> 0.56
       =COMPS_MULTIPLE("MSFT", "EV/EBITDA")  -> 24.1
       =COMPS_PEER_MEDIAN("AMZN", "Fwd P/E") -> peer-median forward P/E
       =COMPS_IMPLIED_PRICE("AMZN", "EV/EBITDA")
"""
from __future__ import annotations

import time
from typing import Optional

import xlwings as xw

from .data import CompanyData, DataError, resolve_ticker
from .excel import MAX_PEERS, payload
from .multiples import MULTIPLES
from .pipeline import analyze

CACHE_SECONDS = 3600
_cache: dict[tuple, tuple[float, object]] = {}
_provider = None


def _get_provider():
    global _provider
    if _provider is None:
        from .data import YahooProvider
        _provider = YahooProvider()
    return _provider


def _cached(key: tuple, fn):
    hit = _cache.get(key)
    if hit and time.time() - hit[0] < CACHE_SECONDS:
        return hit[1]
    value = fn()
    _cache[key] = (time.time(), value)
    return value


# ============================================================== workbook refresh
def _read(book, name: str):
    try:
        return book.names[name].refers_to_range.value
    except Exception:
        return None


def _write(book, name: str, values):
    book.names[name].refers_to_range.value = values


def _parse_peers(raw) -> list[str]:
    if not raw:
        return []
    return [t.strip() for t in str(raw).replace(";", ",").split(",") if t.strip()]


def refresh(book: Optional["xw.Book"] = None, provider=None) -> None:
    """Read the inputs on the Dashboard, run the analysis and write the results back."""
    book = book or xw.Book.caller()
    query = str(_read(book, "CompsInput") or "").strip()
    try:
        n_peers = int(float(_read(book, "CompsNumPeers") or 5))
    except (TypeError, ValueError):
        n_peers = 5
    n_peers = max(1, min(MAX_PEERS, n_peers))
    manual = _parse_peers(_read(book, "CompsCustomPeers"))[:MAX_PEERS]

    if not query:
        _write(book, "CompsStatus", "✖ Type a company name or ticker in the yellow cell first.")
        return
    _write(book, "CompsStatus", f"⏳ Loading {query} and screening peers… (10–30 seconds)")
    try:
        app = book.app
        app.screen_updating = False
    except Exception:
        app = None
    try:
        a = analyze(query, provider=provider or _get_provider(), n_peers=n_peers, manual_peers=manual or None)
        for name, block in payload(a.result, a.candidates, query=query).items():
            _write(book, name, block)
    except DataError as exc:
        _write(book, "CompsStatus", f"✖ {exc}")
    except Exception as exc:  # network hiccups etc. – show it in the sheet instead of a VBA popup
        _write(book, "CompsStatus", f"✖ Unexpected error: {exc}")
    finally:
        if app is not None:
            app.screen_updating = True


# ========================================================== worksheet functions
_ALIASES = {
    "p/e": "pe_ltm", "pe": "pe_ltm", "ltm p/e": "pe_ltm", "trailing p/e": "pe_ltm",
    "fwd p/e": "pe_fwd", "forward p/e": "pe_fwd", "fwd pe": "pe_fwd", "ntm p/e": "pe_fwd",
    "ev/revenue": "ev_revenue", "ev/sales": "ev_revenue", "ev/ltm revenue": "ev_revenue",
    "ev/ntm revenue": "ev_ntm_revenue", "ev/fwd revenue": "ev_ntm_revenue",
    "ev/gp": "ev_gross_profit", "p/b": "p_b", "p/bv": "p_b", "p/fcf": "p_fcf", "p/cf": "p_ocf",
}


def multiple_key(text: str) -> str:
    """'EV / EBITDA', 'ev/ebitda', 'ev_ebitda', 'Fwd P/E' -> internal key."""
    raw = str(text).strip()
    if raw in MULTIPLES:
        return raw
    norm = raw.lower().replace(" ", "").replace("(", "").replace(")", "")
    for key, m in MULTIPLES.items():
        if m.label.lower().replace(" ", "").replace("(", "").replace(")", "") == norm:
            return key
    for alias, key in _ALIASES.items():
        if alias.replace(" ", "") == norm:
            return key
    raise ValueError(f"Unknown multiple '{text}'. Try one of: " + ", ".join(m.label for m in MULTIPLES.values()))


def _company(query: str) -> CompanyData:
    p = _get_provider()
    ticker = _cached(("ticker", query.strip().lower()), lambda: resolve_ticker(p, query))
    return _cached(("company", ticker), lambda: p.get_company(ticker))


def _analysis(query: str, n: int):
    return _cached(("analysis", query.strip().lower(), n), lambda: analyze(query, _get_provider(), n_peers=n))


def _safe(fn):
    try:
        return fn()
    except Exception as exc:
        return f"#ERR: {exc}"


@xw.func
def COMPS_TICKER(query):
    """Resolve a company name or ticker to its ticker symbol."""
    return _safe(lambda: _company(query).ticker)


@xw.func
def COMPS_NAME(query):
    """Company name for a ticker or search term."""
    return _safe(lambda: _company(query).name)


@xw.func
@xw.arg("n", numbers=int)
def COMPS_PEERS(query, n=5):
    """Top-n auto-selected peer tickers (spills down)."""
    def run():
        return [[p.ticker] for p in _analysis(query, n).result.peers]
    return _safe(run)


@xw.func
def COMPS_FIELD(query, field):
    """Any data field, e.g. price, market_cap, enterprise_value, revenue, ebitda, revenue_growth."""
    def run():
        c = _company(query)
        key = str(field).strip().lower().replace(" ", "_")
        key = {"ev": "enterprise_value", "mkt_cap": "market_cap", "marketcap": "market_cap"}.get(key, key)
        d = c.to_dict()
        if key not in d:
            raise ValueError(f"Unknown field '{field}'")
        return d[key]
    return _safe(run)


@xw.func
def COMPS_MULTIPLE(query, multiple):
    """A company's trading multiple, e.g. =COMPS_MULTIPLE("MSFT","EV/EBITDA"). Returns "NM" if not meaningful."""
    def run():
        v = MULTIPLES[multiple_key(multiple)].value(_company(query))
        return "NM" if v is None else v
    return _safe(run)


def _multiple_result(query, multiple, n):
    key = multiple_key(multiple)
    from .analysis import analyze_multiple
    a = _analysis(query, n)
    return analyze_multiple(a.result.target, a.result.peers, key)


@xw.func
@xw.arg("n", numbers=int)
def COMPS_PEER_MEDIAN(query, multiple, n=5):
    """Median of a multiple across the auto-selected peers."""
    def run():
        r = _multiple_result(query, multiple, n)
        return "NM" if r.median is None else r.median
    return _safe(run)


@xw.func
@xw.arg("n", numbers=int)
def COMPS_IMPLIED_PRICE(query, multiple, n=5):
    """Share price implied by applying the peer-median multiple to the company."""
    def run():
        r = _multiple_result(query, multiple, n)
        return "NM" if r.implied_mid is None else r.implied_mid
    return _safe(run)

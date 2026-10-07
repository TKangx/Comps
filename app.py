"""Comps Analyzer – Streamlit web app.

Run locally:   streamlit run app.py
"""
from __future__ import annotations

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from comps.analysis import run_comps
from comps.data import DataError, resolve_ticker
from comps.excel import to_bytes
from comps.multiples import MULTIPLE_REASONS, MULTIPLES, select_multiples
from comps.peers import find_peers, score_manual_peers

st.set_page_config(page_title="Comps Analyzer", page_icon="📊", layout="wide")

PREMIUM, DISCOUNT = "#e34948", "#2a78d6"   # validated diverging pair (light & dark legible)
RANGE, MEDIAN_MARK = "#86b6ef", "#1c5cab"


# ----------------------------------------------------------------- data access
def _provider(sample: bool):
    if sample:
        from comps.sample import SampleProvider
        return SampleProvider()
    from comps.data import YahooProvider
    return YahooProvider()


@st.cache_data(ttl=3600, show_spinner=False)
def load_universe(ticker: str, sample: bool):
    p = _provider(sample)
    target = p.get_company(resolve_ticker(p, ticker))
    return target, find_peers(p, target)


@st.cache_data(ttl=3600, show_spinner=False)
def load_manual(ticker: str, extra: tuple[str, ...], sample: bool):
    p = _provider(sample)
    resolved = []
    for q in extra:
        try:
            resolved.append(resolve_ticker(p, q))
        except DataError:
            pass
    return score_manual_peers(p, p.get_company(ticker), resolved)


def fmt_x(v):
    return "NM" if v is None or pd.isna(v) else f"{v:.1f}x"


def fmt_pp(v):
    return "–" if v is None else f"{v * 100:+.1f} pp"


def fmt_pct(v, signed=False):
    if v is None:
        return "–"
    return f"{v:+.0%}" if signed else f"{v:.1%}"


def fmt_big(v, ccy=""):
    if v is None:
        return "–"
    for div, suf in ((1e12, "T"), (1e9, "B"), (1e6, "M")):
        if abs(v) >= div:
            return f"{ccy}{v / div:,.1f}{suf}"
    return f"{ccy}{v:,.0f}"


# --------------------------------------------------------------------- sidebar
with st.sidebar:
    st.title("📊 Comps Analyzer")
    with st.form("controls"):
        ticker = st.text_input("Company name or ticker", value=st.session_state.get("ticker", "AMZN")).strip()
        n_peers = st.slider("Number of peers", 3, 8, 5)
        sample = st.toggle("Use offline sample data", value=False,
                           help="Bundled illustrative data for AMZN & peers – for demos without internet.")
        submitted = st.form_submit_button("Run analysis", type="primary", use_container_width=True)
    st.caption("Data: Yahoo Finance via yfinance (cached 1 hour). Always sanity-check key "
               "figures against filings before presenting.")

if submitted:
    st.session_state["ticker"] = ticker
    st.session_state.pop("peer_pick", None)
if "ticker" not in st.session_state:
    st.info("Enter a company name or ticker in the sidebar and press **Run analysis**.")
    st.stop()
ticker = st.session_state["ticker"]

try:
    with st.spinner(f"Pulling {ticker} and screening peer candidates…"):
        target, candidates = load_universe(ticker, sample)
except DataError as exc:
    st.error(str(exc))
    st.stop()
except Exception as exc:  # network etc.
    st.error(f"Something went wrong loading data: {exc}")
    st.stop()

# ------------------------------------------------------------------ header
st.header(f"{target.name} ({target.ticker})")
st.caption(f"{target.sector} · {target.industry} · {target.country or ''}  ·  as of {target.as_of}")
h = st.columns(5)
h[0].metric("Share price", f"{target.currency} {target.price:,.2f}" if target.price else "–")
h[1].metric("Market cap", fmt_big(target.market_cap))
h[2].metric("Enterprise value", fmt_big(target.enterprise_value))
h[3].metric("Revenue growth", fmt_pct(target.revenue_growth))
h[4].metric("EBITDA margin", fmt_pct(target.ebitda_margin))

# --------------------------------------------------------------- peer choice
st.subheader("1 · Peer group")
cand_by_ticker = {c.ticker: c for c in candidates}
default_pick = [c.ticker for c in candidates[:n_peers]]
c1, c2 = st.columns([3, 2])
with c1:
    pick = st.multiselect("Peers (auto-ranked by similarity – add or remove freely)",
                          options=list(cand_by_ticker), default=default_pick, key="peer_pick",
                          format_func=lambda t: f"{t} ({cand_by_ticker[t].score:.0f})")
with c2:
    extra_raw = st.text_input("Add other tickers (comma-separated)", placeholder="e.g. MSFT, GOOGL, WMT")
extra = tuple(t.strip().upper() for t in extra_raw.split(",") if t.strip() and t.strip().upper() not in pick)
manual = load_manual(target.ticker, extra, sample) if extra else []
if extra and len(manual) < len(extra):
    st.warning(f"{len(extra) - len(manual)} of the added names couldn't be found – check the spelling or "
               "use the ticker.")

chosen = [cand_by_ticker[t] for t in pick] + manual
if not chosen:
    st.warning("Select at least one peer.")
    st.stop()

with st.expander(f"Why these peers? (screened {len(candidates)} candidates)", expanded=False):
    rows = []
    chosen_set = {c.ticker for c in chosen}
    for s in sorted({c.ticker: c for c in candidates + manual}.values(), key=lambda s: -s.score):
        rows.append({"Selected": "✅" if s.ticker in chosen_set else "", "Ticker": s.ticker,
                     "Company": s.company.name, "Industry": s.company.industry,
                     "Mkt cap": fmt_big(s.company.market_cap), "Score": s.score,
                     "Rationale": "; ".join(s.reasons)})
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True,
                 column_config={"Score": st.column_config.ProgressColumn(min_value=0, max_value=100,
                                                                        format="%.0f")})
    st.caption("Score (0–100): industry match 35 · investor grouping ('people also watch') 10 · "
               "market-cap proximity 20 · revenue growth 15 · EBITDA margin 10 · gross margin 10.")

peers = [s.company for s in chosen]
auto_sel = select_multiples(target, peers)

# ---------------------------------------------------------- multiple choice
st.subheader("2 · Trading multiples")
st.markdown(f"**Profile: {auto_sel.profile.replace('_', ' ').title()}** — {auto_sel.explanation}")
all_keys = list(MULTIPLES)
mult_pick = st.multiselect("Multiples shown (auto-selected; edit to taste)", options=all_keys,
                           default=auto_sel.multiples, format_func=lambda k: MULTIPLES[k].label,
                           key=f"mults_{target.ticker}")
if not mult_pick:
    st.warning("Select at least one multiple.")
    st.stop()
auto_sel.multiples = mult_pick
auto_sel.primary = [k for k in auto_sel.primary if k in mult_pick] or mult_pick[:3]
auto_sel.reasons = {k: MULTIPLE_REASONS[k] for k in mult_pick}
res = run_comps(target, peers, auto_sel)

# Verdict callout
st.info(res.verdict, icon="🧭")

# Comps table
table = []
for c in [target] + peers:
    row = {"Ticker": c.ticker, "Company": c.name[:30], "Mkt cap": fmt_big(c.market_cap)}
    for k in mult_pick:
        row[MULTIPLES[k].label] = MULTIPLES[k].value(c)
    table.append(row)
for stat, attr in (("Peer 25th pct", "low"), ("Peer median", "median"), ("Peer mean", "mean"),
                   ("Peer 75th pct", "high")):
    row = {"Ticker": "", "Company": stat, "Mkt cap": ""}
    for m in res.multiples:
        row[m.label] = getattr(m, attr)
    table.append(row)
df = pd.DataFrame(table)
mult_cols = [MULTIPLES[k].label for k in mult_pick]


def _highlight(row):
    if row.name == 0:
        return ["background-color: rgba(237,161,0,0.18); font-weight: 600"] * len(row)
    if row.name > len(peers):
        return ["background-color: rgba(128,128,128,0.12); font-style: italic"] * len(row)
    return [""] * len(row)


for c in mult_cols:
    df[c] = df[c].map(fmt_x)
st.dataframe(df.style.apply(_highlight, axis=1), hide_index=True, use_container_width=True,
             height=36 * (len(df) + 1) + 4)
st.caption("NM = not meaningful (negative denominator or extreme outlier); NM values are excluded from "
           "peer statistics. Highlighted row = target.")

# ---------------------------------------------------------------- charts
st.subheader("3 · Where it stands")
g1, g2 = st.columns(2)

with g1:
    ms = [m for m in res.multiples if m.premium is not None]
    if ms:
        fig = go.Figure(go.Bar(
            y=[m.label for m in ms], x=[m.premium * 100 for m in ms], orientation="h",
            marker=dict(color=[PREMIUM if m.premium > 0 else DISCOUNT for m in ms],
                        cornerradius=4),
            text=[f"{m.premium:+.0%}" for m in ms], textposition="outside", cliponaxis=False,
            customdata=[[fmt_x(m.target), fmt_x(m.median), m.rank] for m in ms],
            hovertemplate="<b>%{y}</b><br>Target %{customdata[0]} vs peer median %{customdata[1]}"
                          "<br>Premium %{x:+.0f}%<br>%{customdata[2]}<extra></extra>",
        ))
        fig.add_vline(x=0, line_width=1, line_color="rgba(128,128,128,0.6)")
        fig.update_layout(title=dict(text=f"{target.ticker} premium (+) / discount (−) vs. peer median",
                                     font_size=14),
                          xaxis=dict(ticksuffix="%", zeroline=False, gridcolor="rgba(128,128,128,0.15)",
                                     range=[min(0, min(m.premium for m in ms) * 100) * 1.3 - 15,
                                            max(0, max(m.premium for m in ms) * 100) * 1.3 + 15]),
                          yaxis=dict(autorange="reversed"), height=80 + 50 * len(ms),
                          margin=dict(l=10, r=40, t=50, b=30), showlegend=False, bargap=0.45)
        st.plotly_chart(fig, use_container_width=True)
        st.caption("Red = premium (more expensive than peers) · Blue = discount.")

with g2:
    ff = [m for m in res.multiples if m.implied_low is not None and m.implied_high is not None]
    if ff and target.price:
        fig = go.Figure()
        fig.add_trace(go.Bar(
            y=[m.label for m in ff], x=[m.implied_high - m.implied_low for m in ff],
            base=[m.implied_low for m in ff], orientation="h", name="Peer 25th–75th pct",
            marker=dict(color=RANGE, cornerradius=4),
            customdata=[[m.implied_low, m.implied_mid or 0, m.implied_high] for m in ff],
            hovertemplate="<b>%{y}</b><br>25th: %{customdata[0]:,.2f}<br>Median: %{customdata[1]:,.2f}"
                          "<br>75th: %{customdata[2]:,.2f}<extra></extra>"))
        fig.add_trace(go.Scatter(
            y=[m.label for m in ff], x=[m.implied_mid for m in ff], mode="markers", name="Peer median",
            marker=dict(symbol="line-ns", size=18, line=dict(width=3, color=MEDIAN_MARK)),
            hoverinfo="skip"))
        blended_left = bool(res.blended_price and res.blended_price < target.price)
        lines = [(target.price, f"Current {target.price:,.2f}", "dash", None, not blended_left)]
        if res.blended_price:
            lines.append((res.blended_price, f"Blended {res.blended_price:,.2f}", "dot", MEDIAN_MARK,
                          blended_left))
        for x, label, dash, color, left in lines:
            fig.add_vline(x=x, line_dash=dash, line_width=2 if dash == "dash" else 1.5,
                          line_color=color or "rgba(128,128,128,0.9)")
            fig.add_annotation(x=x, y=1.0, yref="paper", yanchor="bottom", showarrow=False,
                               xanchor="right" if left else "left", text=label, font_size=11)
        fig.update_layout(title=dict(text="Implied share price (football field)", font_size=14),
                          yaxis=dict(autorange="reversed"), xaxis=dict(gridcolor="rgba(128,128,128,0.15)"),
                          height=120 + 50 * len(ff), margin=dict(l=10, r=20, t=80, b=40),
                          legend=dict(orientation="h", y=-0.15), bargap=0.45)
        st.plotly_chart(fig, use_container_width=True)

# Implied valuation table
iv = pd.DataFrame([{
    "Multiple": m.label + (" ★" if m.is_primary else ""),
    f"{target.ticker}": fmt_x(m.target), "Peer median": fmt_x(m.median),
    "Premium": fmt_pct(m.premium, signed=True), "Rank": m.rank or "–",
    "Implied @ 25th": m.implied_low, "Implied @ median": m.implied_mid, "Implied @ 75th": m.implied_high,
    "Why this multiple": res.selection.reasons.get(m.key, ""),
} for m in res.multiples])
st.dataframe(iv.style.format({c: (lambda v: "–" if pd.isna(v) else f"{v:,.2f}")
                              for c in ("Implied @ 25th", "Implied @ median", "Implied @ 75th")}),
             hide_index=True, use_container_width=True)
if res.blended_price:
    st.markdown(f"**Blended implied price (★ multiples, peer median): {target.currency} "
                f"{res.blended_price:,.2f}** → {fmt_pct(res.blended_upside, signed=True)} vs. current.")

# ------------------------------------------------------- operating context
st.subheader("4 · Is the premium / discount justified?")
st.caption("Multiples only make sense next to fundamentals: faster growth and higher margins usually "
           "earn a premium.")
op = pd.DataFrame([{"Metric": v["label"], target.ticker: fmt_pct(v["target"]),
                    "Peer median": fmt_pct(v["peer_median"]),
                    "Difference": fmt_pp(None if v["target"] is None or v["peer_median"] is None
                                         else v["target"] - v["peer_median"])}
                   for v in res.operating.values()])
st.dataframe(op, hide_index=True, use_container_width=True)

# --------------------------------------------------------------- export
st.subheader("5 · Export")
st.download_button("⬇️ Download Excel comps model", data=to_bytes(res, candidates + manual, query=ticker),
                   file_name=f"{target.ticker}_comps.xlsx", type="primary",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.caption("The workbook is formula-driven: overwrite any blue input or the Y/N blend flags and every "
           "multiple, statistic and implied price recalculates.")

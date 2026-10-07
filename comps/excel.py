"""Excel export: a live, formula-driven comps workbook.

Raw data is written as blue hard-coded inputs; every multiple, statistic and
implied price is an Excel formula, so the user can override any input (or
add/remove a peer row) and the whole workbook recalculates.
"""
from __future__ import annotations

from io import BytesIO
from typing import Optional

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .analysis import CompsResult
from .multiples import MULTIPLES
from .peers import PeerScore

FONT = "Arial"
BLUE, BLACK, GREEN, WHITE = "0000FF", "000000", "008000", "FFFFFF"
NAVY = "1F3864"
FILL_HEADER = PatternFill("solid", fgColor=NAVY)
FILL_TARGET = PatternFill("solid", fgColor="FFF2CC")
FILL_STATS = PatternFill("solid", fgColor="F2F2F2")
FILL_SELECTED = PatternFill("solid", fgColor="DDEBF7")
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
THIN = Side(style="thin", color="BFBFBF")

FMT_MM = '#,##0;(#,##0);"-"'
FMT_PRICE = '#,##0.00;(#,##0.00);"-"'
FMT_PCT = '0.0%;(0.0%);"-"'
FMT_X = '0.0"x";(0.0"x");"-"'

# (attribute, header, format, scale)  -- scale 1e-6 converts to millions
INPUT_COLUMNS = [
    ("ticker", "Ticker", None, None),
    ("name", "Company", None, None),
    ("price", "Share Price", FMT_PRICE, 1),
    ("market_cap", "Market Cap ($mm)", FMT_MM, 1e-6),
    ("total_debt", "Total Debt ($mm)", FMT_MM, 1e-6),
    ("cash", "Cash ($mm)", FMT_MM, 1e-6),
    ("__ev__", "Enterprise Value ($mm)", FMT_MM, None),
    ("revenue", "LTM Revenue ($mm)", FMT_MM, 1e-6),
    ("ntm_revenue", "NTM Revenue ($mm)", FMT_MM, 1e-6),
    ("gross_profit", "LTM Gross Profit ($mm)", FMT_MM, 1e-6),
    ("ebitda", "LTM EBITDA ($mm)", FMT_MM, 1e-6),
    ("ebit", "LTM EBIT ($mm)", FMT_MM, 1e-6),
    ("net_income", "LTM Net Income ($mm)", FMT_MM, 1e-6),
    ("forward_earnings", "Forward Earnings ($mm)", FMT_MM, 1e-6),
    ("operating_cash_flow", "LTM Op. Cash Flow ($mm)", FMT_MM, 1e-6),
    ("free_cash_flow", "LTM FCF ($mm)", FMT_MM, 1e-6),
    ("book_value", "Book Value ($mm)", FMT_MM, 1e-6),
    ("revenue_growth", "Revenue Growth", FMT_PCT, 1),
    ("gross_margin", "Gross Margin", FMT_PCT, 1),
    ("ebitda_margin", "EBITDA Margin", FMT_PCT, 1),
    ("net_margin", "Net Margin", FMT_PCT, 1),
]

STAT_ROWS = [
    ("Maximum", "MAX({r})"),
    ("75th Percentile", "PERCENTILE({r},0.75)"),
    ("Mean", "AVERAGE({r})"),
    ("Median", "MEDIAN({r})"),
    ("25th Percentile", "PERCENTILE({r},0.25)"),
    ("Minimum", "MIN({r})"),
]


def _style(cell, *, bold=False, color=BLACK, fmt=None, fill=None, align=None, size=10, italic=False):
    cell.font = Font(name=FONT, bold=bold, color=color, size=size, italic=italic)
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    if align:
        if align == "wrap":
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        else:
            cell.alignment = Alignment(horizontal=align, vertical="center")
    return cell


def _put(ws, row, col, value, **kw):
    c = ws.cell(row=row, column=col, value=value)
    return _style(c, **kw)


def _header_row(ws, row, headers, start_col=1):
    for i, h in enumerate(headers):
        c = _put(ws, row, start_col + i, h, bold=True, color=WHITE, fill=FILL_HEADER)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def _title(ws, title, subtitle):
    _put(ws, 1, 1, title, bold=True, size=14, color=NAVY)
    _put(ws, 2, 1, subtitle, italic=True, size=9)


def build_workbook(res: CompsResult, candidates: Optional[list[PeerScore]] = None) -> Workbook:
    wb = Workbook()
    t = res.target
    ccy = t.currency
    selected_keys = res.selection.multiples

    # ------------------------------------------------------------------ Comps
    ws = wb.active
    ws.title = "Comps"
    _title(ws, f"{t.name} ({t.ticker}) – Comparable Companies Analysis",
           f"{ccy} in millions, except per-share data. Source: Yahoo Finance, as of {t.as_of}. "
           "Blue = input, black = formula. Highlighted columns are the multiples selected for this profile.")

    hdr_row, cap_row = 5, 4
    first = hdr_row + 1
    companies = [t] + list(res.peers)
    last = first + len(companies) - 1
    peer_first, peer_last = first + 1, last

    col = {}
    for i, (attr, header, fmt, scale) in enumerate(INPUT_COLUMNS, start=1):
        col[attr] = i
    mult_start = len(INPUT_COLUMNS) + 2  # one spacer column
    for j, key in enumerate(MULTIPLES):
        col[key] = mult_start + j
    L = {k: get_column_letter(v) for k, v in col.items()}

    _header_row(ws, hdr_row, [h for _, h, _, _ in INPUT_COLUMNS])
    _header_row(ws, hdr_row, [MULTIPLES[k].label for k in MULTIPLES], start_col=mult_start)
    _put(ws, cap_row, mult_start - 1, "NM if above →", italic=True, size=8, align="right")
    for key, m in MULTIPLES.items():
        c = _put(ws, cap_row, col[key], m.max_meaningful, color=BLUE, fmt=FMT_X, size=8, align="center")
        c.comment = Comment("Multiples above this cap (or with a negative denominator) are shown "
                            "as NM and excluded from peer statistics. Edit to change.", "Comps")
        if key in selected_keys:
            ws.cell(row=hdr_row, column=col[key]).fill = PatternFill("solid", fgColor="2E75B6")

    for r, c in enumerate(companies, start=first):
        is_target = r == first
        fill = FILL_TARGET if is_target else None
        for attr, header, fmt, scale in INPUT_COLUMNS:
            if attr == "__ev__":
                v = f"={L['market_cap']}{r}+N({L['total_debt']}{r})-N({L['cash']}{r})"
                _put(ws, r, col[attr], v, fmt=fmt, fill=fill, bold=is_target)
                continue
            raw = getattr(c, attr)
            if scale and raw is not None:
                raw = raw * scale
            _put(ws, r, col[attr], raw, color=BLUE, fmt=fmt, fill=fill, bold=is_target)
        for key, m in MULTIPLES.items():
            num = f"{L['__ev__']}{r}" if m.numerator == "ev" else f"{L['market_cap']}{r}"
            den = f"{L[m.metric]}{r}"
            cap = f"{L[key]}${cap_row}"
            f = (f'=IF(AND(ISNUMBER({den}),ISNUMBER({num})),IF(AND({den}>0,{num}>0),'
                 f'IF({num}/{den}<={cap},{num}/{den},"NM"),"NM"),"NM")')
            cfill = fill or (FILL_SELECTED if key in selected_keys else None)
            _put(ws, r, col[key], f, fmt=FMT_X, fill=cfill, bold=is_target, align="right")

    # Peer statistics (exclude the target row)
    stats_first = last + 2
    _put(ws, stats_first - 1, 1, "Peer statistics (excluding target)", bold=True, italic=True, size=9)
    stat_row_of = {}
    stat_cols = [a for a, _, fmt, _ in INPUT_COLUMNS if fmt in (FMT_PCT,)] + list(MULTIPLES)
    for i, (label, tmpl) in enumerate(STAT_ROWS):
        r = stats_first + i
        stat_row_of[label] = r
        _put(ws, r, 2, label, bold=True, fill=FILL_STATS)
        for attr in stat_cols:
            rng = f"{L[attr]}{peer_first}:{L[attr]}{peer_last}"
            fmt = FMT_PCT if attr in L and attr not in MULTIPLES else FMT_X
            _put(ws, r, col[attr], f'=IF(COUNT({rng})=0,"NM",{tmpl.format(r=rng)})',
                 fmt=fmt, fill=FILL_STATS, align="right")
    prem_row = stats_first + len(STAT_ROWS) + 1
    _put(ws, prem_row, 2, f"{t.ticker} vs. Peer Median", bold=True, fill=FILL_TARGET)
    for key in MULTIPLES:
        tc, mc = f"{L[key]}{first}", f"{L[key]}{stat_row_of['Median']}"
        _put(ws, prem_row, col[key], f'=IF(AND(ISNUMBER({tc}),ISNUMBER({mc})),{tc}/{mc}-1,"NM")',
             fmt='+0.0%;-0.0%;0.0%', fill=FILL_TARGET, bold=True, align="right")
    for attr, _, fmt, _ in INPUT_COLUMNS:
        if fmt == FMT_PCT:
            tc, mc = f"{L[attr]}{first}", f"{L[attr]}{stat_row_of['Median']}"
            _put(ws, prem_row, col[attr], f'=IF(AND(ISNUMBER({tc}),ISNUMBER({mc})),{tc}-{mc},"NM")',
                 fmt='+0.0%;-0.0%;0.0%', fill=FILL_TARGET, bold=True, align="right")
    _put(ws, prem_row + 1, 2, "Margin/growth columns show the difference in percentage points; "
         "multiple columns show the premium (+) or discount (−) vs. the peer median.", italic=True, size=8)

    ws.column_dimensions["A"].width = 9
    ws.column_dimensions["B"].width = 30
    for i in range(3, mult_start + len(MULTIPLES)):
        ws.column_dimensions[get_column_letter(i)].width = 13
    ws.column_dimensions[get_column_letter(mult_start - 1)].width = 3
    ws.freeze_panes = ws.cell(row=first, column=3)

    # -------------------------------------------------------------- Valuation
    wv = wb.create_sheet("Valuation")
    _title(wv, f"{t.ticker} – Implied Valuation from Peer Multiples",
           "Implied price = (peer multiple × target metric − net debt for EV multiples) ÷ shares. "
           "Edit the 'Use in Blend' flags (Y/N) to change which multiples drive the blended value.")
    _put(wv, 4, 1, "Current Share Price", bold=True)
    _put(wv, 4, 3, f"=Comps!{L['price']}{first}", color=GREEN, fmt=FMT_PRICE)
    _put(wv, 5, 1, "Diluted Shares (mm, implied)", bold=True)
    _put(wv, 5, 3, f"=Comps!{L['market_cap']}{first}/Comps!{L['price']}{first}", fmt=FMT_MM)
    _put(wv, 6, 1, f"Net Debt ({ccy}mm)", bold=True)
    _put(wv, 6, 3, f"=N(Comps!{L['total_debt']}{first})-N(Comps!{L['cash']}{first})", fmt=FMT_MM)
    price_ref, shares_ref, nd_ref = "$C$4", "$C$5", "$C$6"

    vh = 8
    _header_row(wv, vh, ["Multiple", "Target Metric", f"Metric ({ccy}mm)", f"{t.ticker} Multiple",
                         "Peer 25th", "Peer Median", "Peer 75th", "Implied Price @ 25th",
                         "Implied Price @ Median", "Implied Price @ 75th", "Upside @ Median",
                         "Use in Blend (Y/N)"])
    vrow = {}
    for i, key in enumerate(selected_keys):
        r = vh + 1 + i
        vrow[key] = r
        m = MULTIPLES[key]
        _put(wv, r, 1, m.label, bold=True)
        _put(wv, r, 2, m.metric_label)
        _put(wv, r, 3, f"=Comps!{L[m.metric]}{first}", color=GREEN, fmt=FMT_MM)
        _put(wv, r, 4, f"=Comps!{L[key]}{first}", color=GREEN, fmt=FMT_X, align="right")
        for c_off, stat in ((5, "25th Percentile"), (6, "Median"), (7, "75th Percentile")):
            _put(wv, r, c_off, f"=Comps!{L[key]}{stat_row_of[stat]}", color=GREEN, fmt=FMT_X, align="right")
        for c_off, src in ((8, "E"), (9, "F"), (10, "G")):
            mult, metric = f"{src}{r}", f"C{r}"
            val = f"{mult}*{metric}" + (f"-{nd_ref}" if m.numerator == "ev" else "")
            _put(wv, r, c_off, f'=IF(AND(ISNUMBER({mult}),ISNUMBER({metric})),IF({metric}>0,'
                 f'MAX(0,({val})/{shares_ref}),"NM"),"NM")', fmt=FMT_PRICE, align="right")
        _put(wv, r, 11, f'=IF(ISNUMBER(I{r}),I{r}/{price_ref}-1,"NM")', fmt='+0.0%;-0.0%;0.0%', align="right")
        c = _put(wv, r, 12, "Y" if key in res.selection.primary else "N", color=BLUE, fill=FILL_INPUT,
                 align="center")
    vlast = vh + len(selected_keys)
    br = vlast + 2
    _put(wv, br, 1, "Blended Implied Price (avg. of 'Y' rows)", bold=True, fill=FILL_TARGET)
    for c_off, src in ((8, "H"), (9, "I"), (10, "J")):
        rng = f"{src}{vh + 1}:{src}{vlast}"
        _put(wv, br, c_off, f'=IFERROR(AVERAGEIFS({rng},$L${vh + 1}:$L${vlast},"Y"),"NM")',
             fmt=FMT_PRICE, bold=True, fill=FILL_TARGET, align="right")
    _put(wv, br, 11, f'=IF(ISNUMBER(I{br}),I{br}/{price_ref}-1,"NM")', fmt='+0.0%;-0.0%;0.0%',
         bold=True, fill=FILL_TARGET, align="right")
    _put(wv, br + 1, 1, "Premium / (discount) to peers on blended basis", bold=True)
    _put(wv, br + 1, 11, f'=IF(ISNUMBER(I{br}),{price_ref}/I{br}-1,"NM")', fmt='+0.0%;-0.0%;0.0%',
         bold=True, align="right")
    wv.conditional_formatting.add(f"K{vh + 1}:K{br}", FormulaRule(
        formula=[f'AND(ISNUMBER(K{vh + 1}),K{vh + 1}<0)'], font=Font(name=FONT, color="C00000")))
    wv.column_dimensions["A"].width = 38
    wv.column_dimensions["B"].width = 22
    for c in "CDEFGHIJKL":
        wv.column_dimensions[c].width = 14

    # ---------------------------------------------------------------- Summary
    wsum = wb.create_sheet("Summary", 0)
    _title(wsum, f"{t.name} ({t.ticker}) – Relative Valuation Summary",
           f"{t.sector} / {t.industry} · {ccy} · Data: Yahoo Finance, {t.as_of}")
    _put(wsum, 4, 1, "Verdict", bold=True, color=NAVY, size=11)
    c = _put(wsum, 5, 1, res.verdict, align="wrap")
    wsum.merge_cells("A5:H7")
    wsum.row_dimensions[5].height = 30
    _put(wsum, 8, 1, "Verdict text reflects the data at export time; the tables below are live formulas.",
         italic=True, size=8)

    _put(wsum, 10, 1, "Valuation profile", bold=True, color=NAVY, size=11)
    _put(wsum, 11, 1, res.selection.profile.replace("_", " ").title(), bold=True)
    _put(wsum, 12, 1, res.selection.explanation, align="wrap")
    wsum.merge_cells("A12:H13")
    wsum.row_dimensions[12].height = 30

    sh = 15
    _header_row(wsum, sh, ["Multiple", f"{t.ticker}", "Peer Median", "Premium / (Discount)",
                           "Implied Price @ Median", "Upside", "Why this multiple"])
    for i, key in enumerate(selected_keys):
        r = sh + 1 + i
        vr = vrow[key]
        _put(wsum, r, 1, MULTIPLES[key].label + (" ★" if key in res.selection.primary else ""), bold=True)
        _put(wsum, r, 2, f"=Valuation!D{vr}", color=GREEN, fmt=FMT_X, align="right")
        _put(wsum, r, 3, f"=Valuation!F{vr}", color=GREEN, fmt=FMT_X, align="right")
        _put(wsum, r, 4, f'=IF(AND(ISNUMBER(B{r}),ISNUMBER(C{r})),B{r}/C{r}-1,"NM")',
             fmt='+0.0%;-0.0%;0.0%', align="right")
        _put(wsum, r, 5, f"=Valuation!I{vr}", color=GREEN, fmt=FMT_PRICE, align="right")
        _put(wsum, r, 6, f"=Valuation!K{vr}", color=GREEN, fmt='+0.0%;-0.0%;0.0%', align="right")
        _put(wsum, r, 7, res.selection.reasons.get(key, ""), size=9)
    sr = sh + len(selected_keys) + 1
    _put(wsum, sr, 1, "Blended implied price", bold=True, fill=FILL_TARGET)
    _put(wsum, sr, 5, f"=Valuation!I{br}", color=GREEN, fmt=FMT_PRICE, bold=True, fill=FILL_TARGET, align="right")
    _put(wsum, sr, 6, f"=Valuation!K{br}", color=GREEN, fmt='+0.0%;-0.0%;0.0%', bold=True, fill=FILL_TARGET,
         align="right")
    _put(wsum, sr + 1, 1, "★ = primary multiple used in the blended price", italic=True, size=8)

    ph = sr + 3
    _put(wsum, ph - 1, 1, "Selected peers", bold=True, color=NAVY, size=11)
    _header_row(wsum, ph, ["Ticker", "Company", "Industry", "Market Cap ($mm)", "Similarity Score",
                           "Why it's a comp"])
    score_of = {s.ticker: s for s in (candidates or [])}
    for i, p in enumerate(res.peers):
        r = ph + 1 + i
        s = score_of.get(p.ticker)
        _put(wsum, r, 1, p.ticker, bold=True)
        _put(wsum, r, 2, p.name)
        _put(wsum, r, 3, p.industry)
        _put(wsum, r, 4, (p.market_cap or 0) * 1e-6, color=BLUE, fmt=FMT_MM)
        _put(wsum, r, 5, s.score if s else None, color=BLUE, fmt="0.0")
        _put(wsum, r, 6, "; ".join(s.reasons) if s else "", size=9)
    for c, w in zip("ABCDEFGH", (34, 30, 26, 16, 20, 14, 60, 10)):
        wsum.column_dimensions[c].width = w

    # ---------------------------------------------------------- Peer screen
    if candidates:
        wp = wb.create_sheet("Peer Screen")
        _title(wp, "Peer Candidate Screen",
               "All candidates considered, ranked by similarity score (0–100). Weights: industry 35, "
               "investor grouping 10, size 20, revenue growth 15, EBITDA margin 10, gross margin 10.")
        _header_row(wp, 4, ["Rank", "Ticker", "Company", "Industry", "Market Cap ($mm)", "Rev Growth",
                            "EBITDA Margin", "Industry pts", "Investor pts", "Size pts", "Growth pts",
                            "EBITDA Mgn pts", "Gross Mgn pts", "Total Score", "Selected?"])
        chosen = {p.ticker for p in res.peers}
        for i, s in enumerate(candidates, start=1):
            r = 4 + i
            c = s.company
            fill = FILL_SELECTED if c.ticker in chosen else None
            vals = [i, c.ticker, c.name, c.industry, (c.market_cap or 0) * 1e-6, c.revenue_growth,
                    c.ebitda_margin]
            fmts = [None, None, None, None, FMT_MM, FMT_PCT, FMT_PCT]
            for j, (v, f) in enumerate(zip(vals, fmts), start=1):
                _put(wp, r, j, v, color=BLUE if j >= 5 else BLACK, fmt=f, fill=fill)
            for j, k in enumerate(["industry", "market", "size", "growth", "ebitda_margin", "gross_margin"],
                                  start=8):
                _put(wp, r, j, round(s.components.get(k, 0.0), 1), color=BLUE, fmt="0.0", fill=fill)
            _put(wp, r, 14, f"=SUM(H{r}:M{r})", fmt="0.0", bold=True, fill=fill)
            _put(wp, r, 15, "Yes" if c.ticker in chosen else "", fill=fill, align="center")
        for c, w in zip("ABCDEFGHIJKLMNO", (6, 9, 30, 26, 14, 11, 11, 10, 10, 10, 10, 11, 11, 11, 10)):
            wp.column_dimensions[c].width = w
        wp.freeze_panes = "C5"

    # ------------------------------------------------------------------ Notes
    wn = wb.create_sheet("Notes")
    notes = [
        ("Methodology", None),
        ("Peers", "Candidates come from Yahoo Finance's industry constituents, Yahoo's 'people also watch' "
                  "list and (if thin) sector leaders. Each is scored on industry match, investor grouping, "
                  "market-cap proximity, revenue growth and margin similarity. The top-scoring names are used."),
        ("Multiples", "Chosen by company profile (financial, REIT, unprofitable, high-growth, capital-intensive, "
                      "mature). EV-based multiples use EV = market cap + total debt − cash."),
        ("NM", "A multiple is 'NM' (not meaningful) when the denominator is negative/zero or the result exceeds "
               "the cap in row 4 of the Comps sheet. NM values are excluded from peer statistics."),
        ("NTM Revenue", "Time-weighted blend of consensus current-FY and next-FY revenue estimates."),
        ("Forward Earnings", "Market cap ÷ Yahoo forward P/E (consensus next fiscal year EPS)."),
        ("EBIT", "LTM revenue × LTM operating margin."),
        ("Currency", "Financials reported in a different currency (e.g. ADRs) are converted to the trading "
                     "currency at the spot FX rate."),
        ("Limitations", "Yahoo data can lag filings and does not adjust for one-offs, leases, minority interest "
                        "or preferreds. Always sanity-check key inputs against the 10-K/10-Q before presenting."),
        ("", None),
        ("Colour legend", None),
        ("Blue text", "Hard-coded input pulled from Yahoo Finance – safe to overwrite."),
        ("Black text", "Formula."),
        ("Green text", "Link to another sheet."),
        ("Yellow fill", "Assumption you are expected to adjust (blend flags)."),
    ]
    for i, (k, v) in enumerate(notes, start=1):
        _put(wn, i, 1, k, bold=v is None or True, color=NAVY if v is None else BLACK)
        if v:
            _put(wn, i, 2, v, align="wrap")
    _style(wn["A12"], bold=True, color=BLUE)
    _style(wn["A13"], bold=True, color=BLACK)
    _style(wn["A14"], bold=True, color=GREEN)
    _style(wn["A15"], bold=True, fill=FILL_INPUT)
    wn.column_dimensions["A"].width = 18
    wn.column_dimensions["B"].width = 110

    for sheet in wb.worksheets:
        sheet.sheet_view.showGridLines = False
    return wb


def to_bytes(res: CompsResult, candidates: Optional[list[PeerScore]] = None) -> bytes:
    buf = BytesIO()
    build_workbook(res, candidates).save(buf)
    return buf.getvalue()

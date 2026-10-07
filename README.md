# Comps Analyzer

Type a ticker and get a full comparable-companies analysis:

1. **Peers:** the tool screens up to 20 candidates from Yahoo Finance and ranks them by similarity, then picks the strongest 3–8.
2. **Multiples:** it chooses the multiples that suit the business (banks get P/E and P/B, unprofitable growth names get EV/Revenue, and so on).
3. **Relative value:** it shows where the stock trades versus the peer median and what price peer multiples imply.
4. **Fundamentals check:** it compares growth and margins with the peers, to judge whether a premium or discount is deserved.
5. **Excel export:** a formatted, **formula-driven** Excel comps model you can tweak or drop into a pitch.

The app runs in your browser via [Streamlit](https://streamlit.io). Data comes from Yahoo Finance via `yfinance`, so you don't need an API key.

---

## Quick start (about 5 minutes)

You need **Python 3.10+** ([python.org/downloads](https://www.python.org/downloads/)).

```bash
git clone https://github.com/tkangx/comps.git
cd comps
pip install -r requirements.txt
streamlit run app.py
```

A browser tab opens at `http://localhost:8501`. Enter a ticker (e.g. `AMZN`) and press **Run analysis**.

> No internet? Toggle **Use offline sample data** in the sidebar to demo with bundled, illustrative AMZN data.

### Share it with the whole fund (free)

1. Push this repo to GitHub (already done if you're reading this there).
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in with GitHub, click **Create app**.
3. Pick this repo, branch, and `app.py` → **Deploy**.

You get a URL like `https://smif-comps.streamlit.app` that any member can open without installing anything.

### Command-line / Excel only

```bash
python -m comps AMZN                       # auto peers, saves AMZN_comps.xlsx
python -m comps AMZN --peers 4             # top 4 peers
python -m comps AMZN --with MSFT,GOOGL,WMT # your own peer set
```

---

## How it works

### Peer selection (`comps/peers.py`)
Candidates come from three sources:
- the top companies in the stock's **Yahoo industry**
- Yahoo's **"people also watch"** list, which shows how investors group the stock
- the **sector** leaders, used only when the first two lists are thin

Each candidate is scored 0–100:

| Factor | Points |
|---|---|
| Same industry (same sector only = 14) | 35 |
| In Yahoo "people also watch" | 10 |
| Market-cap proximity (log scale, 0 at 100× apart) | 20 |
| Revenue growth similarity | 15 |
| EBITDA margin similarity | 10 |
| Gross margin similarity | 10 |

The screen also removes duplicate share classes (GOOG/GOOGL), non-equities and foreign duplicate listings. You can add or remove peers in the app at any time.

> **Judgment still matters.** Conglomerates won't fit one industry. Amazon, for example, is classed as *Internet Retail*, so the screen finds BABA, PDD and MELI and misses the AWS business. Add MSFT/GOOGL in the "Add other tickers" box to see how the picture changes. Use the score as a starting point, not the answer.

### Multiple selection (`comps/multiples.py`)

| Profile | Trigger | Multiples |
|---|---|---|
| Financial | Banks, insurers, asset managers, capital markets | Fwd P/E, LTM P/E, P/B |
| REIT | Real Estate + REIT | P/Op. Cash Flow (FFO proxy), EV/EBITDA, P/B |
| Unprofitable | EBITDA ≤ 0 | EV/NTM Revenue, EV/LTM Revenue, EV/Gross Profit |
| High growth | Revenue growth ≥ 20% | EV/NTM Revenue, EV/EBITDA, Fwd P/E, … |
| Capital intensive | Energy, Materials, Utilities | EV/EBITDA, Fwd P/E, P/OCF, … |
| Mature | Everything else | EV/EBITDA, Fwd P/E, EV/EBIT, LTM P/E, P/FCF |

A multiple is **NM** (not meaningful) when its denominator is negative or when it is an extreme outlier (e.g. P/E > 150×). NM values are excluded from peer statistics. The ★ multiples feed the **blended implied price**.

### Implied price
- EV multiples: `(peer multiple × target metric − net debt) ÷ shares`
- Equity multiples: `peer multiple × target metric ÷ shares`

### Excel workbook
| Sheet | Contents |
|---|---|
| **Summary** | Verdict, multiples vs. peers, implied prices, peer rationale |
| **Comps** | All inputs (blue) and every multiple as a live formula, plus peer max/75th/mean/median/25th/min and the target's premium |
| **Valuation** | Implied price at peer 25th/median/75th; toggle the yellow **Y/N** cells to change which multiples feed the blend |
| **Peer Screen** | Every candidate with its score breakdown |
| **Notes** | Methodology and colour legend |

## Limitations
- Yahoo data can lag filings, and Yahoo does not adjust for one-offs, leases, minority interest or preferreds. **Check key inputs against the 10-K/10-Q before presenting.**
- EBIT is approximated as LTM revenue × operating margin.
- NTM revenue is a time-weighted blend of the current-FY and next-FY consensus.
- Forward P/E comes from Yahoo's consensus figure, so it can be stale for ADRs and thinly covered names.

## Development
```bash
pip install pytest
python -m pytest
```
The tests run fully offline using `comps/sample.py` and a fake Yahoo client.

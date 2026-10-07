# Comps Analyzer

Type a ticker and get a full comparable-companies analysis:

1. **Peers:** the tool screens up to 20 candidates from Yahoo Finance and ranks them by similarity, then picks the strongest 3–8.
2. **Multiples:** it chooses the multiples that suit the business (banks get P/E and P/B, unprofitable growth names get EV/Revenue, and so on).
3. **Relative value:** it shows where the stock trades versus the peer median and what price peer multiples imply.
4. **Fundamentals check:** it compares growth and margins with the peers, to judge whether a premium or discount is deserved.
5. **Excel:** a formatted, **formula-driven** comps model. Type a new company into one cell and it rebuilds itself.

The app runs in your browser via [Streamlit](https://streamlit.io). Data comes from Yahoo Finance via `yfinance`, so you don't need an API key.

---

## Three ways to use it

| | Best for | You do |
|---|---|---|
| **Excel: `CompsModel.xlsx`** | Day-to-day work, pitches, tweaking numbers | Type a company in a cell → click **Run main** |
| **Web app** | Quick looks, sharing a link with the fund | Type a company → **Run analysis** |
| **Command line** | Batch runs | `python -m comps AMZN` |

All three run the same engine. You can type a ticker (`AMZN`) or a name (`Amazon`, `Coca-Cola`).

---

## Excel setup (Windows, one time, about 10 minutes)

1. **Install Python** from [python.org/downloads](https://www.python.org/downloads/). On the first installer screen, tick **"Add python.exe to PATH"**.
2. **Download this repo:** green **Code** button → **Download ZIP**, then unzip it somewhere permanent (e.g. `Documents\Comps`). If you use git, you can clone it instead.
3. **Open Command Prompt in that folder:** in File Explorer, click the address bar, type `cmd` and press Enter. Then run:
   ```
   pip install -r requirements.txt
   xlwings addin install
   ```
4. **Open `CompsModel.xlsx`** from that folder. You'll see a new **xlwings** tab in the Excel ribbon.

### Daily use
1. Type a company name or ticker into the **yellow cell** on the Dashboard (e.g. `Nike` or `NKE`).
2. Optionally change the number of peers, or list your own peers (`MSFT, GOOGL`).
3. Click **xlwings** tab → **Run main**. After 10–30 seconds the status line shows ✔ and the whole model updates.

Once data has loaded, everything is live Excel:
- **Multiple dropdowns** on the Dashboard swap which multiples are shown.
- **Y/N blend flags** choose which multiples feed the blended implied price.
- **Blue inputs** on the Comps sheet can be overwritten with your own numbers (say, adjusted EBITDA from the 10-K). Every multiple, statistic and implied price recalculates instantly.

> **Keep the file name `CompsModel.xlsx`.** It must stay in this folder, next to `CompsModel.py` and the `comps` folder. "Run main" works by finding the Python file with the same name as the workbook. To keep a copy for a pitch, use **File → Save a Copy**. The copy keeps all values and formulas.

### Optional: refresh automatically when the cell changes
1. **File → Save As → Excel Macro-Enabled Workbook** and save it as `CompsModel.xlsm` in the same folder.
2. Press **Alt+F11**, then **Tools → References**, and tick **xlwings**.
3. In the left panel, double-click the **Dashboard** sheet and paste:
   ```vba
   Private Sub Worksheet_Change(ByVal Target As Range)
       Dim inputs As Range
       Set inputs = Union(Me.Range("CompsInput"), Me.Range("CompsNumPeers"), Me.Range("CompsCustomPeers"))
       If Not Intersect(Target, inputs) Is Nothing Then
           RunPython "import comps.xl; comps.xl.refresh()"
       End If
   End Sub
   ```
4. *(Optional button)* **Insert → Module**, then paste the macro below. Back in Excel, choose **Insert → Shapes**, draw a button, right-click it → **Assign Macro** → `RefreshComps`.
   ```vba
   Sub RefreshComps()
       RunPython "import comps.xl; comps.xl.refresh()"
   End Sub
   ```

Now typing a new company and pressing Enter rebuilds the model by itself.

### Optional: custom functions in any cell
In a macro-enabled workbook (`.xlsm`):
1. Turn on **File → Options → Trust Center → Trust Center Settings → Macro Settings → "Trust access to the VBA project object model"**.
2. On the **xlwings** tab, type `comps.xl` in the **UDF Modules** box and click **Import Functions**.

Then use these anywhere:

| Formula | Returns |
|---|---|
| `=COMPS_TICKER("Coca-Cola")` | `KO` |
| `=COMPS_NAME("NKE")` | `NIKE, Inc.` |
| `=COMPS_PEERS("AMZN", 5)` | 5 peer tickers, spilling down |
| `=COMPS_FIELD("MSFT", "ebitda_margin")` | any data field: `price`, `market_cap`, `ev`, `revenue`, `ebitda`, `revenue_growth`, … |
| `=COMPS_MULTIPLE("MSFT", "EV/EBITDA")` | the multiple, or `NM` |
| `=COMPS_PEER_MEDIAN("AMZN", "Fwd P/E")` | peer-median multiple |
| `=COMPS_IMPLIED_PRICE("AMZN", "EV/EBITDA")` | implied price at the peer median |

Multiple names are flexible: `EV / EBITDA`, `ev/ebitda`, `P/E`, `Fwd P/E`, `EV/Sales`, `P/B`, `P/FCF`. Results are cached for an hour, so a sheet full of formulas stays fast.

### Troubleshooting
| Problem | Fix |
|---|---|
| No **xlwings** tab | Run `xlwings addin install` again, then restart Excel |
| "Python not found" / nothing happens | Run `where python` in Command Prompt and paste that path into the **Interpreter** box on the xlwings tab |
| `ModuleNotFoundError: comps` | The workbook isn't in the repo folder; move it back next to the `comps` folder |
| Status says ✖ Couldn't find a stock | Try the exact ticker, e.g. `BRK-B`, `RDS.A` |
| Some multiples show NM | The company has negative/zero earnings for that metric, or it's an outlier; see the caps in row 4 of the Comps sheet |

---

## Web app

```bash
pip install -r requirements.txt
streamlit run app.py
```
A browser tab opens at `http://localhost:8501`. Toggle **Use offline sample data** to demo without internet.

**Share it with the whole fund (free):** go to [share.streamlit.io](https://share.streamlit.io), sign in with GitHub, click **Create app**, pick this repo and `app.py`, then **Deploy**. Members get a link and don't need to install anything. The app's **Download Excel** button produces the same `CompsModel` workbook.

## Command line

```bash
python -m comps AMZN                       # auto peers, saves AMZN_comps.xlsx
python -m comps "Coca-Cola" --peers 4      # names work too
python -m comps AMZN --with MSFT,GOOGL,WMT # your own peer set
python -m comps.excel CompsModel.xlsx      # regenerate the blank template (sample data)
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
| **Dashboard** | Input cell, headline stats, multiples (dropdowns) with premium and implied price, blended value, peer list with rationale |
| **Comps** | All inputs (blue) and every multiple as a live formula, plus peer max/75th/mean/median/25th/min and the target's premium |
| **Valuation** | Implied price at peer 25th/median/75th for each Dashboard multiple |
| **Peer Screen** | Every candidate with its score breakdown |
| **Lists** | Lookup table behind the dropdowns (edit the "why" text freely) |
| **Notes** | How-to and methodology |

Python only writes the blue cells, through named ranges such as `CompsInput` and `CompsData`. Everything else is a formula, so you can restyle the sheets or add your own calculations, and refreshes won't overwrite them.

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
The tests run fully offline. They use `comps/sample.py`, a fake Yahoo client and a fake xlwings workbook, and if LibreOffice is installed they also recalculate the Excel model to check for formula errors.

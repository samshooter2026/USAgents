# SPX Pricer — Handoff (2026-05-07)

## What this is

SPX Combo & Revcon pricer. Reads live Bloomberg from `Combos_RevCon_Pricer.xlsx`
via xlwings. Streamlit UI at localhost:8501. CLI also available.

**Code lives at:** `C:\Users\sshooter\US Pricers\spx_pricer\`

**To start:** `python -m streamlit run app.py` (from `spx_pricer\`)

---

## Status

All phases complete and the pricer is **calibration-verified**. 52/52 tests
passing.  BBG Data sheet wired with live ASD and AXW formulas.  AXW units
confirmed as **bps** (handoff May 6).

UI now has **five tabs** in this order: **Run · Ladder · Revcon · Reverse Borrow · Agent Deployment**.

---

## What changed on 2026-05-07

This is a record of every behaviour change since the May 6 handoff.

### UI polish

- Top-bar clock and Run-tab header timestamps now show **America/New_York
  time** (auto-adjusts EDT/EST).  Previously showed UK local.
- **Run** is now the leftmost tab.  Order: Run · Ladder · Revcon · Reverse
  Borrow · Agent Deployment.
- The Run tab's `% of spot` block was replaced with a styled HTML panel:
  numbers are 1.25rem bold monospace, and any bid/ask that moves up flashes
  light blue, any that moves down flashes light pink (1.8s ease-out).  Plain
  text for copying still available below in a code block.

### Run-tab width controls

- Default width is now **25bps** (was 10bps).  Slider extends to **1% (100bps)**.
- Per-expiry width sliders sit to the **left of the run output**, one per
  expiry.  Each slider uses a unique `session_state` key so adjustments
  persist across the 5-second auto-refresh.
- `defaults.yaml` carries per-expiry baselines under `per_expiry_widths:` —
  current values:
  - DEC26 = 15bp
  - DEC27 = 15bp
  - DEC28 = 20bp
  - DEC29 = 35bp
  - DEC30 = 50bp
  - DEC31 = 50bp
- The agent applies these defaults automatically.  Whenever the YAML changes
  (or a fresh browser session opens), every slider snaps back to the
  configured default.  Within a stable YAML version, your manual moves
  persist.
- Two buttons in the top-right of the Run tab:
  - **Set all = default**: applies the global slider value to every expiry
  - **Reset to per-expiry**: re-applies the YAML-configured per-expiry values

### Bloomberg staleness — three new safety nets

These were added after a debugging session where ES turned out to be a
hard-coded `7370.5` (someone had typed over `=BDP("ES1 Index","PX_LAST")`)
and SPX was frozen at `7345.52` because Bloomberg's RTD subscription had
silently stalled.  Diagnostic at `spx_pricer/debug_bbg.py` (run with
`python debug_bbg.py`).

| Safety net | What it does |
|---|---|
| **Fix BDP** button (top bar) | Clears and re-enters the SPX and ES BDP formulas, forcing a fresh Bloomberg request.  One click resolves both "frozen RTD" and "someone typed over a BDP cell". |
| **Staleness detector** | Compares SPX across reads.  If SPX is identical for >60s, the SPX number turns orange in the top bar and a banner directs you to press Ctrl+Alt+F9 or click Fix BDP.  Catches frozen Bloomberg silently — the old "0s ago" timestamp was just when Python last ran, not when Bloomberg last pushed. |
| **Cell protection** in `_rebuild_sheet.py` | All BDP formula cells (B3, B4, B5, SOFR strip, ASD strip, AXW strip) are locked when the sheet is rebuilt.  Green override cells (div yield + borrow per expiry) remain editable.  Sheet is protected without password — power users can unprotect via Excel ribbon if needed. |

### Workbook path mirroring

Excel had been opening the workbook from `C:\blp\API\Office Tools\` (Bloomberg
add-in directory) while the pricer's `excel_map.yaml` pointed at
`C:\Users\sshooter\US Pricers\`.  The pricer worked because xlwings matches
open workbooks by filename, not full path, but the two copies were diverging.
On 2026-05-07 the live copy at `C:\blp\...` was saved and then mirrored back
to `C:\Users\sshooter\US Pricers\Combos_RevCon_Pricer.xlsx`, with a backup
written at `Combos_RevCon_Pricer.xlsx.bak.20260507_155543`.  The two paths are
now in sync.

### Agent Deployment tab — AXW Calibration Agent

A fifth tab.  Auto-logs every street combo print and flags AXW drift.

**Storage**: `spx_pricer/logs/axw_calibration.parquet`.

**Paste workflow** (primary input):
- Paste a one-line print like
  `20:38:10 SPX Dec26 combo trades 650mm ~101.9275% makes 1.2b`
- Parser extracts: trade time, expiry, notional (`Xmm` or `Xb`), price (% of
  spot or pts bid/ask), and ES level (`Xf`) if present
- For % of spot quotes, implied borrow is computed directly from the F/S
  ratio (no spot needed — `b = ln((F/S)/(1−qT))/T − r`).  Avoids the
  imprecision of guessing the spot at trade time.
- Persisted with: timestamp, expiry, size, F/S, ES level, spot-at-trade,
  source/counterparty, implied borrow, AXW at log time, delta

**Manual form** (fallback): in an expander below the paste box.  Use it for
points-format prints where the strike must be specified explicitly.

**Drift dashboard**: per-expiry table showing AXW now, count of recent prints,
mean delta vs current AXW, status indicator, and recommended override.

| Indicator | Rule |
|---|---|
| 🟢 | within ±5bp of current AXW |
| 🟡 | mean drift exceeds 5bp but not consistent same-side |
| 🔴 | ≥3 consecutive same-side prints, mean \|Δ\| > 5bp |

When 🔴 fires, a "Recommend" column suggests an override of `AXW + 0.5 ×
mean_drift` (conservative shift halfway toward observed flow).

**Recent prints log**: filterable by expiry and time window (1d/7d/30d/90d/365d).
Shows per row: Time, Exp, K, Combo, ES@trade, mm, Impl bp, AXW@trade, AXW now,
Δ today, Source.

**Drift thresholds** are constants at the top of `pricer/calibration.py`.

---

## Bloomberg sheet layout (`BBG Data` in `Combos_RevCon_Pricer.xlsx`)

Built by `_rebuild_sheet.py`. Run this script whenever you add a new expiry
or need to re-map cells. It also regenerates `config/excel_map.yaml` and
applies cell protection.

Sections in order:

1. **SPOT & FUTURES** — SPX spot, ES1, ES basis (auto ES−SPX). All live `BDP`.
2. **SOFR CURVE** — 9-tenor curve (ON/1M/3M/6M/1Y/2Y/3Y/5Y/10Y). All live `BDP`.
3. **RAW ASD STRIP** — `=BDP("ASDZxx Index","LAST_PRICE")` for Z26–Z33. Blue cells = live Bloomberg.
4. **DIV YIELD PER EXPIRY** — formula: `=ASDZxx / SPX * 100` (% ann). Green cells = computed; type a number to override.
5. **RAW AXW STRIP** — `=BDP("AXWZxx Index","LAST_PRICE")` for Z26–Z33. Blue cells = live Bloomberg.
6. **BORROW PER EXPIRY** — formula: `=AXWZxx` (bps). Green cells = computed; type a number to override.

After running `_rebuild_sheet.py`, all Bloomberg formula cells are locked
(blue cells, basis, SOFR strip, raw ASD/AXW).  Green override cells stay
editable.  Sheet protection has no password — right-click any sheet tab →
Unprotect Sheet to remove if needed.

---

## Confirmed working Bloomberg tickers

| What | Ticker | Field |
|---|---|---|
| SPX spot | `SPX Index` | `PX_LAST` |
| ES front | `ES1 Index` | `PX_LAST` |
| SOFR ON | `SOFRRATE Index` | `PX_LAST` |
| SOFR 1M | `TSFR1M Index` | `PX_LAST` |
| SOFR 3M | `TSFR3M Index` | `PX_LAST` |
| SOFR 6M | `TSFR6M Index` | `PX_LAST` |
| SOFR 1Y–10Y | `USSO1/2/3/5/10 Curncy` | **`LAST_PRICE`** (not PX_LAST!) |
| ASD strip | `ASDZ26`…`ASDZ33 Index` | **`LAST_PRICE`** |
| AXW (AIR) strip | `AXWZ26`…`AXWZ33 Index` | **`LAST_PRICE`** |

ASD and AXW are CME delayed — allow ~40s after Bloomberg connects for cells
to populate. If cells show `#N/A`: press **Ctrl+Alt+F9** in Excel.  If a BDP
cell silently freezes (returns the same number across reads while the market
is open), click **Fix BDP** in the pricer top bar.

`PX_LAST` vs `LAST_PRICE`: not interchangeable.  Each ticker only populates
the field Bloomberg has assigned to it.  Equity-class instruments
(SPX, ES1, SOFR fixings) use `PX_LAST`.  Derivatives / OIS swaps / CME-AIR /
CME-ASD use `LAST_PRICE`.  Querying the wrong field returns `#N/A`.

---

## Live values as of 2026-05-07 (open of session)

| | |
|---|---|
| SPX spot | 7,365.12 |
| ES1 | 7,401.25 |
| ES basis | +36.13 |
| SOFR ON | 3.62% |
| SOFR 1Y | 3.712% |
| SOFR 10Y | 3.892% |

**ASD strip (index pts/year):**
ASDZ26=82.75, ASDZ27=83.85, ASDZ28=83.75, ASDZ29=83.9, ASDZ30=83.85,
ASDZ31=84.15, ASDZ32=84.75, ASDZ33=84.5

**AXW strip (bps):**
AXWZ26=56.5, AXWZ27=66.0, AXWZ28=74.0, AXWZ29=81.5, AXWZ30=89.5,
AXWZ31=97.0, AXWZ32=103.0, AXWZ33=115.0

**Per-expiry (computed):**

| Expiry | Settle | Div yield | Borrow |
|---|---|---|---|
| DEC26 | 2026-12-18 | 1.1218% | 56.5 bps |
| DEC27 | 2027-12-19 | 1.1367% | 66.0 bps |
| DEC28 | 2028-12-15 | 1.1355% | 74.0 bps |
| DEC29 | 2029-12-21 | 1.1374% | 81.5 bps |
| DEC30 | 2030-12-20 | 1.1380% | 89.5 bps |
| DEC31 | 2031-12-19 | 1.1441% | 97.0 bps |

---

## Calibration verification (cumulative log)

The pricer is validated against real prints.  Each row records the observed
flow vs the pricer at trade-time conditions.

**1. SPX Dec31 Combo (7153f, LIVE) 121.52/121.78 — Apr 22, 2026**
- "7153f" = ES front-month at trade time = 7153 (gives implied SPX spot ~7118)
- Back-solved strike: K = 8575 (clean SPX 25-pt grid)
- Pricer at K=8575: combo mid ≈ 121.4 vs market 121.65 → **gap of 0.25 pts**,
  well inside the 0.26 bid-ask ✓

**2. SPX Dec26 combo — 550mm at 101.9275% / 7338f — May 6, 2026**
- 250mm @ 15:06 + 300mm @ 15:56, both at the identical print
- Pricer F/S at trade-time conditions = 101.9503% → gap of 2.28 bps in F/S =
  **3.6 bps in implied borrow**
- Implied DEC26 borrow on the print = 51.4 bps vs AXWZ26's 55.0 bps
- Within ±5bp tolerance.  Front-end may sit a touch rich.

**3. SPX Dec30 combo — 100mm at 117.21% / 7329f — May 6, 2026**
- Pricer F/S = 116.9694% → gap of 24bp in F/S = **3.8 bps in implied borrow**
- Implied DEC30 borrow on the print = 93.3 bps vs AXWZ30's 89.5 bps
- Year-end → small Q4-squeeze premium expected.  4bp fits.

**Calibration verdict (May 6):** AXW = bps direct unit confirmed by 650mm of
flow.  Front-end may run a touch rich, year-end a touch cheap — both sides
within ±5bp of model.  Pricer is in calibration.

**Calibration agent (May 7 onward):** the Agent Deployment tab now persists
every street print observed and flags drift automatically.  Past prints
should be backfilled via the paste box as you encounter them in chat history.

---

## Key math

```
F        = (S − S×q×T) × exp((r+b)×T)     # linear div, ACT/365
Combo(K) = exp(−rT) × (F − K)
F_combo  = K + Combo × exp(rT)             # only r, not r+b — borrow embedded
bps_ann  = (Edge / Spot) / T × 10000
Verdict  : LIFT if bps_ann > +1, PASS if < −1, FLAT otherwise

Implied borrow from a % of spot quote (no spot needed):
F/S = (1 − qT) × exp((r+b)T)
b   = ln((F/S) / (1 − qT)) / T − r
```

PDF revcon §6 rounds 20.4775→20.48 → P&L stated as −$204,800; full precision
gives −$204,775. Tests use ±$30 tolerance.

---

## Quote conventions you'll see in chat

| Format | What it means | Example |
|---|---|---|
| `XXX.XX / YYY.YY  (Zf)` | Combo mid in **points** at a known strike, ES ref Z | `121.52 / 121.78  (7153f)` |
| `XXX.XXXX%  (Zf)` | Forward as **% of SPX spot**, ES ref Z | `101.9275%  (7338f)` |
| `HH:MM:SS SPX <Exp> combo trades <X>mm ~<P>% makes <Y>b` | Time-stamped chat one-liner.  `<X>mm` = trade size, `<P>%` = F/S, `<Y>b` = running cumulative.  Paste these directly into the Agent Deployment tab. | `20:38:10 SPX Dec26 combo trades 650mm ~101.9275% makes 1.2b` |

The Run tab outputs the first two formats.  The Agent Deployment tab parses
all three.

---

## Borrow convention

Excel stores borrow as **bps** (e.g., 55 = 55 bps).  `_bps_to_decimal()` in
`excel.py` divides by 10,000.  Never use `_to_decimal()` for borrow — it
would read 55 as 55%.

---

## UI guide — `localhost:8501`

### Top bar (always visible)

`Snapshot HH:MM:SS NY (Xs ago)   SPX 7,365.12   ES 7,401.25 basis +36.1   [Refresh] [Fix BDP]`

- **Green** = fresh (<30s), **orange** = stale soon, **red** = stop, refresh Bloomberg
- SPX number turns orange + yellow banner if SPX hasn't moved for >60s during market hours (Bloomberg RTD probably stalled)
- **Refresh** clears the cache and re-reads Excel
- **Fix BDP** clears and re-enters the SPX + ES BDP formulas (use when SPX is frozen or ES has been overwritten with a typed number)

### Tab 1 — Run (NEW first tab)

Generates a full run across all expiries in both quote formats:

- **Default width** slider (5–100bp, default 25)
- **Per-expiry width** sliders (one per expiry, default per `defaults.yaml`,
  persistent across refresh, auto-reset when YAML changes)
- **Strike-mode** radio: ATM forward (round-100) or ES level (round-25)
- **Override ES level** checkbox + ES input — reprices the whole run against
  any ES level (e.g. matching a broker print at "7338f").  Header shows
  `ES X (LIVE Y)` when overridden.
- **Set all = default** / **Reset to per-expiry** buttons
- Output:
  - styled `% of spot` block with flashing numbers (light blue = up, light pink = down)
  - run table with `Width / F / F/S Bid/Mid/Ask / K / Pts Bid/Mid/Ask / Borrow`
  - copy blocks (% of spot text, points text)

### Tab 2 — Ladder

Centred on the SPX **forward** for the chosen expiry.  Footer shows
`F  T  r  b  q  DF  ES`.

**🎯 Match a counterparty quote** expander:
- **Override ES futures ref** — type a specific ES level (e.g. 7153)
- **Show one specific strike** — replaces the 7-strike ladder with a single row
- **Quoted bid / ask** — get a side-by-side comparison with delta in points and bps of forward

### Tab 3 — Revcon

Inputs (expiry / strike / combo mid / ES / basis / side / lots / **Basis
confirmed?**), output is LIFT / PASS / FLAT (±1 bps annualised threshold)
plus breakdown table.

### Tab 4 — Reverse Borrow

Backs out implied borrow from a street combo print.  RICH/CHEAP tile shows
delta vs your AXW model.  Expander shows the implied borrow curve across all
expiries.

### Tab 5 — Agent Deployment (NEW)

AXW calibration agent.  See "What changed on 2026-05-07 → Agent Deployment"
above for the full feature list.  Daily workflow:

1. Whenever a street print drops in chat, paste it into the box at the top of
   the tab and click **Parse and log →**
2. Glance at the drift dashboard.  Any 🔴 = AXW is off.  Recommend column
   suggests where to override.
3. To override: open Excel, edit the green AXW cell for that expiry to the
   recommended value.  The pricer reads the override on next refresh.
4. Don't override on a single print — flow noise is real.  Wait for the 🔴
   alert (≥3 consecutive same-side, mean >5bp).

---

## Reading the run / sanity checks before sending

- Top bar **green**?  SPX not stuck (no orange/banner)?
- F values ladder up by ~3-3.5%/year?
- Borrow column ladders up DEC26 → DEC31?  Any inversion = investigate
- ES level shown on every line matches live ES?
- Per-expiry width sliders show what you expect?

---

## Verifying any new street print (30-second drill)

Two paths now:

**Path A — quick check (Reverse Borrow tab):**
1. Open **Tab 4 — Reverse Borrow**
2. Pick the expiry, type the strike (use the ES round number for %-of-spot trades)
3. Type the street's combo mid in points
4. Read the **Delta** in bps.  |Δ| < 5bp = pricer right.  ≥10bp = investigate.

**Path B — log it for drift tracking (Agent Deployment tab):**
1. Open **Tab 5 — Agent Deployment**
2. Paste the chat one-liner (or use manual form for points format)
3. Hit **Parse and log →**
4. The print is persisted, drift dashboard updates, and 🔴 fires automatically
   when the rolling pattern crosses the threshold

Use Path B for every print you see.  Path A is for one-off sanity checks.

For %-of-spot prints (e.g. "101.9275% / 7338f"), the **Run** tab with ES
override is also good for visual comparison:
1. Tick **Override ES level**, type the broker's ES (e.g. 7338)
2. Compare your run line to the print
3. Or compute by hand:
   ```python
   import math
   gap_bps_borrow = math.log(trade_FoS / pricer_FoS) / T * 10000
   ```

---

## When to manually override AXW

AXW gives a baseline borrow each day.  Override the green Excel cell when:

- The **🔴 alert** in the Agent Deployment tab fires (≥3 consecutive
  same-side prints with mean >5bp drift)
- Year-end / quarter-end where AXW lags the squeeze (Q4 DEC26/DEC30
  typically rich 3-6 bps)
- A specific event (dividend recapture, special borrow demand) AXW can't
  see yet

Don't override on a single print — flow-of-the-day noise is real.  The 550mm
DEC26 print on May 6 at 3.6bp off was *not* worth overriding (one cluster,
within tolerance).  If the same level holds across the next morning's flow
too, the agent will flash 🔴 and that's when to nudge.

---

## File map

| File | Purpose |
|---|---|
| `spx_pricer/_rebuild_sheet.py` | Rebuilds BBG Data sheet + writes `config/excel_map.yaml` + applies cell protection.  Re-run to add expiries. |
| `spx_pricer/debug_bbg.py` | xlwings/Bloomberg connectivity diagnostic.  Run when prices look frozen. |
| `spx_pricer/config/excel_map.yaml` | Auto-generated cell map.  Do not edit by hand. |
| `spx_pricer/config/defaults.yaml` | Width defaults (25bp D2D, 30bp client), per-expiry widths, stale thresholds, Q4 extra 5bp. |
| `spx_pricer/pricer/math.py` | Pure math, no I/O.  All PDF formulas + `implied_borrow_from_fs`. |
| `spx_pricer/pricer/excel.py` | xlwings I/O.  `read_snapshot()` → `MarketSnapshot`. |
| `spx_pricer/pricer/ladder.py` | `build_ladder(snap, expiry, width_bps, es_override=None, specific_strike=None)` |
| `spx_pricer/pricer/revcon.py` | `calc_revcon(...)` |
| `spx_pricer/pricer/reverse.py` | `calc_reverse(...)` — implies borrow from street combo |
| `spx_pricer/pricer/calibration.py` | **NEW** AXW calibration agent: `parse_print_text`, `log_print`, `read_log`, `detect_drift`, `implied_borrow_from_pct_print`. |
| `spx_pricer/pricer/cli.py` | Typer CLI: `ladder`, `revcon`, `reverse` |
| `spx_pricer/app.py` | Streamlit UI.  5 tabs: Run, Ladder, Revcon, Reverse Borrow, Agent Deployment. |
| `spx_pricer/logs/quotes.parquet` | Quote log (auto-created on first Revcon calc). |
| `spx_pricer/logs/axw_calibration.parquet` | **NEW** AXW calibration log (auto-created on first paste). |
| `spx_pricer/tests/` | 52 tests.  Run: `pytest` from `spx_pricer\`. |

---

## To add a new expiry

1. Add to `EXPIRIES` list in `_rebuild_sheet.py`
2. Add to `EXPIRY_TO_YEAR` dict in `_rebuild_sheet.py`
3. Add the new code to `defaults.yaml` under `per_expiry_widths:` with a
   sensible bps default
4. Run `python _rebuild_sheet.py`
5. `config/excel_map.yaml` is regenerated automatically; new expiry appears
   in all five tab dropdowns

ASD/AXW strips currently extend to Z33 — expiries through DEC33 can be added
without changing the BDP layout.

---

## Daily ritual

1. Open `Combos_RevCon_Pricer.xlsx` in Excel — Bloomberg connects automatically
2. Wait ~40s for delayed CME data (ASD/AXW)
3. Glance at green cells: any `#N/A`?  Press **Ctrl+Alt+F9**
4. Run `python -m streamlit run app.py` from `spx_pricer\` (or refresh the browser tab)
5. **Check top bar**: SPX/ES live (no orange/banner), basis sensible
6. **Tab 1 → Run** to see your morning quotes; sanity-check borrow curve
7. **Tab 5 → Agent Deployment** — paste any overnight prints into the log
8. Watch for 🔴 drift alerts during the day; override green AXW cells when justified

---

## Open calibration questions

All resolved.  AXW units (bps direct) confirmed by flow.  The agent now
maintains calibration continuously.

---

## Path / file gotchas

- Excel sometimes opens the workbook from `C:\blp\API\Office Tools\` rather
  than `C:\Users\sshooter\US Pricers\`.  xlwings finds it by filename either
  way.  If you save the workbook from Excel's Save dialog, check the location
  it suggests — accepting the default may save into the Bloomberg directory
  and the two copies will diverge.  Mirror with:
  ```python
  import shutil
  shutil.copy2(
      r"C:\blp\API\Office Tools\Combos_RevCon_Pricer.xlsx",
      r"C:\Users\sshooter\US Pricers\Combos_RevCon_Pricer.xlsx",
  )
  ```

- The pricer's "0s ago" timestamp is when **Python** last ran, not when
  Bloomberg last pushed.  The staleness detector works around this for SPX,
  but if any Bloomberg cell looks frozen, click **Fix BDP** or
  Ctrl+Alt+F9 in Excel.

- `BDP` is a "pull" function.  Excel only re-evaluates BDP when it
  recalculates (forced via xlwings every 5s).  If Bloomberg's RTD
  subscription stalls, BDP keeps returning the last cached value silently.
  Clearing and re-entering the formula (what **Fix BDP** does) is what
  reliably forces a fresh request.

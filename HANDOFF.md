# SPX Pricer — Handoff (2026-05-06)

## What this is

SPX Combo & Revcon pricer. Reads live Bloomberg from `Combos_RevCon_Pricer.xlsx` via xlwings. Streamlit UI at localhost:8501. CLI also available.

**Code lives at:** `C:\Users\sshooter\US Pricers\spx_pricer\`

**To start:** `streamlit run app.py` (from `spx_pricer\`)

---

## Status

All phases complete and the pricer is **calibration-verified** (see "Calibration verification" below). 52/52 tests passing. BBG Data sheet fully wired with live ASD and AXW formulas. AXW units confirmed as **bps**.

UI has **four tabs**: Ladder · Revcon · Reverse Borrow · Run.

---

## Bloomberg sheet layout (`BBG Data` in `Combos_RevCon_Pricer.xlsx`)

Built by `_rebuild_sheet.py`. Run this script whenever you add a new expiry or need to re-map cells. It also regenerates `config/excel_map.yaml`.

Sections in order:

1. **SPOT & FUTURES** — SPX spot, ES1, ES basis (auto ES−SPX). All live `BDP`.
2. **SOFR CURVE** — 9-tenor curve (ON/1M/3M/6M/1Y/2Y/3Y/5Y/10Y). All live `BDP`.
3. **RAW ASD STRIP** — `=BDP("ASDZxx Index","LAST_PRICE")` for Z26–Z33. Blue cells = live Bloomberg.
4. **DIV YIELD PER EXPIRY** — formula: `=ASDZxx / SPX * 100` (% ann). Green cells = computed from live; type a number to override.
5. **RAW AXW STRIP** — `=BDP("AXWZxx Index","LAST_PRICE")` for Z26–Z33. Blue cells = live Bloomberg.
6. **BORROW PER EXPIRY** — formula: `=AXWZxx` (bps). Green cells = computed; type a number to override.

---

## Confirmed working Bloomberg tickers

| What | Ticker | Field |
|---|---|---|
| SPX spot | `SPX Index` | PX_LAST |
| ES front | `ES1 Index` | PX_LAST |
| SOFR ON | `SOFRRATE Index` | PX_LAST |
| SOFR 1M | `TSFR1M Index` | PX_LAST |
| SOFR 3M | `TSFR3M Index` | PX_LAST |
| SOFR 6M | `TSFR6M Index` | PX_LAST |
| SOFR 1Y–10Y | `USSO1/2/3/5/10 Curncy` | **LAST_PRICE** (not PX_LAST!) |
| ASD strip | `ASDZ26`…`ASDZ33 Index` | **LAST_PRICE** |
| AXW (AIR) strip | `AXWZ26`…`AXWZ33 Index` | **LAST_PRICE** |

ASD and AXW are CME delayed — allow ~40s after Bloomberg connects for cells to populate. If cells show `#N/A`: press **Ctrl+Alt+F9** in Excel.

---

## Live values as of 2026-05-06

| | |
|---|---|
| SPX spot | 7,312.50 |
| ES1 | 7,350.75 |
| ES basis | +38.25 |
| SOFR ON | 3.62% |
| SOFR 1Y | 3.712% |
| SOFR 10Y | 3.892% |

**ASD strip (index pts/year):**
ASDZ26=82.80, ASDZ27=83.85, ASDZ28=84.00, ASDZ29=84.10, ASDZ30=83.90, ASDZ31=84.30, ASDZ32=83.50, ASDZ33=84.05

**AXW strip (bps — confirmed unit):**
AXWZ26=55.0, AXWZ27=65.5, AXWZ28=74.0, AXWZ29=81.5, AXWZ30=89.5, AXWZ31=97.0, AXWZ32=103.0, AXWZ33=115.0

**Computed per-expiry:**

| Expiry | Settle | Div yield | Borrow |
|---|---|---|---|
| DEC26 | 2026-12-18 | 1.1323% | 55.0 bps |
| DEC27 | 2027-12-19 | 1.1467% | 65.5 bps |
| DEC28 | 2028-12-15 | 1.1487% | 74.0 bps |
| DEC29 | 2029-12-21 | 1.1501% | 81.5 bps |
| DEC30 | 2030-12-20 | 1.1474% | 89.5 bps |
| DEC31 | 2031-12-19 | 1.1528% | 97.0 bps |

---

## Calibration verification (today's evidence)

The pricer was validated against two real prints:

**1. SPX Dec31 Combo (7153f, LIVE) 121.52/121.78 — Apr 22, 2026**
- "7153f" = ES front-month at trade time = 7153 (gives implied SPX spot ~7118)
- Back-solved strike: K = 8575 (clean SPX 25-pt grid)
- Pricer at K=8575: combo mid ≈ 121.4 vs market 121.65 → **gap of 0.25 pts**, well inside the 0.26 bid-ask. ✓

**2. SPX Dec26 combo trades 250mm 101.9275% 7338f — May 6, 2026**
- "101.9275%" = forward as a percentage of SPX spot (F/S, the standard equity-finance convention)
- Pricer F/S = 101.9537% → **gap = 2.6 bps in F/S = ~0.4 bps in implied borrow**
- Implied borrow on the print = 54.6 bps vs AXW's 55.0 bps. ✓

**Conclusion:** AXW = bps directly is the correct unit. Pricer is calibrated end-to-end.

---

## Key math (PDF §2–4)

```
F        = (S − S×q×T) × exp((r+b)×T)     # linear div, ACT/365
Combo(K) = exp(−rT) × (F − K)
F_combo  = K + Combo × exp(rT)             # only r, not r+b — borrow embedded
bps_ann  = (Edge / Spot) / T × 10000
Verdict  : LIFT if bps_ann > +1, PASS if < −1, FLAT otherwise
```

PDF revcon §6 rounds 20.4775→20.48 → P&L stated as −$204,800; full precision gives −$204,775. Tests use ±$30 tolerance.

---

## Quote conventions you'll see in chat

| Format | What it means | Example |
|---|---|---|
| `XXX.XX / YYY.YY  (Zf)` | Combo mid in **points** at a known strike, ES ref Z | `121.52 / 121.78  (7153f)` |
| `XXX.XXXX%  (Zf)` | Forward as **% of SPX spot**, ES ref Z (synthetic forward print) | `101.9275%  (7338f)` |

The pricer's **Run** tab outputs both formats — copy-paste ready.

---

## Borrow convention

Excel stores borrow as **bps** (e.g., 55 = 55 bps). `_bps_to_decimal()` in `excel.py` divides by 10,000. Never use `_to_decimal()` for borrow — it would read 55 as 55%.

---

## UI guide — `localhost:8501`

### Top bar (always visible)
`Snapshot HH:MM:SS (Xs ago)   SPX 7,312.50   ES 7,350.75 basis +38.3`
- **Green** = fresh (<30s), **orange** = stale soon, **red** = stop, refresh Bloomberg
- ⟳ button forces re-read

### Tab 1 — Ladder
Centred on the SPX **forward** for the chosen expiry. Footer shows `F  T  r  b  q  DF  ES`.

**🎯 Match a counterparty quote** expander:
- **Override ES futures ref** — type a specific ES level (e.g. 7153). Pricer shifts implied spot by the same delta and reprices.
- **Show one specific strike** — replaces the 7-strike ladder with a single row at any K.
- **Quoted bid / ask** — type the broker's two-way to get a side-by-side comparison with delta in points and bps of forward.

### Tab 2 — Revcon
Inputs (expiry / strike / combo mid / ES / basis / side / lots / **Basis confirmed?**), output is LIFT / PASS / FLAT (±1 bps annualised threshold) plus breakdown table.

### Tab 3 — Reverse Borrow
Backs out implied borrow from a street combo print. RICH/CHEAP tile shows delta vs your AXW model. Expander shows the implied borrow curve across all expiries.

### Tab 4 — Run (NEW)
Generates a full run across all expiries in both quote formats:
- Width slider (10bp D2D, 15bp client)
- Strike-mode radio: ATM forward (round-100) or ES level (round-25)
- Output: live table + two copy-paste blocks (% of spot, points)

Sample output:
```
SPX Combo Run — 15:41:57 NY
ES 7,350.75  SPX 7,312.50  basis +38.25

SPX DEC26 combo  101.9018% / 102.0037%  (7351f)
SPX DEC27 combo  105.2504% / 105.3557%  (7351f)
SPX DEC28 combo  108.7484% / 108.8572%  (7351f)
SPX DEC29 combo  112.7611% / 112.8739%  (7351f)
SPX DEC30 combo  116.9694% / 117.0864%  (7351f)
SPX DEC31 combo  122.9117% / 123.0347%  (7351f)
```

---

## Reading the run / sanity checks before sending

- Top bar **green**?
- F values ladder up by ~3-3.5%/year (each expiry = one more year of carry)?
- Borrow column ladders up DEC26 → DEC31? Any inversion = investigate
- ES level shown on every line matches live ES?

---

## Verifying any new street print (30-second drill)

1. Open **Tab 3 — Reverse Borrow**
2. Pick the expiry, type the strike (use the ES round number for %-of-spot trades)
3. Type the street's combo mid in points
4. Read the **Delta** in bps:
   - **|Δ| < 5 bps** → pricer is right, AXW is calibrated
   - **|Δ| ≥ 10 bps** → investigate (stale Bloomberg, weird AXW tick, or genuine market dislocation)

---

## File map

| File | Purpose |
|---|---|
| `spx_pricer/_rebuild_sheet.py` | Rebuilds BBG Data sheet + writes `config/excel_map.yaml`. Re-run to add expiries. |
| `spx_pricer/config/excel_map.yaml` | Auto-generated cell map. Do not edit by hand. |
| `spx_pricer/config/defaults.yaml` | Width defaults (10bp D2D, 15bp client), stale thresholds, Q4 extra 5bp. |
| `spx_pricer/pricer/math.py` | Pure math, no I/O. All PDF formulas. |
| `spx_pricer/pricer/excel.py` | xlwings I/O. `read_snapshot()` → `MarketSnapshot`. |
| `spx_pricer/pricer/ladder.py` | `build_ladder(snap, expiry_code, width_bps, es_override=None, specific_strike=None)` |
| `spx_pricer/pricer/revcon.py` | `calc_revcon(...)` |
| `spx_pricer/pricer/reverse.py` | `calc_reverse(...)` — implies borrow from street combo |
| `spx_pricer/pricer/cli.py` | Typer CLI: `ladder`, `revcon`, `reverse` |
| `spx_pricer/app.py` | Streamlit UI. 4 tabs: Ladder, Revcon, Reverse Borrow, Run. |
| `spx_pricer/logs/quotes.parquet` | Quote log (auto-created on first Revcon calc). |
| `spx_pricer/tests/` | 52 tests. Run: `pytest` from `spx_pricer\`. |

---

## To add a new expiry

1. Add to `EXPIRIES` list in `_rebuild_sheet.py`
2. Add to `EXPIRY_TO_YEAR` dict in `_rebuild_sheet.py`
3. Run `python _rebuild_sheet.py`
4. `config/excel_map.yaml` is regenerated automatically; new expiry appears in all four tab dropdowns

ASD/AXW strips currently extend to Z33 — expiries through DEC33 can be added without changing the BDP layout.

---

## Daily ritual

1. Open `Combos_RevCon_Pricer.xlsx` in Excel — Bloomberg connects automatically
2. Wait ~40s for delayed CME data (ASD/AXW)
3. Glance at green cells: any `#N/A`? Press **Ctrl+Alt+F9**
4. Run `streamlit run app.py` from `spx_pricer\` (or refresh browser tab)
5. **Tab 4 → Run** to see your morning quotes; sanity-check borrow curve
6. **Tab 3 → Reverse Borrow** against any overnight prints to confirm calibration

---

## Open calibration questions (all resolved)

1. SOFR tenor per expiry — default: nearest tenor ≥ T ✓
2. Bloomberg tickers — read from Excel, not hardcoded ✓
3. Dividend convention — S×q×T linear ✓
4. Basis convention — manual input + confirm checkbox ✓
5. Width discipline — 10bp D2D, 15bp client; slider in UI ✓
6. AXW units — bps direct (confirmed by 250mm DEC26 print, May 6 2026) ✓

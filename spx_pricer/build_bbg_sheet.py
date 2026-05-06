"""
Build the Bloomberg formula sheet inside Combos_RevCon_Pricer.xlsx.

Run once with Excel open:
    python build_bbg_sheet.py

Creates a sheet called "BBG Data" with all live BDP formulas,
then writes config/excel_map.yaml automatically.

Tickers that need verification with Athan (PDF §8 Q2) are flagged
in column D of the sheet.
"""

from __future__ import annotations
import time, sys
from pathlib import Path
import xlwings as xw
import yaml

WB_PATH   = r"C:\Users\sshooter\US Pricers\Combos_RevCon_Pricer.xlsx"
SHEET_NAME = "BBG Data"
CONFIG_DIR  = Path(__file__).parent / "config"
CONFIG_PATH = CONFIG_DIR / "excel_map.yaml"

# ── Expiry schedule ──────────────────────────────────────────────────────────
EXPIRIES = [
    ("DEC26", "2026-12-18"),
    ("DEC27", "2027-12-19"),
    ("DEC28", "2028-12-15"),
    ("DEC29", "2029-12-21"),
    ("DEC30", "2030-12-20"),
]

# ── SOFR tickers (Bloomberg) ─────────────────────────────────────────────────
# Short-end: CME Term SOFR.  Long-end: USD SOFR OIS swap mid (BGN).
# Verify with Athan per PDF §8 Q1 — "biggest flagged unknown".
SOFR_ROWS = [
    ("Overnight",  "SOFRRATE Index",        0.0027, "Fed-published daily SOFR fixing"),
    ("1M",         "TSFR1M Index",           0.083,  "CME Term SOFR 1M"),
    ("3M",         "TSFR3M Index",           0.25,   "CME Term SOFR 3M"),
    ("6M",         "TSFR6M Index",           0.5,    "CME Term SOFR 6M"),
    ("1Y",         "USOSFR1Y BGN Curncy",    1.0,    "SOFR OIS 1Y mid"),
    ("2Y",         "USOSFR2Y BGN Curncy",    2.0,    "SOFR OIS 2Y mid"),
    ("3Y",         "USOSFR3Y BGN Curncy",    3.0,    "SOFR OIS 3Y mid"),
    ("5Y",         "USSW5 Curncy",           5.0,    "USD SOFR swap 5Y"),
    ("10Y",        "USSW10 Curncy",          10.0,   "USD SOFR swap 10Y"),
]

# ── Dividend strip tickers ───────────────────────────────────────────────────
# ASDZ futures = S&P 500 Annual Dividend Index futures (ICE, Dec expiry)
# Stored as dollar amounts.  Yield = strip_value / SPX_spot.
ASDZ_YEARS = [2026, 2027, 2028, 2029, 2030, 2031, 2032, 2033]

# ── Borrow tickers ───────────────────────────────────────────────────────────
# AIR futures = S&P 500 Implied Repo Rate futures (CBOE)
# Ticker format: AIR<month_code><year> Index, e.g. AIRZ26 Index for Dec 2026.
# ⚠ VERIFY WITH ATHAN — these may differ from your BBG setup (PDF §8 Q2).
# If not found, borrow cells will show #N/A and you can override manually.
BORROW_TICKERS = {
    "DEC26": "AIRZ26 Index",
    "DEC27": "AIRZ27 Index",
    "DEC28": "AIRZ28 Index",
    "DEC29": "AIRZ29 Index",
    "DEC30": "AIRZ30 Index",
}


# ── Helpers ───────────────────────────────────────────────────────────────────

def col(n: int) -> str:
    """1-indexed column number → letter (A, B, … Z, AA, …)"""
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def write_header(sh, row: int, text: str):
    cell = sh.range(f"A{row}")
    cell.value = text
    cell.font.bold = True
    cell.font.color = (0x1F, 0x49, 0x7D)   # dark blue


def write_row(sh, row: int, label: str, formula: str, note: str = "", extra_label: str = ""):
    sh.range(f"A{row}").value = label
    sh.range(f"B{row}").formula = formula
    if note:
        sh.range(f"D{row}").value = note
    if extra_label:
        sh.range(f"C{row}").value = extra_label


# ── Main ──────────────────────────────────────────────────────────────────────

def build_sheet():
    print(f"Attaching to: {WB_PATH}")
    try:
        wb  = xw.Book(WB_PATH)
        app = wb.app
        app.visible = True
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)

    # Delete old BBG Data sheet if it exists
    for sh in wb.sheets:
        if sh.name == SHEET_NAME:
            sh.delete()

    sh = wb.sheets.add(name=SHEET_NAME, after=wb.sheets[0])
    sh.activate()

    # Column widths
    sh.range("A:A").column_width = 20
    sh.range("B:B").column_width = 28
    sh.range("C:C").column_width = 14
    sh.range("D:D").column_width = 50

    R = 1   # current row pointer

    # ── Section 1: Spot & Futures ────────────────────────────────────────────
    write_header(sh, R, "SPOT & FUTURES"); R += 1
    sh.range(f"A{R}").value = "Label";  sh.range(f"B{R}").value = "Value"
    sh.range(f"A{R}:B{R}").font.bold = True; R += 1

    SPX_ROW = R
    write_row(sh, R, "SPX Spot",    '=BDP("SPX Index","PX_LAST")',  "SPX cash index"); R += 1

    ES_ROW = R
    write_row(sh, R, "ES1 Future",  '=BDP("ES1 Index","PX_LAST")',   "ES front-month continuous"); R += 1

    BASIS_ROW = R
    sh.range(f"A{R}").value = "ES Basis (ES-SPX)"
    sh.range(f"B{R}").formula = f"=B{ES_ROW}-B{SPX_ROW}"
    sh.range(f"D{R}").value  = "Auto-computed: ES minus SPX. Override manually if needed."
    R += 2

    # ── Section 2: SOFR Curve ────────────────────────────────────────────────
    write_header(sh, R, "SOFR CURVE  (verify tickers with Athan — PDF §8 Q1)"); R += 1
    sh.range(f"A{R}").value = "Tenor"; sh.range(f"B{R}").value = "Rate (%)"
    sh.range(f"C{R}").value = "T (yrs)"; sh.range(f"D{R}").value = "Bloomberg ticker / note"
    sh.range(f"A{R}:D{R}").font.bold = True; R += 1

    SOFR_START_ROW = R
    for label, ticker, tenor_yrs, note in SOFR_ROWS:
        write_row(sh, R, label, f'=BDP("{ticker}","PX_LAST")', note, str(tenor_yrs))
        R += 1
    SOFR_END_ROW = R - 1
    R += 1

    # ── Section 3: ASDZ Dividend Strip ───────────────────────────────────────
    write_header(sh, R, "DIVIDEND STRIP  (ASDZ — S&P 500 Annual Dividend Futures)"); R += 1
    sh.range(f"A{R}").value = "Year"; sh.range(f"B{R}").value = "$ Strip (dollar amt)"
    sh.range(f"C{R}").value = "Approx Yield %"; sh.range(f"D{R}").value = "Bloomberg ticker"
    sh.range(f"A{R}:D{R}").font.bold = True; R += 1

    ASDZ_START_ROW = R
    asdz_cells: dict[int, str] = {}       # year → "B{row}"
    for yr in ASDZ_YEARS:
        ticker = f"ASD{yr}Z Index"
        approx_yield_formula = f"=IF(B{SPX_ROW}>0, B{R}/B{SPX_ROW}*100, \"\")"
        write_row(sh, R, str(yr), f'=BDP("{ticker}","PX_LAST")', ticker, "")
        sh.range(f"C{R}").formula = approx_yield_formula
        asdz_cells[yr] = f"B{R}"
        R += 1
    ASDZ_END_ROW = R - 1
    R += 1

    # ── Section 4: Div Yield Per Expiry (aggregated) ─────────────────────────
    write_header(sh, R, "DIV YIELD PER EXPIRY  (aggregated from ASDZ strip)"); R += 1
    sh.range(f"A{R}").value = "Expiry"; sh.range(f"B{R}").value = "Yield %"
    sh.range(f"C{R}").value = "Settle date"; sh.range(f"D{R}").value = "Method: mean(ASDZ years covered) / SPX"
    sh.range(f"A{R}:D{R}").font.bold = True; R += 1

    DIV_YIELD_START_ROW = R
    div_yield_cells: dict[str, str] = {}   # expiry_code → "B{row}"
    for code, settle_str in EXPIRIES:
        settle_yr = int(settle_str[:4])
        # Include all years from current year through settle year
        from datetime import date
        current_yr = date.today().year
        relevant_rows = [asdz_cells[y] for y in range(current_yr, settle_yr + 1) if y in asdz_cells]
        n = len(relevant_rows)
        if relevant_rows and n > 0:
            sum_formula = "+".join(relevant_rows)
            # yield = sum(strips) / spot / n_years  (simple average annual yield)
            formula = f"=IF(B{SPX_ROW}>0, ({sum_formula}) / B{SPX_ROW} / {n} * 100, \"\")"
        else:
            formula = f'=IF(B{SPX_ROW}>0, {list(asdz_cells.values())[-1]}/B{SPX_ROW}*100, "")'
        sh.range(f"A{R}").value = code
        sh.range(f"B{R}").formula = formula
        sh.range(f"C{R}").value = settle_str
        div_yield_cells[code] = f"B{R}"
        R += 1
    DIV_YIELD_END_ROW = R - 1
    R += 1

    # ── Section 5: Borrow Curve ───────────────────────────────────────────────
    write_header(sh, R, "BORROW CURVE  ⚠ VERIFY TICKERS WITH ATHAN — PDF §8 Q2"); R += 1
    sh.range(f"A{R}").value = "Expiry"; sh.range(f"B{R}").value = "Borrow (%)"
    sh.range(f"D{R}").value = "Bloomberg ticker — confirm with Athan before going live"
    sh.range(f"A{R}:D{R}").font.bold = True; R += 1

    BORROW_START_ROW = R
    borrow_cells: dict[str, str] = {}
    for code, settle_str in EXPIRIES:
        ticker = BORROW_TICKERS.get(code, f"⚠ NO TICKER FOR {code}")
        write_row(sh, R, code, f'=BDP("{ticker}","PX_LAST")',
                  f"{ticker}  ← verify with Athan", settle_str)
        borrow_cells[code] = f"B{R}"
        R += 1
    BORROW_END_ROW = R - 1
    R += 2

    # ── Instructions cell ─────────────────────────────────────────────────────
    sh.range(f"A{R}").value = (
        "INSTRUCTIONS: Bloomberg formulas refresh automatically when connected. "
        "If cells show #N/A, press Ctrl+Alt+F9 to force recalc. "
        "Borrow tickers (Section 5) must be verified with Athan before trading."
    )
    sh.range(f"A{R}").font.italic = True
    sh.range(f"A{R}:D{R}").merge()

    # ── Save workbook ─────────────────────────────────────────────────────────
    wb.save()
    print(f"Sheet '{SHEET_NAME}' written and saved.")

    # ── Force recalc & read back spot to confirm Bloomberg is live ─────────────
    print("Waiting 6s for Bloomberg to populate cells…")
    time.sleep(6)
    try:
        app.calculate()
        if hasattr(app.api, "CalculateFull"):
            app.api.CalculateFull()
        time.sleep(2)
    except Exception:
        pass

    spot_val = sh.range(f"B{SPX_ROW}").value
    es_val   = sh.range(f"B{ES_ROW}").value
    sofr1y   = sh.range(f"B{SOFR_START_ROW + 4}").value   # row index 4 = 1Y tenor
    asd26    = sh.range(f"{list(asdz_cells.values())[0]}").value

    print()
    print("  Bloomberg live check:")
    print(f"    SPX spot   : {spot_val!r}")
    print(f"    ES1 future : {es_val!r}")
    print(f"    SOFR 1Y    : {sofr1y!r}")
    print(f"    ASD2026Z   : {asd26!r}")
    print()

    bloomberg_live = (
        isinstance(spot_val, (int, float))
        and isinstance(es_val, (int, float))
    )
    if bloomberg_live:
        print("  ✓ Bloomberg is live — values look good.")
    else:
        print("  ⚠ Some cells still show #N/A or None.")
        print("    In Excel: press Ctrl+Alt+F9 to force a Bloomberg recalc,")
        print("    then re-run `python build_bbg_sheet.py` OR proceed to discover.py.")

    # ── Write excel_map.yaml ──────────────────────────────────────────────────
    CONFIG_DIR.mkdir(exist_ok=True)

    sofr_range = f"B{SOFR_START_ROW}:B{SOFR_END_ROW}"
    tenors     = [t for _, _, t, _ in SOFR_ROWS]

    # Div strip range (raw dollar values)
    asdz_range = f"B{ASDZ_START_ROW}:B{ASDZ_END_ROW}"

    # Div yield per expiry — individual cells
    # Store as a range from DIV_YIELD_START_ROW to DIV_YIELD_END_ROW
    div_yield_range = f"B{DIV_YIELD_START_ROW}:B{DIV_YIELD_END_ROW}"
    div_labels      = [c for c, _ in EXPIRIES]

    borrow_range = f"B{BORROW_START_ROW}:B{BORROW_END_ROW}"

    cfg = {
        "workbook": WB_PATH,
        "platform": "windows",
        "spot": {
            "sheet": SHEET_NAME,
            "cell":  f"B{SPX_ROW}",
        },
        "sofr_curve": {
            "sheet":        SHEET_NAME,
            "range":        sofr_range,
            "tenors_years": tenors,
        },
        "div_strip": {
            "sheet":         SHEET_NAME,
            "range":         div_yield_range,
            "expiry_labels": div_labels,
            "_note": (
                "Stores aggregated div yield (%) per expiry, not raw dollar strip. "
                "Raw ASDZ dollar values are in rows "
                f"{ASDZ_START_ROW}-{ASDZ_END_ROW} columns A-B for reference."
            ),
        },
        "borrow_curve": {
            "sheet":         SHEET_NAME,
            "range":         borrow_range,
            "expiry_labels": div_labels,
            "_note": "Verify AIR future tickers with Athan — PDF §8 Q2",
        },
        "es_future": {
            "sheet": SHEET_NAME,
            "cell":  f"B{ES_ROW}",
        },
        "es_basis": {
            "sheet": SHEET_NAME,
            "cell":  f"B{BASIS_ROW}",
        },
        "expiries": [
            {"code": code, "date": settle}
            for code, settle in EXPIRIES
        ],
    }

    with open(CONFIG_PATH, "w") as f:
        yaml.dump(cfg, f, default_flow_style=False, sort_keys=False)

    print(f"  Wrote: {CONFIG_PATH}")
    print()
    print("  Next steps:")
    print("    1. Check Bloomberg cells in Excel (Ctrl+Alt+F9 if any #N/A).")
    print("    2. Run:  python -m pricer.snapshot")
    print("       to verify the live snapshot matches your Bloomberg data.")
    print("    3. If borrow tickers show #N/A — ask Athan for the correct")
    print("       AIR ticker format, then update column B rows")
    print(f"       {BORROW_START_ROW}–{BORROW_END_ROW} in the '{SHEET_NAME}' sheet.")
    print()
    print("  Once snapshot is confirmed:")
    print("    streamlit run app.py")


if __name__ == "__main__":
    build_sheet()

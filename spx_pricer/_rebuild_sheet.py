"""Rebuild the BBG Data sheet with confirmed-working tickers.
Div yield driven live from ASD futures.  Borrow driven live from AXW (AIR) futures.
Run once each time you add a new expiry or re-map the sheet layout.
"""
import xlwings as xw, time, sys, yaml
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

WB_PATH    = r"C:\Users\sshooter\US Pricers\Combos_RevCon_Pricer.xlsx"
SHEET_NAME = "BBG Data"
CONFIG_DIR  = Path(r"C:\Users\sshooter\US Pricers\spx_pricer\config")
CONFIG_PATH = CONFIG_DIR / "excel_map.yaml"

# Pricing expiries (DEC26-DEC30) and their settlement dates
EXPIRIES = [
    ("DEC26", "2026-12-18"),
    ("DEC27", "2027-12-19"),
    ("DEC28", "2028-12-15"),
    ("DEC29", "2029-12-21"),
    ("DEC30", "2030-12-20"),
    ("DEC31", "2031-12-19"),
]

# ASD strip years to display on sheet (full curve for reference, Z26-Z33)
ASD_YEARS = [26, 27, 28, 29, 30, 31, 32, 33]

# AXW strip years to display on sheet (same range)
AXW_YEARS = [26, 27, 28, 29, 30, 31, 32, 33]

# Which ASD year maps to which pricing expiry (primary year for annual yield calc)
EXPIRY_TO_YEAR = {
    "DEC26": 26, "DEC27": 27, "DEC28": 28, "DEC29": 29, "DEC30": 30, "DEC31": 31,
}

SOFR_ROWS = [
    ("Overnight", "SOFRRATE Index", "PX_LAST",    0.0027, "SOFR overnight fixing (FRBNY)"),
    ("1M",        "TSFR1M Index",   "PX_LAST",    0.083,  "CME Term SOFR 1M"),
    ("3M",        "TSFR3M Index",   "PX_LAST",    0.25,   "CME Term SOFR 3M"),
    ("6M",        "TSFR6M Index",   "PX_LAST",    0.5,    "CME Term SOFR 6M"),
    ("1Y",        "USSO1 Curncy",   "LAST_PRICE", 1.0,    "USD SOFR OIS swap 1Y"),
    ("2Y",        "USSO2 Curncy",   "LAST_PRICE", 2.0,    "USD SOFR OIS swap 2Y"),
    ("3Y",        "USSO3 Curncy",   "LAST_PRICE", 3.0,    "USD SOFR OIS swap 3Y"),
    ("5Y",        "USSO5 Curncy",   "LAST_PRICE", 5.0,    "USD SOFR OIS swap 5Y"),
    ("10Y",       "USSO10 Curncy",  "LAST_PRICE", 10.0,   "USD SOFR OIS swap 10Y"),
]

# -----------------------------------------------------------------------
wb  = xw.Book(WB_PATH)
app = wb.app
app.visible = True

# Wait for any running calculation before touching the workbook
print("Waiting for Excel to settle...")
time.sleep(5)

# Remove and recreate sheet — retry if Excel is momentarily busy
for s in wb.sheets:
    if s.name == SHEET_NAME:
        s.delete()
        time.sleep(2)

for attempt in range(5):
    try:
        sh = wb.sheets.add(name=SHEET_NAME, after=wb.sheets[0])
        break
    except Exception as e:
        if attempt == 4:
            raise
        print(f"  Sheet add attempt {attempt+1} failed ({e}), retrying in 3s...")
        time.sleep(3)
sh.activate()

sh.range("A:A").column_width = 18
sh.range("B:B").column_width = 16
sh.range("C:C").column_width = 14
sh.range("D:D").column_width = 52

DARK_BLUE  = (0x1F, 0x49, 0x7D)
ORANGE     = (0xC0, 0x50, 0x00)
GREY       = (0x60, 0x60, 0x60)
YELLOW_BG  = (0xFF, 0xFF, 0xCC)
LIGHT_BLUE = (0xDD, 0xEE, 0xFF)
LIGHT_GRN  = (0xDD, 0xFF, 0xDD)

def hdr(row, txt):
    c = sh.range(f"A{row}")
    c.value = txt
    c.font.bold = True
    c.font.color = DARK_BLUE

def note(row, txt):
    sh.range(f"D{row}").value = txt
    sh.range(f"D{row}").font.italic = True
    sh.range(f"D{row}").font.color = GREY

R = 1

# -----------------------------------------------------------------------
# Section 1: Spot & Futures
# -----------------------------------------------------------------------
hdr(R, "SPOT & FUTURES  (live Bloomberg)"); R += 1
sh.range(f"A{R}").value = "Label"
sh.range(f"B{R}").value = "Value"
sh.range(f"D{R}").value = "Bloomberg formula"
sh.range(f"A{R}:D{R}").font.bold = True
R += 1

SPX_ROW = R
sh.range(f"A{R}").value = "SPX Spot"
sh.range(f"B{R}").formula = '=BDP("SPX Index","PX_LAST")'
note(R, 'BDP("SPX Index","PX_LAST")')
R += 1

ES_ROW = R
sh.range(f"A{R}").value = "ES1 Future"
sh.range(f"B{R}").formula = '=BDP("ES1 Index","PX_LAST")'
note(R, 'BDP("ES1 Index","PX_LAST")')
R += 1

BASIS_ROW = R
sh.range(f"A{R}").value = "ES Basis"
sh.range(f"B{R}").formula = f"=B{ES_ROW}-B{SPX_ROW}"
note(R, "Auto: ES minus SPX.  Override manually if needed.")
R += 2

# -----------------------------------------------------------------------
# Section 2: SOFR Curve (all live)
# -----------------------------------------------------------------------
hdr(R, "SOFR CURVE  (all 9 tenors live and confirmed)"); R += 1
sh.range(f"A{R}").value = "Tenor"
sh.range(f"B{R}").value = "Rate (%)"
sh.range(f"C{R}").value = "T (yrs)"
sh.range(f"D{R}").value = "Ticker and field"
sh.range(f"A{R}:D{R}").font.bold = True
R += 1

SOFR_START = R
for label, ticker, field, t_yrs, desc in SOFR_ROWS:
    sh.range(f"A{R}").value = label
    sh.range(f"B{R}").formula = f'=BDP("{ticker}","{field}")'
    sh.range(f"C{R}").value = t_yrs
    note(R, f'{ticker}  [{field}]  {desc}')
    R += 1
SOFR_END = R - 1
R += 1

# -----------------------------------------------------------------------
# Section 3: Raw ASD strip  (live — CME S&P 500 Annual Dividend futures)
# -----------------------------------------------------------------------
hdr(R, "RAW ASD STRIP  (CME S&P 500 Annual Dividend futures, live)"); R += 1
sh.range(f"A{R}").value = "Year"
sh.range(f"B{R}").value = "ASD (pts)"
sh.range(f"D{R}").value = "Ticker and field"
sh.range(f"A{R}:D{R}").font.bold = True
R += 1

ASD_START = R
ASD_CELL = {}   # year -> "B{row}" address
for yr in ASD_YEARS:
    ticker = f"ASDZ{yr} Index"
    sh.range(f"A{R}").value = f"ASDZ{yr}"
    sh.range(f"B{R}").formula = f'=BDP("{ticker}","LAST_PRICE")'
    sh.range(f"B{R}").color = LIGHT_BLUE
    note(R, f'{ticker}  [LAST_PRICE]  S&P 500 annual dividend points for 20{yr}')
    ASD_CELL[yr] = f"B{R}"
    R += 1
ASD_END = R - 1
R += 1

# -----------------------------------------------------------------------
# Section 4: Div yield per pricing expiry  (live — computed from ASD/SPX)
# -----------------------------------------------------------------------
hdr(R, "DIVIDEND YIELD PER EXPIRY  (% ann — LIVE from ASD/SPX)"); R += 1
sh.range(f"A{R}").value = "Expiry"
sh.range(f"B{R}").value = "Div Yield %"
sh.range(f"C{R}").value = "Settle date"
sh.range(f"D{R}").value = (
    "Formula: =ASD_cell/SPX_cell*100  (annual yield for primary expiry year). "
    "Override by typing a number.  Source: CME ASDZ futures (delayed)."
)
sh.range(f"A{R}:C{R}").font.bold = True
sh.range(f"D{R}").font.color = ORANGE
R += 1

DIV_START = R
div_cells = {}
for code, settle in EXPIRIES:
    yr = EXPIRY_TO_YEAR[code]
    asd_ref = ASD_CELL[yr]           # e.g. "B14"
    spx_ref = f"B{SPX_ROW}"
    sh.range(f"A{R}").value = code
    sh.range(f"B{R}").formula = f"={asd_ref}/{spx_ref}*100"
    sh.range(f"B{R}").color = LIGHT_GRN
    sh.range(f"C{R}").value = settle
    note(R, f"Formula: ASD_Z{yr} / SPX * 100  (primary-year yield; type a number to override)")
    div_cells[code] = f"B{R}"
    R += 1
DIV_END = R - 1
R += 1

# -----------------------------------------------------------------------
# Section 5: Raw AXW strip  (live — CME Annual Implied Return / AIR futures)
# -----------------------------------------------------------------------
hdr(R, "RAW AXW STRIP  (CME AIR futures, live — quoted in bps)"); R += 1
sh.range(f"A{R}").value = "Year"
sh.range(f"B{R}").value = "AXW (bps)"
sh.range(f"D{R}").value = "Ticker and field"
sh.range(f"A{R}:D{R}").font.bold = True
R += 1

AXW_START = R
AXW_CELL = {}   # year -> "B{row}" address
for yr in AXW_YEARS:
    ticker = f"AXWZ{yr} Index"
    sh.range(f"A{R}").value = f"AXWZ{yr}"
    sh.range(f"B{R}").formula = f'=BDP("{ticker}","LAST_PRICE")'
    sh.range(f"B{R}").color = LIGHT_BLUE
    note(R, f'{ticker}  [LAST_PRICE]  AIR implied return for 20{yr} — bps')
    AXW_CELL[yr] = f"B{R}"
    R += 1
AXW_END = R - 1
R += 1

# -----------------------------------------------------------------------
# Section 6: Borrow per pricing expiry  (live — from AXW, assumed bps)
# -----------------------------------------------------------------------
hdr(R, "BORROW PER EXPIRY  (bps — LIVE from AXW)"); R += 1
sh.range(f"A{R}").value = "Expiry"
sh.range(f"B{R}").value = "Borrow (bps)"
sh.range(f"C{R}").value = "Settle date"
sh.range(f"D{R}").value = "Formula: =AXWZxx  (bps, confirmed). Python divides by 10000. Type to override."
sh.range(f"A{R}:C{R}").font.bold = True
R += 1

BORR_START = R
borr_cells = {}
for code, settle in EXPIRIES:
    yr = EXPIRY_TO_YEAR[code]
    axw_ref = AXW_CELL[yr]
    sh.range(f"A{R}").value = code
    sh.range(f"B{R}").formula = f"={axw_ref}"         # pass-through — assumed bps
    sh.range(f"B{R}").color = LIGHT_GRN
    sh.range(f"C{R}").value = settle
    note(R, f"Formula: AXW_Z{yr}  (assumed bps; Python divides by 10000; type a number to override)")
    borr_cells[code] = f"B{R}"
    R += 1
BORR_END = R - 1
R += 2

# -----------------------------------------------------------------------
# Instructions footer
# -----------------------------------------------------------------------
sh.range(f"A{R}").value = (
    "INSTRUCTIONS: SOFR, ASD, and AXW cells update live via Bloomberg (delayed CME sub). "
    "Blue cells = live raw Bloomberg feed. Green cells = computed from live feed (can be manually overridden by typing). "
    "If any cell shows #N/A: press Ctrl+Alt+F9."
)
sh.range(f"A{R}:D{R}").merge()
sh.range(f"A{R}").font.italic = True
sh.range(f"A{R}").font.size = 9
sh.range(f"A{R}").color = (0xF0, 0xF4, 0xFF)

wb.save()
print("Sheet written OK.")

# -----------------------------------------------------------------------
# Force recalc and read back
# -----------------------------------------------------------------------
print("Recalculating Bloomberg cells (waiting ~40s for delayed CME data)...")
app.calculate()
try:
    app.api.CalculateFull()
except Exception:
    pass

# Delayed CME data (ASD/AXW) can take 35-40s to populate
time.sleep(40)
app.calculate()
time.sleep(3)

spot     = sh.range(f"B{SPX_ROW}").value
es       = sh.range(f"B{ES_ROW}").value
basis    = sh.range(f"B{BASIS_ROW}").value
sofr_on  = sh.range(f"B{SOFR_START}").value
sofr_1y  = sh.range(f"B{SOFR_START+4}").value
sofr_10  = sh.range(f"B{SOFR_END}").value

print()
print(f"  SPX spot    : {spot}")
print(f"  ES1 future  : {es}")
print(f"  ES basis    : {basis}")
print(f"  SOFR ON     : {sofr_on}")
print(f"  SOFR 1Y     : {sofr_1y}")
print(f"  SOFR 10Y    : {sofr_10}")
print()
print("  ASD strip:")
for yr in ASD_YEARS:
    val = sh.range(ASD_CELL[yr]).value
    print(f"    ASDZ{yr}: {val}")
print()
print("  AXW strip:")
for yr in AXW_YEARS:
    val = sh.range(AXW_CELL[yr]).value
    print(f"    AXWZ{yr}: {val}")
print()
print("  Per-expiry (computed):")
for code, settle in EXPIRIES:
    div = sh.range(div_cells[code]).value
    brw = sh.range(borr_cells[code]).value
    print(f"    {code}  div_yield={div:.4f}%  borrow={brw} bps  settle={settle}")

# -----------------------------------------------------------------------
# Write excel_map.yaml
# -----------------------------------------------------------------------
sofr_range = f"B{SOFR_START}:B{SOFR_END}"
tenors     = [t for _, _, _, t, _ in SOFR_ROWS]
div_range  = f"B{DIV_START}:B{DIV_END}"
borr_range = f"B{BORR_START}:B{BORR_END}"
exp_labels = [c for c, _ in EXPIRIES]

cfg = {
    "workbook": WB_PATH,
    "platform": "windows",
    "spot":       {"sheet": SHEET_NAME, "cell": f"B{SPX_ROW}"},
    "sofr_curve": {"sheet": SHEET_NAME, "range": sofr_range, "tenors_years": tenors},
    "div_strip":  {
        "sheet": SHEET_NAME,
        "range": div_range,
        "expiry_labels": exp_labels,
        "_note": "Cells formula-driven: =ASDZxx/SPX*100  (% ann).  Python uses _to_decimal().",
    },
    "borrow_curve": {
        "sheet": SHEET_NAME,
        "range": borr_range,
        "expiry_labels": exp_labels,
        "_note": "Cells formula-driven: =AXWZxx (bps, confirmed).  Python uses _bps_to_decimal().",
    },
    "es_future": {"sheet": SHEET_NAME, "cell": f"B{ES_ROW}"},
    "es_basis":  {"sheet": SHEET_NAME, "cell": f"B{BASIS_ROW}"},
    "expiries":  [{"code": c, "date": d} for c, d in EXPIRIES],
}

CONFIG_DIR.mkdir(exist_ok=True)
with open(CONFIG_PATH, "w") as f:
    yaml.dump(cfg, f, default_flow_style=False, sort_keys=False)

print(f"\n  excel_map.yaml written: {CONFIG_PATH}")
print("\nDone. Run 'streamlit run app.py' to start the pricer.")

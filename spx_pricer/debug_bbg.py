"""
Bloomberg / xlwings diagnostic.
Run: python debug_bbg.py
Tells you exactly what xlwings is reading, whether Bloomberg is still
alive, and whether a forced recalc changes anything.
"""
import time
import datetime
import xlwings as xw

WB_PATH = r"C:\Users\sshooter\US Pricers\Combos_RevCon_Pricer.xlsx"
SHEET   = "BBG Data"

# Cells we care about
SPOT_CELL   = "B3"
ES_CELL     = "B4"
BASIS_CELL  = "B5"
# Raw ASD strip (rows vary — scan a window)
ASD_RANGE   = "B20:B27"
AXW_RANGE   = "B40:B47"


def hr(title=""):
    print(f"\n{'='*60}")
    if title:
        print(f"  {title}")
    print('='*60)


def read_key_cells(sh):
    rows = []
    for label, cell in [("SPX spot", SPOT_CELL), ("ES1", ES_CELL), ("ES basis", BASIS_CELL)]:
        val     = sh.range(cell).value
        formula = sh.range(cell).formula
        rows.append((label, cell, val, formula))
    return rows


def print_cells(rows):
    for label, cell, val, formula in rows:
        print(f"  {cell} ({label:12s}):  value={val!r:>18}   formula={formula!r}")


hr("OPEN EXCEL APPS")
print(f"Timestamp: {datetime.datetime.now()}\n")

if not xw.apps:
    print("ERROR: No Excel instance found. Open Combos_RevCon_Pricer.xlsx first.")
    raise SystemExit(1)

for i, app in enumerate(xw.apps):
    print(f"App[{i}]  pid={app.pid}  calc_mode={app.calculation}")
    for bk in app.books:
        print(f"          book: {bk.name}")

# ── Connect ──────────────────────────────────────────────────────────────────
try:
    wb  = xw.Book(WB_PATH)
    app = wb.app
    print(f"\nConnected: {wb.fullname}")
    print(f"Calc mode: {app.calculation}  (should be 'automatic')")
except Exception as exc:
    print(f"CONNECT FAILED: {exc}")
    raise SystemExit(1)

sh = wb.sheets[SHEET]

# ── Read BEFORE recalc ────────────────────────────────────────────────────────
hr("VALUES — BEFORE RECALC")
before = read_key_cells(sh)
print_cells(before)

# ── Formulas ─────────────────────────────────────────────────────────────────
hr("FORMULAS IN KEY CELLS")
for label, cell, val, formula in before:
    print(f"  {cell} ({label}):  {formula!r}")

# ── Try Bloomberg RunAll ──────────────────────────────────────────────────────
hr("BLOOMBERG REFRESH ATTEMPT")
for macro in ("RefreshAllStaticData", "BLP_Refresh", "Bloomberg_Refresh"):
    try:
        app.api.Run(macro)
        print(f"  app.api.Run('{macro}')  OK")
        break
    except Exception as e:
        print(f"  app.api.Run('{macro}')  FAILED: {e}")

# ── Force Excel recalc ────────────────────────────────────────────────────────
hr("FORCE EXCEL RECALC")
try:
    app.calculation = "automatic"
    app.calculate()
    print("  app.calculate()  ->  OK")
except Exception as e:
    print(f"  app.calculate()  ->  FAILED: {e}")

try:
    if hasattr(app.api, "CalculateFull"):
        app.api.CalculateFull()
        print("  CalculateFull()  ->  OK")
    else:
        print("  CalculateFull  ->  not available")
except Exception as e:
    print(f"  CalculateFull()  ->  FAILED: {e}")

print("\n  Waiting 3s for Bloomberg RTD to settle...")
time.sleep(3)

# ── Read AFTER recalc ─────────────────────────────────────────────────────────
hr("VALUES — AFTER RECALC (3s wait)")
after = read_key_cells(sh)
print_cells(after)

# ── Change detection ──────────────────────────────────────────────────────────
hr("CHANGE DETECTION")
any_change = False
for (lbl, cell, v_before, _), (_, _, v_after, _) in zip(before, after):
    changed = (v_before != v_after)
    any_change = any_change or changed
    status = "CHANGED" if changed else "FROZEN"
    print(f"  {cell} ({lbl}):  {v_before!r}  ->  {v_after!r}   [{status}]")

if not any_change:
    print("\n  *** ALL VALUES FROZEN — Bloomberg is NOT updating these cells ***")
    print("  Likely causes:")
    print("    1. Bloomberg Terminal lost its DDE/RTD connection to Excel")
    print("    2. Calc mode is Manual (Ctrl+Alt+F9 in Excel to force refresh)")
    print("    3. Bloomberg Add-In needs to be re-enabled (File > Options > Add-ins)")
    print("    4. The BDP tickers themselves are returning no new data (market closed?)")
else:
    print("\n  At least some values changed — Bloomberg IS pushing data.")

# ── ASD/AXW raw strip scan ────────────────────────────────────────────────────
hr("RAW ASD STRIP (B20:B27)")
for i, val in enumerate(sh.range(ASD_RANGE).value or []):
    formula = sh.range(f"B{20+i}").formula
    print(f"  B{20+i}:  value={val!r:>12}   formula={formula!r}")

hr("RAW AXW STRIP (B40:B47)")
for i, val in enumerate(sh.range(AXW_RANGE).value or []):
    formula = sh.range(f"B{40+i}").formula
    print(f"  B{40+i}:  value={val!r:>12}   formula={formula!r}")

# ── Consecutive-read canary ───────────────────────────────────────────────────
hr("CANARY: 3 reads × 2s apart — does spot ever move?")
reads = []
for i in range(3):
    app.calculate()
    time.sleep(2)
    v = sh.range(SPOT_CELL).value
    reads.append(v)
    print(f"  Read {i+1}: SPX spot = {v!r}")

if len(set(reads)) == 1:
    print(f"\n  STUCK: same value ({reads[0]}) across all 3 reads.")
    print("  Bloomberg RTD is not delivering fresh ticks to Excel.")
else:
    print(f"\n  LIVE: values are moving. Feed is working.")

hr("DONE")
print(f"Timestamp: {datetime.datetime.now()}")

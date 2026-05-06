"""
Phase 0 — Interactive Excel discovery.

Run this once with your Bloomberg workbook open:
    python discover.py

It will walk you through identifying each required input cell one at a time,
then write config/excel_map.yaml.
"""

from __future__ import annotations

import sys
from pathlib import Path
from datetime import date

# ---------------------------------------------------------------------------
# Guard: confirm workbook path up front
# ---------------------------------------------------------------------------

WORKBOOK_PATH = r"C:\Users\sshooter\US Pricers\Combos_RevCon_Pricer.xlsx"
CONFIG_DIR = Path(__file__).parent / "config"
CONFIG_PATH = CONFIG_DIR / "excel_map.yaml"


def _connect(path: str):
    import xlwings as xw
    print(f"\nConnecting to: {path}")
    try:
        wb = xw.Book(path)
        app = wb.app
    except Exception as exc:
        print(
            f"\nERROR: Cannot attach to '{path}'.\n"
            "Make sure the workbook is open in Excel "
            "and Bloomberg shows 'Connected' in the ribbon.\n"
            f"Detail: {exc}"
        )
        sys.exit(1)
    print("  Connected OK.")
    print("  Forcing recalculation...")
    try:
        app.calculate()
        if hasattr(app.api, "CalculateFull"):
            app.api.CalculateFull()
    except Exception:
        pass
    return wb, app


def _print_sheets(wb) -> list[str]:
    names = [sh.name for sh in wb.sheets]
    print("\nSheets in workbook:")
    for i, n in enumerate(names, 1):
        print(f"  {i}. {n}")
    return names


def _print_named_ranges(wb):
    try:
        nrs = wb.names
        if nrs:
            print("\nNamed ranges:")
            for nr in nrs:
                print(f"  {nr.name}  →  {nr.refers_to}")
    except Exception:
        pass


def _read_cell_safe(wb, sheet_name: str, cell: str):
    try:
        sh = wb.sheets[sheet_name]
        val = sh.range(cell).value
        formula = sh.range(cell).formula
        return val, formula
    except Exception as exc:
        return None, str(exc)


def _search_for_keyword(wb, keyword: str) -> list[tuple[str, str, object]]:
    """Scan all sheets for cells whose value or formula contains keyword."""
    hits = []
    kw = keyword.lower()
    for sh in wb.sheets:
        try:
            used = sh.used_range
            for cell in used:
                try:
                    formula = cell.formula or ""
                    val = cell.value
                    if kw in str(formula).lower() or kw in str(val).lower():
                        hits.append((sh.name, cell.address, val, formula))
                except Exception:
                    pass
        except Exception:
            pass
    return hits


def _ask_confirm(wb, prompt: str, sheet: str, cell: str) -> tuple[str, str]:
    """
    Show the user a candidate cell and ask: y / n / <SheetName!CellRef>
    Returns (confirmed_sheet, confirmed_cell).
    """
    val, formula = _read_cell_safe(wb, sheet, cell)
    print(f"\n  {prompt}")
    print(f"  Candidate: {sheet}!{cell}")
    print(f"    Formula : {formula!r}")
    print(f"    Value   : {val}")
    while True:
        ans = input("  Confirm? [y / n / Sheet!Cell]: ").strip()
        if ans.lower() == "y":
            return sheet, cell
        if ans.lower() == "n":
            ref = input("  Enter the correct cell (e.g. 'Sheet2!D7'): ").strip()
            if "!" in ref:
                s, c = ref.split("!", 1)
            else:
                s, c = sheet, ref
            v2, f2 = _read_cell_safe(wb, s, c)
            print(f"    Formula : {f2!r}")
            print(f"    Value   : {v2}")
            ok = input("  Use this? [y/n]: ").strip().lower()
            if ok == "y":
                return s, c
        else:
            # They typed a cell reference directly
            if "!" in ans:
                s, c = ans.split("!", 1)
            else:
                s, c = sheet, ans
            v2, f2 = _read_cell_safe(wb, s, c)
            print(f"    Formula : {f2!r}")
            print(f"    Value   : {v2}")
            ok = input("  Use this? [y/n]: ").strip().lower()
            if ok == "y":
                return s, c


def _find_candidate(wb, keywords: list[str], expected_type=None) -> tuple[str, str]:
    """Find first hit for any keyword; return (sheet, address)."""
    for kw in keywords:
        hits = _search_for_keyword(wb, kw)
        for sh_name, addr, val, formula in hits:
            if expected_type is None or isinstance(val, expected_type):
                # Strip $ from address
                clean = addr.replace("$", "")
                return sh_name, clean
    # Fallback: first sheet, A1
    return wb.sheets[0].name, "A1"


def _ask_range(wb, prompt: str, sheet: str, start_cell: str, n: int) -> tuple[str, str]:
    """Ask user to confirm a range of n cells starting at start_cell."""
    end_col = start_cell[0]
    end_row = int(start_cell[1:]) + n - 1
    end_cell = f"{end_col}{end_row}"
    rng = f"{start_cell}:{end_cell}"
    try:
        sh = wb.sheets[sheet]
        vals = sh.range(rng).value
    except Exception:
        vals = None
    print(f"\n  {prompt}")
    print(f"  Candidate range: {sheet}!{rng}")
    print(f"    Values: {vals}")
    ans = input("  Confirm? [y / n / Sheet!Range]: ").strip()
    if ans.lower() == "y":
        return sheet, rng
    if ans.lower() == "n":
        ref = input("  Enter the correct range (e.g. 'Rates!B2:B11'): ").strip()
        if "!" in ref:
            s, r = ref.split("!", 1)
        else:
            s, r = sheet, ref
        return s, r
    if "!" in ans:
        s, r = ans.split("!", 1)
        return s, r
    return sheet, ans


def _ask_expiry_list() -> list[dict]:
    print("\n  Expiry list")
    print("  Enter each expiry code and settlement date, one per line.")
    print("  Format: CODE YYYY-MM-DD  (e.g. 'DEC26 2026-12-18')")
    print("  Press Enter on an empty line when done.")
    expiries = []
    while True:
        line = input("  > ").strip()
        if not line:
            break
        parts = line.split()
        if len(parts) == 2:
            code, dt_str = parts
            try:
                date.fromisoformat(dt_str)
                expiries.append({"code": code.upper(), "date": dt_str})
                print(f"    Added: {code.upper()} → {dt_str}")
            except ValueError:
                print("    Invalid date format, skipping.")
        else:
            print("    Expected: CODE YYYY-MM-DD")
    return expiries


# ---------------------------------------------------------------------------
# Main discovery flow
# ---------------------------------------------------------------------------

def run_discovery(path: str) -> None:
    wb, app = _connect(path)
    sheet_names = _print_sheets(wb)
    _print_named_ranges(wb)

    print("\n" + "="*60)
    print("  CELL MAPPING DISCOVERY")
    print("  I will suggest a cell for each input. Answer y / n / Cell.")
    print("="*60)

    cfg: dict = {"workbook": path, "platform": "windows"}

    # 1. SPX Spot
    sh, cell = _find_candidate(wb, ["spx", "px_last", "spot", "index"], float)
    sh, cell = _ask_confirm(wb, "SPX cash spot price", sh, cell)
    cfg["spot"] = {"sheet": sh, "cell": cell}

    # 2. SOFR curve (column of rates)
    sh, cell = _find_candidate(wb, ["sofr", "rate", "libor", "ois"], float)
    sh, rng = _ask_range(
        wb, "SOFR rates (overnight → 10Y, one cell per tenor)", sh, cell, 9
    )
    tenors_input = input(
        "  Tenors in years, comma-separated [default: 0.003,0.083,0.25,0.5,1,2,3,5,10]: "
    ).strip()
    if tenors_input:
        tenors = [float(t.strip()) for t in tenors_input.split(",")]
    else:
        tenors = [0.0027, 0.083, 0.25, 0.5, 1, 2, 3, 5, 10]
    cfg["sofr_curve"] = {"sheet": sh, "range": rng, "tenors_years": tenors}

    # 3. Dividend strip
    sh, cell = _find_candidate(wb, ["asd", "div", "dividend", "strip"], float)
    sh, rng = _ask_range(
        wb, "Dividend strip (one row per expiry year, e.g. ASD2026Z values)", sh, cell, 10
    )
    years_input = input(
        "  Year labels, comma-separated [e.g. DEC26,DEC27,DEC28,DEC29,DEC30]: "
    ).strip()
    if years_input:
        div_labels = [y.strip().upper() for y in years_input.split(",")]
    else:
        div_labels = ["DEC26", "DEC27", "DEC28", "DEC29", "DEC30"]
    cfg["div_strip"] = {"sheet": sh, "range": rng, "expiry_labels": div_labels}

    # 4. Borrow curve
    sh, cell = _find_candidate(wb, ["air", "borrow", "repo", "lend"], float)
    sh, rng = _ask_range(
        wb, "Borrow spread per expiry (AIR futures or implied)", sh, cell, len(div_labels)
    )
    cfg["borrow_curve"] = {"sheet": sh, "range": rng, "expiry_labels": div_labels}

    # 5. ES front-month future
    sh, cell = _find_candidate(wb, ["es1", "es ", "future", "emini"], float)
    sh, cell = _ask_confirm(wb, "ES front-month futures price (ES1 Index)", sh, cell)
    cfg["es_future"] = {"sheet": sh, "cell": cell}

    # 6. ES-SPX basis
    sh, cell = _find_candidate(wb, ["basis", "es-spx", "es_spx"], float)
    sh, cell = _ask_confirm(wb, "ES-SPX basis (ES price minus SPX, or manually agreed)", sh, cell)
    cfg["es_basis"] = {"sheet": sh, "cell": cell}

    # 7. Expiry list
    expiries = _ask_expiry_list()
    if not expiries:
        print("  Using example expiries (edit config/excel_map.yaml to change).")
        expiries = [
            {"code": "DEC26", "date": "2026-12-18"},
            {"code": "DEC27", "date": "2027-12-17"},
            {"code": "DEC30", "date": "2030-12-19"},
        ]
    cfg["expiries"] = expiries

    # Write config
    CONFIG_DIR.mkdir(exist_ok=True)
    import yaml
    with open(CONFIG_PATH, "w") as f:
        yaml.dump(cfg, f, default_flow_style=False, sort_keys=False)

    print(f"\n{'='*60}")
    print(f"  Saved: {CONFIG_PATH}")
    print()
    print("  Next step:")
    print("    python -m pricer.snapshot")
    print("  to verify the snapshot matches your Excel.")
    print("="*60)


if __name__ == "__main__":
    if not Path(WORKBOOK_PATH).exists():
        alt = input(
            f"Workbook not found at:\n  {WORKBOOK_PATH}\n"
            "Enter the full path to the .xlsx file: "
        ).strip().strip('"')
        WORKBOOK_PATH = alt
    run_discovery(WORKBOOK_PATH)

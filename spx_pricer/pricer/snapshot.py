"""
python -m pricer.snapshot

Reads a live MarketSnapshot from the open Bloomberg-connected Excel workbook
and prints all values so the user can confirm they match Excel.
"""

from __future__ import annotations

import sys
from pathlib import Path

CONFIG_PATH = Path(__file__).parent.parent / "config" / "excel_map.yaml"


def main() -> None:
    if not CONFIG_PATH.exists():
        print(
            f"ERROR: {CONFIG_PATH} not found.\n"
            "Run `python discover.py` first to map your Excel cells."
        )
        sys.exit(1)

    from pricer.excel import read_snapshot, StaleSnapshotError

    print("Connecting to Excel and forcing recalculation...")
    try:
        snap = read_snapshot(CONFIG_PATH)
    except ConnectionError as exc:
        print(f"\nEXCEL ERROR: {exc}")
        sys.exit(1)
    except StaleSnapshotError as exc:
        print(f"\nSTALE DATA: {exc}")
        sys.exit(1)
    except Exception as exc:
        print(f"\nUNEXPECTED ERROR: {exc}")
        raise

    W = 55
    print(f"\n{'='*W}")
    print(f"  MARKET SNAPSHOT  {snap.timestamp:%Y-%m-%d %H:%M:%S}")
    print(f"{'='*W}")
    print(f"  Spot (SPX)    : {snap.spot:>12,.2f}")
    print(f"  ES Future     : {snap.es_future:>12,.2f}")
    print(f"  ES Basis      : {snap.es_basis:>+12.2f}")
    print()
    print("  SOFR Curve")
    for tenor, rate in sorted(snap.sofr.items()):
        label = f"{tenor:.3f}y" if tenor < 1 else f"{tenor:.1f}y"
        print(f"    {label:<6}  :  {rate*100:>8.4f}%")
    print()
    print("  Dividend Yields  (per expiry)")
    for code, yld in snap.div_yield_curve.items():
        print(f"    {code:<10}  :  {yld*100:>8.4f}%")
    print()
    print("  Borrow Spreads  (per expiry)")
    for code, b in snap.borrow.items():
        print(f"    {code:<10}  :  {b*10000:>8.1f} bps")
    print()
    print("  Expiries")
    for code, dt in snap.expiries.items():
        print(f"    {code:<10}  :  {dt}")
    print(f"{'='*W}")
    print(
        "\nEyeball the numbers above against your open Excel.\n"
        "If anything looks wrong, re-run `python discover.py` to remap the cells."
    )


if __name__ == "__main__":
    main()

"""
SPX Combo & Revcon — CLI entry point.

Usage:
    python -m pricer.cli ladder  --expiry DEC26 --width-bps 10
    python -m pricer.cli revcon  --expiry DEC26 --strike 7200 --combo 20
                                  --es 7160 --basis 40 --side REV --lots 100
    python -m pricer.cli reverse --expiry DEC26 --strike 7200 --combo-mid 22.5
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path
from typing import Optional

import typer

app = typer.Typer(
    name="pricer",
    help="SPX Combo & Revcon live pricer",
    no_args_is_help=True,
)
CONFIG_PATH = Path(__file__).parent.parent / "config" / "excel_map.yaml"

logging.basicConfig(level=logging.WARNING, format="%(levelname)s %(message)s")


def _snap():
    from pricer.excel import read_snapshot, StaleSnapshotError

    if not CONFIG_PATH.exists():
        typer.echo(
            f"Config not found at {CONFIG_PATH}.\n"
            "Run `python discover.py` first.",
            err=True,
        )
        raise typer.Exit(1)
    try:
        return read_snapshot(CONFIG_PATH)
    except ConnectionError as exc:
        typer.echo(f"EXCEL ERROR: {exc}", err=True)
        raise typer.Exit(1)
    except StaleSnapshotError as exc:
        typer.echo(f"STALE DATA — cannot price: {exc}", err=True)
        raise typer.Exit(1)


@app.command()
def ladder(
    expiry: str = typer.Option(..., help="Expiry code e.g. DEC26"),
    width_bps: float = typer.Option(10.0, "--width-bps", help="Bid/ask half-width in bps of forward"),
):
    """Print a 7-strike combo ladder centred on ATM forward."""
    snap = _snap()
    from pricer.ladder import build_ladder, format_ladder
    result = build_ladder(snap, expiry.upper(), width_bps)
    typer.echo(format_ladder(result))


@app.command()
def revcon(
    expiry: str = typer.Option(..., help="Expiry code e.g. DEC26"),
    strike: float = typer.Option(..., help="Option strike"),
    combo: float = typer.Option(..., help="Combo price (mid) you are pricing"),
    es: float = typer.Option(..., "--es", help="ES futures hedge level"),
    basis: float = typer.Option(..., help="Agreed ES-SPX basis"),
    side: str = typer.Option("REV", help="REV (buy combo, sell future) or CONV (mirror)"),
    lots: int = typer.Option(1, help="Number of lots"),
):
    """Calculate revcon P&L — LIFT / PASS / FLAT verdict."""
    snap = _snap()
    from pricer.revcon import calc_revcon, format_revcon
    result = calc_revcon(snap, expiry.upper(), strike, combo, es, basis, side, lots)
    typer.echo(format_revcon(result, strike, combo, expiry.upper(), side.upper(), lots))


@app.command()
def reverse(
    expiry: str = typer.Option(..., help="Expiry code e.g. DEC26"),
    strike: float = typer.Option(..., help="Option strike"),
    combo_mid: float = typer.Option(..., "--combo-mid", help="Street combo mid to back out borrow from"),
):
    """Back out implied borrow from a street combo mid — RICH or CHEAP signal."""
    snap = _snap()
    from pricer.reverse import calc_reverse, format_reverse
    result = calc_reverse(snap, expiry.upper(), strike, combo_mid)
    typer.echo(format_reverse(result, strike, expiry.upper()))


def main():
    app()


if __name__ == "__main__":
    main()

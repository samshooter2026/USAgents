"""Revcon package calculator."""

from __future__ import annotations

from datetime import date

from pricer import math as m
from pricer.excel import MarketSnapshot, pick_sofr


def calc_revcon(
    snap: MarketSnapshot,
    expiry_code: str,
    strike: float,
    combo: float,
    es: float,
    basis: float,
    side: str,
    lots: int,
) -> dict:
    today = date.today()
    settle = snap.expiries[expiry_code]
    T = (settle - today).days / 365.0
    _, r = pick_sofr(snap.sofr, T)
    return m.revcon_pnl(
        K=strike,
        combo=combo,
        r=r,
        T=T,
        es_future=es,
        basis=basis,
        side=side.upper(),
        lots=lots,
        spot=snap.spot,
    )


def format_revcon(
    result: dict,
    strike: float,
    combo: float,
    expiry: str,
    side: str,
    lots: int,
) -> str:
    verdict_pad = f"  {'='*38}"
    lines = [
        f"\nRevcon Package  [{expiry}  {side}  {lots:,} lots]",
        "-" * 42,
        f"  Strike        = {strike:>10,.0f}",
        f"  Combo input   = {combo:>+10.2f}",
        f"  F_combo       = {result['f_combo']:>10,.2f}",
        f"  F_hedge       = {result['f_hedge']:>10,.2f}",
        f"  Edge          = {result['edge_pts']:>+10.2f} pts",
        f"  P&L $         = ${result['pnl_dollars']:>+12,.0f}",
        f"  bps ann       = {result['bps_ann']:>+10.1f} bps",
        verdict_pad,
        f"  Verdict       = {result['verdict']}",
        "-" * 42,
    ]
    return "\n".join(lines)

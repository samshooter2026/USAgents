"""Reverse implied-borrow solver."""

from __future__ import annotations

from datetime import date

from pricer import math as m
from pricer.excel import MarketSnapshot, pick_sofr


def calc_reverse(
    snap: MarketSnapshot,
    expiry_code: str,
    strike: float,
    combo_mid_market: float,
) -> dict:
    today = date.today()
    settle = snap.expiries[expiry_code]
    T = (settle - today).days / 365.0
    _, r = pick_sofr(snap.sofr, T)
    q = snap.div_yield_curve.get(
        expiry_code,
        next(iter(snap.div_yield_curve.values()), 0.0),
    )
    b_model = snap.borrow.get(expiry_code, 0.0)

    result = m.reverse_borrow(strike, combo_mid_market, r, T, snap.spot, q)
    result["b_model_bps"] = b_model * 10000
    result["delta_bps"] = result["borrow_bps"] - result["b_model_bps"]
    result["signal"] = "RICH" if result["delta_bps"] > 0 else "CHEAP"
    return result


def format_reverse(result: dict, strike: float, expiry: str) -> str:
    lines = [
        f"\nReverse Borrow Solver  [{expiry}  K={strike:,.0f}]",
        "-" * 42,
        f"  Implied forward   = {result['f_implied']:>10,.2f}",
        f"  Implied (r+b)     = {result['r_plus_b']*100:>10.4f}%",
        f"  Implied borrow    = {result['borrow_bps']:>10.0f} bps",
        f"  Your assumption   = {result['b_model_bps']:>10.0f} bps  (from Excel)",
        f"  Delta             = {result['delta_bps']:>+10.0f} bps  → {result['signal']}",
        "-" * 42,
    ]
    return "\n".join(lines)


def format_implied_curve(
    snap: MarketSnapshot,
    combo_mids: dict[str, float],
) -> str:
    """Format a full implied-borrow curve across all expiries."""
    today = date.today()
    lines = [
        f"\n{'Expiry':<10}  {'T':>5}  {'Impl Fwd':>10}  {'Impl Borrow':>12}  {'Model':>7}  {'Delta':>7}  Signal",
        "-" * 72,
    ]
    for code, settle in snap.expiries.items():
        if code not in combo_mids:
            continue
        T = (settle - today).days / 365.0
        _, r = pick_sofr(snap.sofr, T)
        q = snap.div_yield_curve.get(code, next(iter(snap.div_yield_curve.values()), 0.0))
        b_model = snap.borrow.get(code, 0.0)
        strike = round(m.forward(snap.spot, q, r, b_model, T) / 100) * 100
        res = m.reverse_borrow(combo_mids[code], combo_mids[code], r, T, snap.spot, q)
        delta = res["borrow_bps"] - b_model * 10000
        sig = "RICH" if delta > 0 else "CHEAP"
        lines.append(
            f"{code:<10}  {T:>5.2f}  {res['f_implied']:>10,.2f}  "
            f"{res['borrow_bps']:>12.1f}  {b_model*10000:>7.1f}  {delta:>+7.1f}  {sig}"
        )
    return "\n".join(lines)

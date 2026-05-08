"""
SPX Combo & Revcon pricing formulas — spx_combo_revcon_handoff.pdf v0.1

All formulas are pure functions (no I/O). Unit-tested in tests/test_math.py.

Forward (PDF §2):
    F = (S − PV_divs) × exp((r + b) × T)

PV of dividends (simple linear approximation, ACT/365 day count):
    PV_divs = S × q × T
    where q is annual div yield (decimal), T is year-fraction

Discount factor:
    DF = exp(−r × T)

Combo mid (PDF §2, put-call parity):
    Combo(K) = exp(−rT) × (F − K)

Edge per side for the ladder (PDF §4.1):
    edge_per_side = (width_bps / 10000) × F × DF / 2
    Bid = Mid − edge_per_side
    Ask = Mid + edge_per_side

Revcon P&L (PDF §4.2, REV side: buy combo, sell future):
    F_combo  = K + Combo × exp(rT)
    F_hedge  = ES_future + basis
    Edge     = F_hedge − F_combo          # positive on REV = lift
    P&L_$    = Edge × $100 × lots
    bps_ann  = (Edge / S) / T × 10000
    Verdict  = LIFT if bps_ann > +1, PASS if < −1, FLAT otherwise

Reverse-implied borrow (PDF §4.3):
    F_implied   = Combo × exp(rT) + K
    (r + b)     = ln(F_implied / (S − PV_divs)) / T
    borrow_bps  = ((r + b) − r) × 10000

Sensitivities per ATM combo (PDF §4.4, linear approximations):
    dCombo/d(SOFR_bp)  ≈  T × F × DF × 0.0001
    dCombo/d(div_bp)   ≈ −T × S × DF × 0.0001   (higher divs → lower fwd)

CONV is the mirror of REV: sell combo, buy future, sign of Edge flips.
SPX multiplier is $100 per point.
Day count: ACT/365 throughout.
"""

from __future__ import annotations

import math
from typing import Literal

SPX_MULTIPLIER = 100  # $100 per index point


# ---------------------------------------------------------------------------
# Core building blocks
# ---------------------------------------------------------------------------

def discount_factor(r: float, T: float) -> float:
    """DF = exp(−r × T)"""
    return math.exp(-r * T)


def pv_divs(spot: float, div_yield: float, T: float) -> float:
    """PV_divs = S × q × T  (linear approximation, ACT/365)."""
    return spot * div_yield * T


def forward(spot: float, div_yield: float, r: float, borrow: float, T: float) -> float:
    """
    F = (S − PV_divs) × exp((r + b) × T)

    spot      : SPX cash spot
    div_yield : annual dividend yield (decimal, e.g. 0.0125 for 1.25%)
    r         : SOFR matched to expiry tenor (decimal)
    borrow    : borrow spread (decimal, e.g. 0.0012 for 12 bps)
    T         : year-fraction to expiry (ACT/365)
    """
    pvd = pv_divs(spot, div_yield, T)
    return (spot - pvd) * math.exp((r + borrow) * T)


def combo_mid(F: float, K: float, r: float, T: float) -> float:
    """
    Combo(K) = exp(−rT) × (F − K)

    F : implied forward
    K : strike
    r : SOFR rate (decimal)
    T : year-fraction (ACT/365)
    """
    return discount_factor(r, T) * (F - K)


def edge_per_side(width_bps: float, F: float, r: float, T: float) -> float:
    """
    edge_per_side = (width_bps / 10000) × F × DF / 2

    Bid = Mid − edge_per_side
    Ask = Mid + edge_per_side
    """
    DF = discount_factor(r, T)
    return (width_bps / 10000) * F * DF / 2


# ---------------------------------------------------------------------------
# Ladder
# ---------------------------------------------------------------------------

def ladder_strikes(
    F: float,
    spacing: float = 100.0,
    n_above: int = 3,
    n_below: int = 3,
) -> list[float]:
    """7-strike ladder centred on ATM forward, rounded to nearest spacing."""
    atm = round(F / spacing) * spacing
    return [atm + (i - n_below) * spacing for i in range(n_below + 1 + n_above)]


# ---------------------------------------------------------------------------
# Revcon package
# ---------------------------------------------------------------------------

def revcon_pnl(
    K: float,
    combo: float,
    r: float,
    T: float,
    es_future: float,
    basis: float,
    side: Literal["REV", "CONV"],
    lots: int,
    spot: float,
) -> dict:
    """
    Revcon P&L (PDF §4.2).

    REV:  buy combo (synthetic long) + sell future (delta hedge).
          Edge = F_hedge − F_combo  (positive = lift)
    CONV: sell combo + buy future.
          Edge sign flips.

    F_combo  = K + Combo × exp(rT)
    F_hedge  = ES_future + basis
    Edge     = F_hedge − F_combo   (REV sign convention)
    P&L_$    = Edge × $100 × lots
    bps_ann  = (Edge / S) / T × 10000
    """
    f_combo = K + combo * math.exp(r * T)
    f_hedge = es_future + basis
    edge = f_hedge - f_combo
    if side == "CONV":
        edge = -edge
    pnl_dollars = edge * SPX_MULTIPLIER * lots
    bps_ann = (edge / spot) / T * 10000
    verdict = "LIFT" if bps_ann > 1 else ("PASS" if bps_ann < -1 else "FLAT")
    return {
        "f_combo": f_combo,
        "f_hedge": f_hedge,
        "edge_pts": edge,
        "pnl_dollars": pnl_dollars,
        "bps_ann": bps_ann,
        "verdict": verdict,
    }


# ---------------------------------------------------------------------------
# Reverse-implied borrow solver
# ---------------------------------------------------------------------------

def reverse_borrow(
    K: float,
    combo_mid_market: float,
    r: float,
    T: float,
    spot: float,
    div_yield: float,
) -> dict:
    """
    Implied borrow from a market combo mid (PDF §4.3).

    F_implied   = Combo × exp(rT) + K
    (r + b)     = ln(F_implied / (S − PV_divs)) / T
    borrow_bps  = ((r + b) − r) × 10000
    """
    f_implied = combo_mid_market * math.exp(r * T) + K
    pvd = pv_divs(spot, div_yield, T)
    r_plus_b = math.log(f_implied / (spot - pvd)) / T
    borrow_bps = (r_plus_b - r) * 10000
    return {
        "f_implied": f_implied,
        "r_plus_b": r_plus_b,
        "borrow_bps": borrow_bps,
    }


def implied_borrow_from_fs(fs_ratio: float, r: float, T: float, div_yield: float) -> float:
    """
    Implied borrow (decimal) from a forward/spot ratio quote.

    F/S = (1 − qT) × exp((r+b)T)
    => b = ln( (F/S) / (1 − qT) ) / T − r

    fs_ratio  : F divided by S (e.g. 1.019275 for "101.9275%")
    r         : SOFR matched to expiry tenor (decimal)
    T         : year-fraction to expiry (ACT/365)
    div_yield : annual dividend yield (decimal)

    No spot or strike needed — the ratio is dimensionless, so this works for
    "% of spot" street prints where the absolute spot at trade time isn't given.
    """
    if T <= 0 or fs_ratio <= 0:
        return 0.0
    return math.log(fs_ratio / (1 - div_yield * T)) / T - r


# ---------------------------------------------------------------------------
# Sensitivities (PDF §4.4, linear approximations)
# ---------------------------------------------------------------------------

def sensitivity_sofr(T: float, F: float, r: float) -> float:
    """dCombo / +1bp SOFR ≈ T × F × DF × 0.0001  (PDF §4.4)"""
    return T * F * discount_factor(r, T) * 0.0001


def sensitivity_div(T: float, spot: float, r: float) -> float:
    """dCombo / +1bp div yield ≈ −T × S × DF × 0.0001  (PDF §4.4)
    Negative: higher divs lower the forward."""
    return -T * spot * discount_factor(r, T) * 0.0001

"""Combo strike ladder builder."""

from __future__ import annotations

from datetime import date

from pricer import math as m
from pricer.excel import MarketSnapshot, pick_sofr


def build_ladder(
    snap: MarketSnapshot,
    expiry_code: str,
    width_bps: float,
    es_override: float | None = None,
    specific_strike: float | None = None,
) -> dict:
    """
    Build a combo ladder for a given expiry.

    es_override: if set, replaces live ES front-month and shifts the implied
                 spot by the same delta (preserves live basis). Use to match a
                 quote where the counterparty cited a specific ES reference.
    specific_strike: if set, builds a single-row ladder at exactly that strike
                     (useful for matching a known counterparty quote).
                     If None, builds the standard 7-strike ATM ladder.

    Returns a dict with rows (Strike, Bid, Mid, Ask, % Fwd) plus footer params.
    """
    today = date.today()
    settle = snap.expiries[expiry_code]
    T = (settle - today).days / 365.0

    _, r = pick_sofr(snap.sofr, T)

    q = snap.div_yield_curve.get(
        expiry_code,
        next(iter(snap.div_yield_curve.values()), 0.0),
    )
    b = snap.borrow.get(expiry_code, 0.0)

    # Effective spot — if ES override given, shift live spot by the ES delta
    # (assumes front-month basis is locally constant, which is a fine
    # approximation for the small ES moves you'd type to match a quote).
    if es_override is not None:
        spot_eff = snap.spot + (es_override - snap.es_future)
    else:
        spot_eff = snap.spot

    F = m.forward(spot_eff, q, r, b, T)
    DF = m.discount_factor(r, T)
    eps = m.edge_per_side(width_bps, F, r, T)

    if specific_strike is not None:
        strikes = [float(specific_strike)]
    else:
        strikes = m.ladder_strikes(F)

    rows = []
    for K in strikes:
        mid = m.combo_mid(F, K, r, T)
        rows.append({
            "strike": K,
            "bid": mid - eps,
            "mid": mid,
            "ask": mid + eps,
            "pct_fwd": (mid / F) * 100,
        })

    return {
        "expiry": expiry_code,
        "settle": settle,
        "F": F,
        "T": T,
        "r": r,
        "b": b,
        "q": q,
        "DF": DF,
        "width_bps": width_bps,
        "spot_eff": spot_eff,
        "es_used": es_override if es_override is not None else snap.es_future,
        "rows": rows,
    }


def format_ladder(result: dict) -> str:
    header = f"\n{'Strike':>8}  {'Bid':>10}  {'Mid':>10}  {'Ask':>10}  {'% Fwd':>8}"
    sep = "-" * 56
    lines = [header, sep]
    for row in result["rows"]:
        lines.append(
            f"{row['strike']:>8,.0f}  "
            f"{row['bid']:>+10.2f}  "
            f"{row['mid']:>+10.2f}  "
            f"{row['ask']:>+10.2f}  "
            f"{row['pct_fwd']:>+7.2f}%"
        )
    lines.append(sep)
    lines.append(
        f"F={result['F']:,.2f}  T={result['T']:.2f}y  "
        f"r={result['r']*100:.2f}%  b={result['b']*10000:.0f}bp  "
        f"q={result['q']*100:.2f}%  DF={result['DF']:.4f}"
    )
    return "\n".join(lines)

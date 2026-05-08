"""
AXW calibration agent — persists every street combo print observed, computes
implied borrow from each, and flags when AXW drifts from real flow.

Storage: parquet file at logs/axw_calibration.parquet (one row per print).
Drift rules (matched to handoff doc):
  - per-print tolerance: ±5 bps
  - flag when ≥3 consecutive prints in same direction exceed tolerance
"""
from __future__ import annotations

import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

LOG_PATH = Path(__file__).parent.parent / "logs" / "axw_calibration.parquet"

# ── Drift thresholds (overridable via defaults.yaml later if useful) ─────────
DRIFT_TOLERANCE_BPS = 5     # individual print tolerance vs current AXW
DRIFT_WINDOW        = 5     # how many recent prints per expiry to consider
DRIFT_MIN_PRINTS    = 3     # need at least this many same-side prints to alert


# ── Free-form text parsing ──────────────────────────────────────────────────
# Designed for broker chat one-liners.  Examples:
#   "20:38:10 SPX Dec26 combo trades 650mm ~101.9275% makes 1.2b"
#   "SPX Dec26 combo 550mm at 101.9275% / 7338f"
#   "SPX Dec31 Combo (7153f, LIVE) 121.52/121.78"

_EXPIRY_RE   = re.compile(
    r"\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s*(\d{2,4})\b",
    re.IGNORECASE,
)
_TIME_RE     = re.compile(r"\b(\d{1,2}):(\d{2})(?::(\d{2}))?\b")
_NOTIONAL_RE = re.compile(r"\b(\d+(?:\.\d+)?)\s*(mm|m|b)\b", re.IGNORECASE)
_PCT_RE      = re.compile(r"~?\s*(\d{2,3}\.\d+)\s*%")
_ES_LEVEL_RE = re.compile(r"\b(\d{4,5}(?:\.\d+)?)\s*f\b", re.IGNORECASE)
_PTS_BIDASK_RE = re.compile(r"(\d{1,4}\.\d{1,4})\s*/\s*(\d{1,4}\.\d{1,4})")

# AXW direct trade patterns:
#   "AXWZ8 trades at 73.50"
#   "AXWZ28 @ 73.5"
#   "AXWZ7 at 66.00"
#   "DEC26 AXW just traded at 56bps"     (loose form)
#   "AXW DEC26 @ 56"                      (loose form, reversed order)
_AXW_LEG_RE = re.compile(
    r"AXW\s*Z\s*(\d{1,2})\b"
    r"\s*(?:trades?\s*)?(?:at\s+|@\s+)?"
    r"(\d{1,3}(?:\.\d+)?)",
    re.IGNORECASE,
)
# Loose-form: number followed by 'bp' or 'bps' (e.g. "56bps", "73.5 bp")
_BPS_PRICE_RE = re.compile(
    r"(\d{1,3}(?:\.\d+)?)\s*(?:bps?)\b",
    re.IGNORECASE,
)
_AXW_KEYWORD_RE = re.compile(r"\baxw\b", re.IGNORECASE)
# Size: "500x" or "size 500"
_AXW_SIZE_RE  = re.compile(r"\b(\d+)\s*x\b", re.IGNORECASE)
_AXW_SIZE2_RE = re.compile(r"\bsize\s+(\d+)\b", re.IGNORECASE)


def parse_print_text(text: str) -> dict:
    """
    Parse a free-form street print into structured fields.

    Always returns:
        expiry  : str (e.g. "DEC26")
        format  : "pct" or "pts"
        raw     : the original text

    Optional fields, populated if found:
        fs_pct          : float — % of spot (when format="pct")
        combo_bid_pts   : float — when format="pts"
        combo_ask_pts   : float — when format="pts"
        combo_mid_pts   : float — (bid+ask)/2 when format="pts"
        size_mm         : float — notional in millions
        es_level        : float — ES futures level the trade was struck at
        trade_time      : str   — HH:MM:SS

    Raises ValueError if expiry or price can't be located.
    """
    if not text or not text.strip():
        raise ValueError("Empty input.")
    text = text.strip()
    out: dict = {"raw": text}

    m_exp = _EXPIRY_RE.search(text)
    if not m_exp:
        raise ValueError("Could not find an expiry like DEC26 / Dec27 / Mar28.")
    yr = m_exp.group(2)
    yr_short = yr[-2:] if len(yr) > 2 else yr
    out["expiry"] = f"{m_exp.group(1).upper()}{yr_short}"

    m_time = _TIME_RE.search(text)
    if m_time:
        h, mn, s = m_time.group(1), m_time.group(2), m_time.group(3) or "00"
        out["trade_time"] = f"{int(h):02d}:{mn}:{s}"

    m_not = _NOTIONAL_RE.search(text)
    if m_not:
        n = float(m_not.group(1))
        unit = m_not.group(2).lower()
        if unit == "b":
            n *= 1000
        out["size_mm"] = n

    m_es = _ES_LEVEL_RE.search(text)
    if m_es:
        out["es_level"] = float(m_es.group(1))

    # Prefer points bid/ask if present and no '%' anywhere.  Otherwise %.
    m_pts = _PTS_BIDASK_RE.search(text)
    has_pct = "%" in text
    if m_pts and not has_pct:
        out["format"] = "pts"
        out["combo_bid_pts"] = float(m_pts.group(1))
        out["combo_ask_pts"] = float(m_pts.group(2))
        out["combo_mid_pts"] = (out["combo_bid_pts"] + out["combo_ask_pts"]) / 2
    else:
        m_pct = _PCT_RE.search(text)
        if not m_pct:
            raise ValueError(
                "Could not find a price.  Need either 'XXX.XX%' or 'XXX.XX/YYY.YY'."
            )
        out["format"] = "pct"
        out["fs_pct"] = float(m_pct.group(1))

    return out


def parse_axw_direct(text: str) -> dict:
    """
    Parse an AXW direct trade — single or multi-leg.  Examples:
      "AXWZ28 trades at 73.5 size 100"
      "*** WE TRADE 500x *** AXWZ8 trades at 73.50 vs AXWZ7 at 66.00"

    Returns:
        {
            "kind":   "axw_direct",
            "legs":   [{"expiry": "DEC28", "borrow_bps": 73.5}, ...],
            "size":   500 | None,           # contracts, applied to all legs
            "trade_time": "15:38:10" | None,
            "raw":    original text,
        }

    Raises ValueError if no AXWZ reference found.
    """
    if not text or not text.strip():
        raise ValueError("Empty input.")
    text = text.strip()

    legs = []

    # Path 1 — strict ticker form: AXWZ<year> <price>
    for m in _AXW_LEG_RE.finditer(text):
        yr = m.group(1)
        if len(yr) == 1:
            yr = f"2{yr}"            # AXWZ8 -> 28
        legs.append({
            "expiry":      f"DEC{yr}",
            "borrow_bps":  float(m.group(2)),
        })

    # Path 2 — loose form: needs the word "axw" + an explicit expiry + a "Xbps" price
    if not legs and _AXW_KEYWORD_RE.search(text):
        expiry_matches = list(_EXPIRY_RE.finditer(text))
        bps_matches    = list(_BPS_PRICE_RE.finditer(text))
        if not expiry_matches:
            raise ValueError(
                "AXW mentioned but no expiry like DEC26 / DEC27 / DEC28."
            )
        if not bps_matches:
            raise ValueError(
                "AXW + expiry mentioned but no bps price found "
                "(write the price like '56bps' so it's unambiguous)."
            )
        # Pair in order: leg N = expiry N + bps price N
        for em, bm in zip(expiry_matches, bps_matches):
            yr_short = em.group(2)[-2:]
            month = em.group(1).upper()
            legs.append({
                "expiry":      f"{month}{yr_short}",
                "borrow_bps":  float(bm.group(1)),
            })

    if not legs:
        raise ValueError(
            "No AXW reference found.  Expected 'AXWZ28 trades at 73.5' or "
            "'DEC28 AXW traded 73.5bps'."
        )

    m_time = _TIME_RE.search(text)
    trade_time = None
    if m_time:
        h, mn, s = m_time.group(1), m_time.group(2), m_time.group(3) or "00"
        trade_time = f"{int(h):02d}:{mn}:{s}"

    m_size = _AXW_SIZE_RE.search(text) or _AXW_SIZE2_RE.search(text)
    size = int(m_size.group(1)) if m_size else None

    return {
        "kind":       "axw_direct",
        "legs":       legs,
        "size":       size,
        "trade_time": trade_time,
        "raw":        text,
    }


def parse_any(text: str) -> dict:
    """
    Auto-detect format: AXW direct trade or combo print.  Returns the same
    shape as the underlying parser, plus a top-level "kind" key.
    """
    if not text or not text.strip():
        raise ValueError("Empty input.")
    # AXW direct: either strict AXWZ<n> ticker, or "axw" keyword + bps price
    if _AXW_LEG_RE.search(text) or (
        _AXW_KEYWORD_RE.search(text) and _BPS_PRICE_RE.search(text)
    ):
        return parse_axw_direct(text)
    parsed = parse_print_text(text)
    parsed["kind"] = "combo"
    return parsed


def log_axw_direct(
    *,
    expiry: str,
    borrow_bps: float,
    axw_borrow_bps_at_log: float,
    contracts: Optional[int],
    source: Optional[str],
) -> None:
    """Append an AXW direct trade to the calibration log.  Same schema as
    log_print() but with combo_mid/strike/es_at_trade left blank."""
    import pandas as pd

    LOG_PATH.parent.mkdir(exist_ok=True)
    row = {
        "timestamp":          datetime.now().isoformat(),
        "expiry":             expiry,
        "strike":             0.0,
        "combo_mid":          0.0,
        "es_at_trade":        0.0,
        "spx_at_trade":       0.0,
        "notional_mm":        float(contracts) if contracts else None,
        "source":             (source or "AXW direct"),
        "implied_borrow_bps": float(borrow_bps),
        "axw_borrow_bps":     float(axw_borrow_bps_at_log),
        "delta_bps":          float(borrow_bps - axw_borrow_bps_at_log),
        "print_type":         "axw_direct",
    }
    new_df = pd.DataFrame([row])

    if LOG_PATH.exists():
        old_df = pd.read_parquet(LOG_PATH)
        if "print_type" not in old_df.columns:
            old_df["print_type"] = "combo"
        df = pd.concat([old_df, new_df], ignore_index=True)
    else:
        df = new_df
    df.to_parquet(LOG_PATH, index=False)


def implied_borrow_from_pct_print(
    fs_pct: float,
    expiry: str,
    snap,
) -> dict:
    """
    Compute implied borrow (bps) from a % of spot quote.

    Uses the snap's current SOFR + dividend curve + expiry settle.  The F/S
    ratio is dimensionless so absolute spot is not needed.

    Returns: {
        "borrow_bps":   implied b (bps, annualised),
        "axw_bps":      current AXW model value for this expiry (bps),
        "delta_bps":    implied − model,
        "T":            year-fraction to expiry,
        "r":            SOFR used (decimal),
        "q":            div yield used (decimal),
    }
    """
    from datetime import date as _date
    from pricer import math as m
    from pricer.excel import pick_sofr

    if expiry not in snap.expiries:
        raise ValueError(f"Unknown expiry '{expiry}'. Available: {list(snap.expiries.keys())}")

    settle = snap.expiries[expiry]
    T = (settle - _date.today()).days / 365.0
    _, r = pick_sofr(snap.sofr, T)
    q = snap.div_yield_curve.get(expiry, next(iter(snap.div_yield_curve.values()), 0.0))
    axw_decimal = snap.borrow.get(expiry, 0.0)

    fs_ratio = fs_pct / 100.0
    b = m.implied_borrow_from_fs(fs_ratio, r, T, q)
    borrow_bps = b * 10000
    axw_bps    = axw_decimal * 10000
    return {
        "borrow_bps": borrow_bps,
        "axw_bps":    axw_bps,
        "delta_bps":  borrow_bps - axw_bps,
        "T":          T,
        "r":          r,
        "q":          q,
    }


# ── Persistence ──────────────────────────────────────────────────────────────

def log_print(
    *,
    expiry: str,
    strike: float,
    combo_mid: float,
    es_at_trade: float,
    spx_at_trade: float,
    notional_mm: Optional[float],
    source: Optional[str],
    implied_borrow_bps: float,
    axw_borrow_bps: float,
) -> None:
    """Append one street print to the calibration log."""
    import pandas as pd

    LOG_PATH.parent.mkdir(exist_ok=True)
    row = {
        "timestamp":          datetime.now().isoformat(),
        "expiry":             expiry,
        "strike":             float(strike),
        "combo_mid":          float(combo_mid),
        "es_at_trade":        float(es_at_trade),
        "spx_at_trade":       float(spx_at_trade),
        "notional_mm":        float(notional_mm) if notional_mm else None,
        "source":             source or "",
        "implied_borrow_bps": float(implied_borrow_bps),
        "axw_borrow_bps":     float(axw_borrow_bps),
        "delta_bps":          float(implied_borrow_bps - axw_borrow_bps),
    }
    new_df = pd.DataFrame([row])

    if LOG_PATH.exists():
        old_df = pd.read_parquet(LOG_PATH)
        df = pd.concat([old_df, new_df], ignore_index=True)
    else:
        df = new_df
    df.to_parquet(LOG_PATH, index=False)


def read_log(since_days: Optional[int] = 30):
    """Read the calibration log; optionally filter to recent N days."""
    import pandas as pd
    cols = [
        "timestamp", "expiry", "strike", "combo_mid", "es_at_trade",
        "spx_at_trade", "notional_mm", "source",
        "implied_borrow_bps", "axw_borrow_bps", "delta_bps",
    ]
    if not LOG_PATH.exists():
        return pd.DataFrame(columns=cols)
    df = pd.read_parquet(LOG_PATH)
    df["timestamp"] = pd.to_datetime(df["timestamp"])
    if since_days is not None:
        cutoff = datetime.now() - timedelta(days=since_days)
        df = df[df["timestamp"] >= cutoff]
    return df.sort_values("timestamp", ascending=False).reset_index(drop=True)


# ── Drift detection ──────────────────────────────────────────────────────────

def detect_drift(df, axw_now_bps: dict[str, float]) -> dict[str, dict]:
    """
    For each expiry, summarise recent prints vs current AXW.

    Returns: {expiry_code: {
        "n_prints":         int,
        "mean_delta_bps":   float,    # mean of (impl − current AXW) across recent N
        "all_same_side":    bool,
        "alert":            str|None, # populated when DRIFT_MIN_PRINTS hit
        "rec_borrow_bps":   float,    # AXW shifted halfway toward mean implied
    }}
    """
    if df.empty:
        return {}

    out: dict[str, dict] = {}
    for expiry, grp in df.groupby("expiry"):
        recent = grp.head(DRIFT_WINDOW)
        if recent.empty:
            continue
        current_axw = axw_now_bps.get(expiry)
        if current_axw is None:
            continue

        deltas = recent["implied_borrow_bps"] - current_axw
        n           = len(deltas)
        mean_delta  = float(deltas.mean())
        all_pos     = bool((deltas > 0).all())
        all_neg     = bool((deltas < 0).all())
        same_side   = all_pos or all_neg

        alert = None
        if n >= DRIFT_MIN_PRINTS and same_side and abs(mean_delta) > DRIFT_TOLERANCE_BPS:
            direction = "rich" if all_pos else "cheap"
            alert = (
                f"{n} consecutive {direction} prints, "
                f"mean {mean_delta:+.1f}bp vs AXW"
            )

        rec_borrow = current_axw + (0.5 * mean_delta if alert else 0.0)

        out[expiry] = {
            "n_prints":       int(n),
            "mean_delta_bps": mean_delta,
            "all_same_side":  same_side,
            "alert":          alert,
            "rec_borrow_bps": float(rec_borrow),
        }
    return out


# ── Trader-style commentary ──────────────────────────────────────────────────

def _humanise_span(seconds: float) -> str:
    if seconds < 60:
        return "moments ago"
    if seconds < 3600:
        return f"in last {int(seconds/60)}m"
    if seconds < 7200:
        return "in last hour"
    if seconds < 86400:
        return f"in last {int(seconds/3600)}h"
    if seconds < 7 * 86400:
        return f"over {int(seconds/86400)}d"
    return f"over {int(seconds/86400)}d window"


def narrate_drift(
    df,
    axw_now_bps: dict[str, float],
    expiry_codes: list[str],
) -> dict[str, str]:
    """
    Produce a 1–3 sentence trader narrative per expiry.  Captures: direction,
    flow size, trend (widening/narrowing), and a recommendation.

    Returns: {expiry: narrative_string}.  Expiries with no logged prints are
    omitted.
    """
    from datetime import datetime as _dt

    if df.empty:
        return {}
    out: dict[str, str] = {}

    for expiry in expiry_codes:
        sub = df[df["expiry"] == expiry].head(DRIFT_WINDOW).copy()
        if sub.empty:
            continue
        current_axw = axw_now_bps.get(expiry)
        if current_axw is None:
            continue

        sub["delta_today"] = sub["implied_borrow_bps"] - current_axw
        n             = len(sub)
        latest_delta  = float(sub["delta_today"].iloc[0])    # head() = most recent first
        oldest_delta  = float(sub["delta_today"].iloc[-1])
        mean_delta    = float(sub["delta_today"].mean())
        total_size    = float(sub["notional_mm"].fillna(0).sum())
        all_pos       = bool((sub["delta_today"] > 0).all())
        all_neg       = bool((sub["delta_today"] < 0).all())
        same_side     = all_pos or all_neg

        # Time span across the window
        latest_ts  = sub["timestamp"].iloc[0]
        oldest_ts  = sub["timestamp"].iloc[-1]
        span_secs  = (latest_ts - oldest_ts).total_seconds() if n > 1 else 0
        span_str   = _humanise_span(span_secs) if n > 1 else "just now"

        # Trend (only meaningful with ≥2 same-side prints)
        trend = None
        if n >= 2 and same_side:
            if abs(latest_delta) > abs(oldest_delta) * 1.3:
                trend = "widening"
            elif abs(latest_delta) < abs(oldest_delta) * 0.7:
                trend = "narrowing"
            else:
                trend = "stable"

        # Build narrative
        bits: list[str] = []
        direction = "rich" if all_pos else ("cheap" if all_neg else "mixed")

        if n == 1:
            if abs(latest_delta) < 0.5:
                kind_phrase = "in line with AXW"
            elif latest_delta > 0:
                kind_phrase = "rich to AXW"
            else:
                kind_phrase = "cheap to AXW"
            size = f" on {total_size:.0f}mm" if total_size > 0 else ""
            bits.append(
                f"1 print {span_str} at {latest_delta:+.1f}bp ({kind_phrase}){size}."
            )
        else:
            size = f" on {total_size:.0f}mm total flow" if total_size > 0 else ""
            bits.append(
                f"{n} prints {span_str}, {direction} by mean {mean_delta:+.1f}bp{size}."
            )
            if trend == "widening":
                bits.append(
                    f"Gap widening (latest {latest_delta:+.1f}bp vs oldest "
                    f"{oldest_delta:+.1f}bp)."
                )
            elif trend == "narrowing":
                bits.append(
                    f"Gap narrowing (latest {latest_delta:+.1f}bp vs oldest "
                    f"{oldest_delta:+.1f}bp)."
                )

        # Recommendation
        if n >= DRIFT_MIN_PRINTS and same_side and abs(mean_delta) > DRIFT_TOLERANCE_BPS:
            rec = current_axw + 0.5 * mean_delta
            bits.append(
                f"**ALERT** — override AXW from {current_axw:.1f} to {rec:.1f}bp."
            )
        elif n >= 2 and same_side and abs(mean_delta) > DRIFT_TOLERANCE_BPS:
            need = DRIFT_MIN_PRINTS - n
            bits.append(
                f"One more {direction} print would trigger the override alert."
                if need == 1 else
                f"{need} more {direction} prints would trigger the override alert."
            )
        elif n >= 1 and abs(latest_delta) > DRIFT_TOLERANCE_BPS:
            bits.append(
                f"Outside ±{DRIFT_TOLERANCE_BPS}bp tolerance but only {n} print "
                f"so far — wait for confirmation."
            )
        else:
            bits.append("Within tolerance.")

        out[expiry] = " ".join(bits)
    return out


def market_summary(
    df,
    axw_now_bps: dict[str, float],
    expiry_codes: list[str],
) -> str:
    """
    Cross-curve commentary: where is the curve sitting relative to AXW
    overall, and which expiries are doing the heavy lifting.
    """
    if df.empty:
        return "No prints logged yet — log some street prints to start tracking AXW drift."

    # Per-expiry mean delta (vs current AXW), restricted to expiries with prints
    drift = detect_drift(df, axw_now_bps)
    if not drift:
        return "No prints in the active window."

    expiries_with_prints = [c for c in expiry_codes if c in drift]
    n_prints_total = sum(d["n_prints"] for d in drift.values())

    # Front-end vs back-end split (rough: years 1-3 vs 4+)
    front = [c for c in expiries_with_prints if c in ("DEC26", "DEC27", "DEC28")]
    back  = [c for c in expiries_with_prints if c in ("DEC29", "DEC30", "DEC31")]

    front_mean = (
        sum(drift[c]["mean_delta_bps"] for c in front) / len(front) if front else None
    )
    back_mean = (
        sum(drift[c]["mean_delta_bps"] for c in back) / len(back) if back else None
    )

    alerts = [c for c, d in drift.items() if d["alert"]]

    parts: list[str] = []
    parts.append(
        f"{n_prints_total} {'print' if n_prints_total==1 else 'prints'} across "
        f"{len(expiries_with_prints)} "
        f"{'expiry' if len(expiries_with_prints)==1 else 'expiries'}."
    )

    if front_mean is not None and back_mean is not None:
        front_word = "cheap" if front_mean < 0 else "rich"
        back_word  = "cheap" if back_mean < 0 else "rich"
        parts.append(
            f"Front-end ({', '.join(front)}) {front_mean:+.1f}bp ({front_word}); "
            f"back-end ({', '.join(back)}) {back_mean:+.1f}bp ({back_word})."
        )
    elif front_mean is not None:
        parts.append(
            f"Front-end ({', '.join(front)}) running mean {front_mean:+.1f}bp "
            f"vs AXW; no back-end prints yet."
        )
    elif back_mean is not None:
        parts.append(
            f"Back-end ({', '.join(back)}) running mean {back_mean:+.1f}bp "
            f"vs AXW; no front-end prints yet."
        )

    if alerts:
        parts.append(
            f"**🔴 Override alerts active on:** {', '.join(alerts)}."
        )
    else:
        # Soft warnings — expiries with same-side >2bp drift but not yet alerting
        soft = [
            c for c, d in drift.items()
            if d["all_same_side"] and abs(d["mean_delta_bps"]) > DRIFT_TOLERANCE_BPS
        ]
        if soft:
            parts.append(
                f"Watch list (consistent drift, more prints needed to confirm): "
                f"{', '.join(soft)}."
            )
        else:
            parts.append("All expiries within tolerance.")

    return " ".join(parts)

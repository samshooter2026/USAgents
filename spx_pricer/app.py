"""
SPX Combo & Revcon Pricer — Streamlit UI
Run: streamlit run app.py
"""

from __future__ import annotations

import logging
import time
from datetime import date, datetime
from pathlib import Path

import streamlit as st

CONFIG_PATH = Path(__file__).parent / "config" / "excel_map.yaml"
DEFAULTS_PATH = Path(__file__).parent / "config" / "defaults.yaml"

logging.basicConfig(level=logging.WARNING)

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="SPX Pricer",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Monospace numbers, sparse trader look
st.markdown(
    """
    <style>
    .stMetric label { font-size: 0.8rem; }
    .stMetric [data-testid="stMetricValue"] { font-family: monospace; }
    table { font-family: monospace; font-size: 0.9rem; }
    .stDataFrame { font-family: monospace; }
    div[data-testid="stVerticalBlock"] { gap: 0.3rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _load_defaults() -> dict:
    import yaml
    if DEFAULTS_PATH.exists():
        with open(DEFAULTS_PATH) as f:
            return yaml.safe_load(f) or {}
    return {}


@st.cache_data(ttl=5)
def _read_snap(_tick: int):
    """Cached snapshot — shared across all three tabs per refresh cycle."""
    from pricer.excel import read_snapshot, StaleSnapshotError
    try:
        snap = read_snapshot(CONFIG_PATH)
        return snap, None
    except ConnectionError as exc:
        return None, f"EXCEL ERROR: {exc}"
    except StaleSnapshotError as exc:
        return None, f"STALE DATA: {exc}"
    except Exception as exc:
        return None, f"ERROR: {exc}"


def _snap_age_color(snap) -> str:
    if snap is None:
        return "red"
    age = (datetime.now() - snap.timestamp).total_seconds()
    defs = _load_defaults()
    if age < defs.get("stale_amber_secs", 30):
        return "green"
    if age < defs.get("stale_red_secs", 120):
        return "orange"
    return "red"


def _yearend_flag(expiry_code: str, snap) -> bool:
    """True if this expiry crosses a Dec 31 boundary."""
    if snap is None or expiry_code not in snap.expiries:
        return False
    settle = snap.expiries[expiry_code]
    today = date.today()
    # Any Dec 31 between today and settle date
    for year in range(today.year, settle.year + 1):
        ye = date(year, 12, 31)
        if today < ye <= settle:
            return True
    return False


def _is_q4() -> bool:
    return date.today().month in (10, 11, 12)


# ---------------------------------------------------------------------------
# Sticky top bar
# ---------------------------------------------------------------------------

def render_top_bar(snap, err: str | None, tick: int):
    defs = _load_defaults()
    col1, col2, col3, col4 = st.columns([2, 2, 2, 1])
    with col1:
        if snap:
            age = (datetime.now() - snap.timestamp).total_seconds()
            color = _snap_age_color(snap)
            st.markdown(
                f"**Snapshot** &nbsp; "
                f"<span style='color:{color};font-family:monospace'>"
                f"{snap.timestamp:%H:%M:%S} ({age:.0f}s ago)</span>",
                unsafe_allow_html=True,
            )
        else:
            st.markdown("**Snapshot** &nbsp; <span style='color:red'>NOT CONNECTED</span>",
                        unsafe_allow_html=True)
    with col2:
        if snap:
            st.markdown(
                f"**SPX** &nbsp; <span style='font-family:monospace;font-size:1.1rem'>"
                f"{snap.spot:,.2f}</span>",
                unsafe_allow_html=True,
            )
    with col3:
        if snap:
            st.markdown(
                f"**ES** &nbsp; <span style='font-family:monospace'>{snap.es_future:,.2f}"
                f"</span> &nbsp; basis <span style='font-family:monospace'>"
                f"{snap.es_basis:+.1f}</span>",
                unsafe_allow_html=True,
            )
    with col4:
        if st.button("⟳ Refresh"):
            st.cache_data.clear()
            st.rerun()

    if err:
        st.error(err)
    if snap and _snap_age_color(snap) == "red":
        st.error("Data is stale (>120s). Prices are not reliable. Refresh Bloomberg and reconnect.")
        return False
    return True


# ---------------------------------------------------------------------------
# Tab 1 — Ladder
# ---------------------------------------------------------------------------

def render_ladder(snap):
    from pricer.ladder import build_ladder
    import pandas as pd

    defs = _load_defaults()
    expiry_codes = list(snap.expiries.keys()) if snap else []

    c1, c2 = st.columns([2, 3])
    with c1:
        expiry = st.selectbox("Expiry", expiry_codes, key="ladder_expiry")
        default_width = defs.get("width", {}).get("dealer_to_dealer", 10)
        if _is_q4():
            default_width += defs.get("yearend_extra_bps", 5)
        width_bps = st.slider(
            "Width (bps of forward)", min_value=5, max_value=25,
            value=int(default_width), step=1, key="ladder_width"
        )
    with c2:
        if _yearend_flag(expiry, snap):
            st.warning("⚠ Year-end expiry — Q4 borrow squeeze risk. Consider widening 5–10 bps.")

    if not snap:
        st.info("Waiting for Bloomberg connection…")
        return

    # ---------- Match-a-quote panel ----------
    with st.expander("🎯 Match a counterparty quote (ES override / exact strike)"):
        live_es = float(snap.es_future)
        cc1, cc2, cc3 = st.columns(3)
        with cc1:
            use_es_override = st.checkbox(
                "Override ES futures ref",
                value=False,
                key="ladder_use_es_override",
                help="Tick to price the ladder against a specific ES level "
                     "(e.g. the '7153f' the broker cited).",
            )
            es_override = st.number_input(
                "ES futures level",
                value=live_es,
                step=0.25,
                format="%.2f",
                key="ladder_es_override",
                disabled=not use_es_override,
            )
        with cc2:
            use_specific_strike = st.checkbox(
                "Show one specific strike",
                value=False,
                key="ladder_use_strike",
                help="Tick to price a single strike instead of the 7-strike ladder.",
            )
            specific_strike = st.number_input(
                "Strike",
                value=float(round(live_es / 25) * 25),
                step=1.0,
                format="%.0f",
                key="ladder_specific_strike",
                disabled=not use_specific_strike,
            )
        with cc3:
            quoted_bid = st.number_input(
                "Quoted bid (optional)",
                value=0.0,
                step=0.01,
                format="%.2f",
                key="ladder_quoted_bid",
                help="Type the counterparty's bid here to compare with the pricer's bid.",
            )
            quoted_ask = st.number_input(
                "Quoted ask (optional)",
                value=0.0,
                step=0.01,
                format="%.2f",
                key="ladder_quoted_ask",
            )

    es_arg     = float(es_override)      if use_es_override else None
    strike_arg = float(specific_strike)  if use_specific_strike else None

    try:
        result = build_ladder(snap, expiry, float(width_bps),
                              es_override=es_arg, specific_strike=strike_arg)
    except Exception as exc:
        st.error(f"Ladder error: {exc}")
        return

    # Banner if ES override is active
    if use_es_override:
        st.info(
            f"Pricing against ES = {es_arg:,.2f} "
            f"(live {live_es:,.2f}, delta {es_arg - live_es:+,.2f}).  "
            f"Implied spot = {result['spot_eff']:,.2f}."
        )

    # Table
    rows = result["rows"]
    df = pd.DataFrame({
        "Strike": [f"{r['strike']:,.0f}" for r in rows],
        "Bid":    [f"{r['bid']:+.2f}" for r in rows],
        "Mid":    [f"{r['mid']:+.2f}" for r in rows],
        "Ask":    [f"{r['ask']:+.2f}" for r in rows],
        "% Fwd":  [f"{r['pct_fwd']:+.2f}%" for r in rows],
    })
    st.dataframe(df, hide_index=True, width="stretch")

    # Quote-comparison block (only when single strike + at least one side typed)
    if use_specific_strike and (quoted_bid > 0 or quoted_ask > 0):
        row = rows[0]
        ours_bid, ours_mid, ours_ask = row["bid"], row["mid"], row["ask"]
        cp_mid = (quoted_bid + quoted_ask) / 2 if (quoted_bid and quoted_ask) else (quoted_bid or quoted_ask)
        delta_mid = cp_mid - ours_mid
        bps_per_side = (delta_mid / result["F"]) * 10000

        st.markdown("**Quote vs pricer**")
        st.markdown(
            f"""
| | Pricer | Counterparty | Δ |
|---|---|---|---|
| Bid | `{ours_bid:+.2f}` | `{quoted_bid:+.2f}` | `{quoted_bid - ours_bid:+.2f}` |
| Mid | `{ours_mid:+.2f}` | `{cp_mid:+.2f}` | `{delta_mid:+.2f}` |
| Ask | `{ours_ask:+.2f}` | `{quoted_ask:+.2f}` | `{quoted_ask - ours_ask:+.2f}` |
"""
        )
        st.caption(f"Mid difference: {delta_mid:+.2f} pts ≈ {bps_per_side:+.1f} bps of forward")

    # Footer
    st.markdown(
        f"`F={result['F']:,.2f}  T={result['T']:.2f}y  "
        f"r={result['r']*100:.2f}%  b={result['b']*10000:.0f}bp  "
        f"q={result['q']*100:.2f}%  DF={result['DF']:.4f}  "
        f"ES={result['es_used']:,.2f}`"
    )


# ---------------------------------------------------------------------------
# Tab 2 — Revcon
# ---------------------------------------------------------------------------

def render_revcon(snap):
    from pricer.revcon import calc_revcon, format_revcon

    expiry_codes = list(snap.expiries.keys()) if snap else []

    with st.form("revcon_form"):
        c1, c2, c3 = st.columns(3)
        with c1:
            expiry = st.selectbox("Expiry", expiry_codes, key="rev_expiry")
            strike = st.number_input("Strike", value=7200.0, step=100.0, format="%.0f")
            combo = st.number_input("Combo mid", value=20.0, step=0.5, format="%.2f")
        with c2:
            es = st.number_input("ES futures level", value=float(snap.es_future if snap else 7160), step=1.0, format="%.2f")
            basis = st.number_input("Basis (ES − SPX)", value=float(snap.es_basis if snap else 40), step=1.0, format="%.2f")
            side = st.radio("Side", ["REV", "CONV"], horizontal=True)
        with c3:
            lots = st.number_input("Lots", value=1, min_value=1, step=1)
            basis_confirmed = st.checkbox(
                "Basis confirmed in chat?",
                help="Always confirm basis in writing before printing.",
            )

        submitted = st.form_submit_button("Calculate")

    if not submitted:
        return

    if not basis_confirmed:
        st.warning("Confirm the agreed basis in your chat before pricing.")
        return

    if snap is None:
        st.error("No live data — cannot price.")
        return

    if _yearend_flag(expiry, snap):
        st.warning("⚠ Year-end expiry — verify borrow assumption before trading.")

    try:
        result = calc_revcon(snap, expiry, strike, combo, es, basis, side, lots)
    except Exception as exc:
        st.error(f"Revcon error: {exc}")
        return

    verdict = result["verdict"]
    verdict_color = {"LIFT": "green", "PASS": "red", "FLAT": "gray"}[verdict]

    st.markdown(
        f"<div style='background:{verdict_color};color:white;padding:1rem;"
        f"border-radius:6px;font-size:2rem;font-family:monospace;text-align:center'>"
        f"{verdict}</div>",
        unsafe_allow_html=True,
    )

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Breakdown**")
        st.markdown(
            f"""
| Field | Value |
|---|---|
| F_combo | `{result['f_combo']:,.2f}` |
| F_hedge | `{result['f_hedge']:,.2f}` |
| Edge | `{result['edge_pts']:+.2f} pts` |
| P&L $ | `${result['pnl_dollars']:+,.0f}` |
| bps ann | `{result['bps_ann']:+.1f} bps` |
"""
        )
    with c2:
        if _log_quote(snap, "revcon", {
            "expiry": expiry, "strike": strike, "combo": combo,
            "es": es, "basis": basis, "side": side, "lots": lots,
        }, result):
            st.caption("Quote logged.")


# ---------------------------------------------------------------------------
# Tab 3 — Reverse borrow
# ---------------------------------------------------------------------------

def render_reverse(snap):
    from pricer.reverse import calc_reverse

    expiry_codes = list(snap.expiries.keys()) if snap else []

    with st.form("reverse_form"):
        c1, c2 = st.columns(2)
        with c1:
            expiry = st.selectbox("Expiry", expiry_codes, key="rev_borrow_expiry")
            strike = st.number_input("Strike", value=7200.0, step=100.0, format="%.0f", key="rb_strike")
        with c2:
            combo_mid_mkt = st.number_input("Street combo mid", value=22.5, step=0.5, format="%.2f", key="rb_combo")
        submitted = st.form_submit_button("Solve")

    if not submitted:
        return

    if snap is None:
        st.error("No live data — cannot solve.")
        return

    try:
        result = calc_reverse(snap, expiry, strike, combo_mid_mkt)
    except Exception as exc:
        st.error(f"Solver error: {exc}")
        return

    sig = result["signal"]
    sig_color = "green" if sig == "RICH" else "orange"

    st.markdown(
        f"<div style='background:{sig_color};color:white;padding:0.8rem;"
        f"border-radius:6px;font-size:1.5rem;font-family:monospace;text-align:center'>"
        f"{sig}  ({result['delta_bps']:+.0f} bps vs your model)</div>",
        unsafe_allow_html=True,
    )

    st.markdown(
        f"""
| | |
|---|---|
| Implied forward | `{result['f_implied']:,.2f}` |
| Implied (r+b) | `{result['r_plus_b']*100:.4f}%` |
| Implied borrow | `{result['borrow_bps']:.0f} bps` |
| Your assumption | `{result['b_model_bps']:.0f} bps` |
| Delta | `{result['delta_bps']:+.0f} bps` |
"""
    )

    # Side panel: implied curve across all expiries
    with st.expander("Implied borrow curve (all expiries)"):
        import pandas as pd
        rows = []
        for code, settle in snap.expiries.items():
            try:
                r = calc_reverse(snap, code, strike, combo_mid_mkt)
                rows.append({
                    "Expiry": code,
                    "Settle": str(settle),
                    "Impl Fwd": f"{r['f_implied']:,.2f}",
                    "Impl Borrow": f"{r['borrow_bps']:.0f} bps",
                    "Model": f"{r['b_model_bps']:.0f} bps",
                    "Delta": f"{r['delta_bps']:+.0f} bps",
                    "Signal": r["signal"],
                })
            except Exception:
                pass
        if rows:
            st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")


# ---------------------------------------------------------------------------
# Tab 4 — Quote Run (generate your own runs in broker format)
# ---------------------------------------------------------------------------

def render_run(snap):
    from pricer.ladder import build_ladder
    from pricer import math as m
    import pandas as pd

    defs = _load_defaults()
    expiry_codes = list(snap.expiries.keys()) if snap else []

    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        default_width = defs.get("width", {}).get("dealer_to_dealer", 10)
        width_bps = st.slider(
            "Width (bps of forward)", 5, 25, int(default_width), 1, key="run_width",
            help="Total bid-ask in bps of the forward. 10 = D2D, 15 = client.",
        )
    with c2:
        strike_mode = st.radio(
            "Strike for points quote",
            ["ATM forward (round 100)", "ES level (round 25)"],
            key="run_strike_mode",
            help="Choose what strike to use for the points-format quote.",
        )
    with c3:
        st.caption(
            "Generates your run across every expiry. Two formats: "
            "**% of spot** (101.9275% style) and **points** (121.52/121.78 style). "
            "Copy the code blocks straight into your chat."
        )

    if not snap:
        st.info("Waiting for Bloomberg connection…")
        return

    now_str = datetime.now().strftime("%H:%M:%S")
    rows = []
    text_pct = [
        f"SPX Combo Run — {now_str} NY",
        f"ES {snap.es_future:,.2f}  SPX {snap.spot:,.2f}  basis {snap.es_basis:+.2f}",
        "",
    ]
    text_pts = [
        f"SPX Combo Run (points) — {now_str} NY",
        f"ES {snap.es_future:,.2f}  SPX {snap.spot:,.2f}  basis {snap.es_basis:+.2f}",
        "",
    ]

    half_spread = (width_bps / 2) / 10000  # half-width as fraction of forward

    for code in expiry_codes:
        try:
            res = build_ladder(snap, code, float(width_bps))
            F, S, T, r, b, q, DF = res["F"], snap.spot, res["T"], res["r"], res["b"], res["q"], res["DF"]

            # % of spot format
            FoverS = F / S
            FoS_bid = FoverS * (1 - half_spread)
            FoS_ask = FoverS * (1 + half_spread)

            # Points format — pick strike per user's choice
            if strike_mode.startswith("ATM"):
                K = round(F / 100) * 100
            else:
                K = round(snap.es_future / 25) * 25

            mid = m.combo_mid(F, K, r, T)
            eps = m.edge_per_side(width_bps, F, r, T)
            bid_pts, ask_pts = mid - eps, mid + eps

            rows.append({
                "Expiry":   code,
                "F":        f"{F:,.2f}",
                "F/S Bid":  f"{FoS_bid*100:.4f}%",
                "F/S Mid":  f"{FoverS*100:.4f}%",
                "F/S Ask":  f"{FoS_ask*100:.4f}%",
                "K":        f"{K:,.0f}",
                "Pts Bid":  f"{bid_pts:+.2f}",
                "Pts Mid":  f"{mid:+.2f}",
                "Pts Ask":  f"{ask_pts:+.2f}",
                "Borrow":   f"{b*10000:.0f}bp",
            })

            text_pct.append(
                f"SPX {code} combo  {FoS_bid*100:.4f}% / {FoS_ask*100:.4f}%  "
                f"({snap.es_future:,.0f}f)"
            )
            text_pts.append(
                f"SPX {code} combo K={K:,.0f}  {bid_pts:+.2f} / {ask_pts:+.2f}  "
                f"({snap.es_future:,.0f}f)"
            )
        except Exception as exc:
            rows.append({"Expiry": code, "F": f"ERR: {exc}"})

    st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")

    cc1, cc2 = st.columns(2)
    with cc1:
        st.markdown("**Copy — % of spot format**")
        st.code("\n".join(text_pct), language="text")
    with cc2:
        st.markdown("**Copy — points format**")
        st.code("\n".join(text_pts), language="text")


# ---------------------------------------------------------------------------
# Quote logging (Phase 3)
# ---------------------------------------------------------------------------

def _log_quote(snap, tab: str, inputs: dict, result: dict) -> bool:
    try:
        import pandas as pd
        import hashlib, json

        log_dir = Path(__file__).parent / "logs"
        log_dir.mkdir(exist_ok=True)
        log_file = log_dir / "quotes.parquet"

        snap_hash = hashlib.md5(
            f"{snap.spot}{snap.sofr}{snap.timestamp}".encode()
        ).hexdigest()[:8]

        record = {
            "timestamp": [datetime.now().isoformat()],
            "snap_hash": [snap_hash],
            "tab": [tab],
            "inputs": [json.dumps(inputs)],
            "output": [json.dumps({k: float(v) if isinstance(v, (int, float)) else v
                                   for k, v in result.items()})],
        }
        new_df = pd.DataFrame(record)

        if log_file.exists():
            old_df = pd.read_parquet(log_file)
            combined = pd.concat([old_df, new_df], ignore_index=True)
        else:
            combined = new_df

        combined.to_parquet(log_file, index=False)
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if not CONFIG_PATH.exists():
        st.error(
            f"Config not found at `{CONFIG_PATH}`.\n\n"
            "Run `python discover.py` first to map your Excel cells."
        )
        return

    # Auto-refresh every 5 seconds
    try:
        from streamlit_autorefresh import st_autorefresh
        tick = st_autorefresh(interval=5000, key="autorefresh")
    except ImportError:
        tick = int(time.time() // 5)

    snap, err = _read_snap(tick)
    data_ok = render_top_bar(snap, err, tick)

    st.divider()

    tab1, tab2, tab3, tab4 = st.tabs(["Ladder", "Revcon", "Reverse Borrow", "Run"])

    with tab1:
        if data_ok and snap:
            render_ladder(snap)
        elif not data_ok:
            st.error("Stale data — refresh Bloomberg.")

    with tab2:
        if snap:
            render_revcon(snap)

    with tab3:
        if snap:
            render_reverse(snap)

    with tab4:
        if snap:
            render_run(snap)


if __name__ == "__main__":
    main()

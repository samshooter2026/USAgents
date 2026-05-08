"""
SPX Combo & Revcon Pricer — Streamlit UI
Run: streamlit run app.py
"""

from __future__ import annotations

import logging
import time
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

_NY = ZoneInfo("America/New_York")

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
    @keyframes flash-blue {
        0%   { background: #add8e6; color: #000; border-radius: 3px; }
        100% { background: transparent; color: inherit; }
    }
    @keyframes flash-pink {
        0%   { background: #ffb6c1; color: #000; border-radius: 3px; }
        100% { background: transparent; color: inherit; }
    }
    .num-up   { animation: flash-blue 1.8s ease-out 1; }
    .num-down { animation: flash-pink 1.8s ease-out 1; }
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


def _fix_bloomberg_bdp() -> str:
    """Clear and re-enter SPX + ES BDP formulas to force a fresh Bloomberg evaluation.
    Returns a status message."""
    import xlwings as xw
    import yaml
    cfg = yaml.safe_load(open(CONFIG_PATH))
    wb  = xw.Book(cfg["workbook"])
    app = wb.app
    sh  = wb.sheets["BBG Data"]

    spot_cell = cfg["spot"]["cell"]
    es_cell   = cfg["es_future"]["cell"]

    # Clear + re-enter forces Bloomberg to issue a new data request
    for cell, formula in [
        (spot_cell, '=BDP("SPX Index","PX_LAST")'),
        (es_cell,   '=BDP("ES1 Index","PX_LAST")'),
    ]:
        sh.range(cell).clear_contents()
        sh.range(cell).formula = formula

    app.calculate()
    time.sleep(3)
    app.calculate()

    spx = sh.range(spot_cell).value
    es  = sh.range(es_cell).value
    st.cache_data.clear()
    return f"BDP restored. SPX={spx:,.2f}  ES={es:,.2f}"


# ---------------------------------------------------------------------------
# Sticky top bar
# ---------------------------------------------------------------------------

def render_top_bar(snap, err: str | None, tick: int):
    defs = _load_defaults()
    col1, col2, col3, col4, col5 = st.columns([2, 2, 2, 1, 1])
    with col1:
        if snap:
            age = (datetime.now() - snap.timestamp).total_seconds()
            color = _snap_age_color(snap)
            ny_ts = snap.timestamp.astimezone(_NY)
            st.markdown(
                f"**Snapshot** &nbsp; "
                f"<span style='color:{color};font-family:monospace'>"
                f"{ny_ts:%H:%M:%S} NY ({age:.0f}s ago)</span>",
                unsafe_allow_html=True,
            )
        else:
            st.markdown("**Snapshot** &nbsp; <span style='color:red'>NOT CONNECTED</span>",
                        unsafe_allow_html=True)
    with col2:
        if snap:
            frozen = st.session_state.get("_spx_frozen", False)
            color  = "orange" if frozen else "inherit"
            st.markdown(
                f"**SPX** &nbsp; <span style='font-family:monospace;font-size:1.1rem;color:{color}'>"
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
        if st.button("Refresh"):
            st.cache_data.clear()
            st.rerun()
    with col5:
        if st.button("Fix BDP", help="Restores SPX + ES Bloomberg formulas if they've been overwritten or frozen"):
            with st.spinner("Restoring BDP formulas..."):
                try:
                    msg = _fix_bloomberg_bdp()
                    st.success(msg)
                    st.rerun()
                except Exception as exc:
                    st.error(f"Fix failed: {exc}")

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
# Tab 5 — AXW Calibration Agent
# ---------------------------------------------------------------------------

def render_agent_deployment(snap):
    """AXW calibration agent — log every street print, flag AXW drift."""
    from pricer import calibration as cal
    from pricer.reverse import calc_reverse
    from dataclasses import replace
    import pandas as pd

    st.markdown("### AXW Calibration Agent")
    st.caption(
        f"Logs every street combo print, computes implied borrow, and flags when AXW "
        f"drifts from real flow.  "
        f"Alert rule: ≥{cal.DRIFT_MIN_PRINTS} consecutive prints same-side and "
        f"mean |Δ| > {cal.DRIFT_TOLERANCE_BPS}bp vs current AXW."
    )

    if not snap:
        st.info("Waiting for Bloomberg connection…")
        return

    expiry_codes = list(snap.expiries.keys())
    axw_now_bps  = {code: snap.borrow[code] * 10000 for code in snap.borrow}

    # ── Paste a street print (primary input) ────────────────────────────────
    st.markdown("#### 📋 Paste street print")
    paste_default = (
        ""
        if "cal_paste_text" not in st.session_state
        else st.session_state["cal_paste_text"]
    )
    with st.form("axw_paste_form", clear_on_submit=True):
        text = st.text_area(
            "Paste a one-line print",
            value=paste_default,
            placeholder="20:38:10 SPX Dec26 combo trades 650mm ~101.9275% makes 1.2b",
            height=70,
            key="cal_paste_text_input",
            help="Free-form. Picks up expiry, %-of-spot or pts-bid/ask, "
                 "size, ES level, trade time. ES level is optional for % quotes.",
        )
        paste_source = st.text_input(
            "Source / counterparty (optional)",
            placeholder="JPM / GS / chat / etc.",
            key="cal_paste_source",
        )
        paste_submit = st.form_submit_button("Parse and log →")

    if paste_submit and text.strip():
        try:
            parsed = cal.parse_any(text)
        except ValueError as exc:
            st.error(f"Couldn't parse: {exc}")
        else:
            try:
                # ── AXW direct trade (one or many legs) ────────────────────
                if parsed["kind"] == "axw_direct":
                    logged = []
                    for leg in parsed["legs"]:
                        code = leg["expiry"]
                        if code not in snap.expiries:
                            st.warning(f"Skipping unknown expiry {code}.")
                            continue
                        axw_at_log = axw_now_bps.get(code, 0.0)
                        cal.log_axw_direct(
                            expiry=code,
                            borrow_bps=leg["borrow_bps"],
                            axw_borrow_bps_at_log=axw_at_log,
                            contracts=parsed.get("size"),
                            source=(paste_source or "AXW direct"),
                        )
                        delta = leg["borrow_bps"] - axw_at_log
                        logged.append(f"{code} @ {leg['borrow_bps']:.2f}bp (Δ {delta:+.2f}bp)")
                    if logged:
                        size_str = (
                            f" · {parsed['size']} contracts"
                            if parsed.get("size") else ""
                        )
                        st.success("AXW direct logged: " + " · ".join(logged) + size_str)

                # ── Combo print, % of spot ─────────────────────────────────
                elif parsed.get("format") == "pct":
                    if parsed["expiry"] not in snap.expiries:
                        st.error(
                            f"Parsed expiry '{parsed['expiry']}' not in snapshot. "
                            f"Known: {list(snap.expiries.keys())}"
                        )
                    else:
                        res = cal.implied_borrow_from_pct_print(
                            parsed["fs_pct"], parsed["expiry"], snap,
                        )
                        es_used = parsed.get("es_level") or float(snap.es_future)
                        spx_at_trade = snap.spot + (es_used - snap.es_future)
                        cal.log_print(
                            expiry=parsed["expiry"],
                            strike=0.0,
                            combo_mid=parsed["fs_pct"],
                            es_at_trade=es_used,
                            spx_at_trade=spx_at_trade,
                            notional_mm=parsed.get("size_mm"),
                            source=paste_source or None,
                            implied_borrow_bps=res["borrow_bps"],
                            axw_borrow_bps=res["axw_bps"],
                        )
                        st.success(
                            f"Logged: {parsed['expiry']} @ {parsed['fs_pct']:.4f}% "
                            f"({parsed.get('size_mm', 0):.0f}mm)  "
                            f"→ implied {res['borrow_bps']:.1f}bp vs AXW {res['axw_bps']:.1f}bp "
                            f"({res['delta_bps']:+.1f}bp)"
                        )

                # ── Combo print, points format ─────────────────────────────
                else:
                    st.warning(
                        "Points-format quote detected, but parser doesn't know the strike. "
                        "Use the manual form below to enter strike + combo mid."
                    )
            except Exception as exc:
                st.error(f"Error computing implied borrow: {exc}")

    # ── Manual form (fallback / points-format / specific strike) ────────────
    with st.expander("📝 Manual entry (for points-format prints or to override)"):
        with st.form("axw_log_form", clear_on_submit=True):
            r1c1, r1c2, r1c3 = st.columns(3)
            with r1c1:
                expiry = st.selectbox("Expiry", expiry_codes, key="cal_expiry")
                strike = st.number_input(
                    "Strike", value=float(round(snap.es_future / 25) * 25),
                    step=25.0, format="%.0f",
                )
            with r1c2:
                combo_mid = st.number_input(
                    "Combo mid (points)", value=0.0, step=0.5, format="%.2f",
                    help="The street's combo mid in index points (not %).",
                )
                es_at_trade = st.number_input(
                    "ES at trade time", value=float(snap.es_future),
                    step=0.25, format="%.2f",
                    help="Use the ES level the print was struck against (e.g. '7338f').",
                )
            with r1c3:
                notional_mm = st.number_input(
                    "Notional (mm)", value=0, step=10, min_value=0,
                )
                source = st.text_input(
                    "Source / counterparty", placeholder="JPM / GS / chat / etc.",
                    key="cal_manual_source",
                )

            submitted = st.form_submit_button("Log print →")

        if submitted:
            if combo_mid == 0.0:
                st.error("Combo mid is required.")
            else:
                try:
                    spot_at_trade = snap.spot + (es_at_trade - snap.es_future)
                    snap_at_trade = replace(snap, spot=spot_at_trade, es_future=es_at_trade)

                    res = calc_reverse(snap_at_trade, expiry, strike, combo_mid)
                    cal.log_print(
                        expiry=expiry,
                        strike=strike,
                        combo_mid=combo_mid,
                        es_at_trade=es_at_trade,
                        spx_at_trade=spot_at_trade,
                        notional_mm=notional_mm or None,
                        source=source or None,
                        implied_borrow_bps=res["borrow_bps"],
                        axw_borrow_bps=res["b_model_bps"],
                    )
                    st.success(
                        f"Logged: {expiry} K={strike:.0f} combo={combo_mid:+.2f} "
                        f"→ implied {res['borrow_bps']:.1f}bp vs AXW {res['b_model_bps']:.1f}bp "
                        f"({res['delta_bps']:+.1f}bp · {res['signal']})"
                    )
                except Exception as exc:
                    st.error(f"Error: {exc}")

    # ── Drift dashboard ─────────────────────────────────────────────────────
    df = cal.read_log()
    drift = cal.detect_drift(df, axw_now_bps)

    st.markdown("#### Drift status (vs current AXW)")
    if df.empty:
        st.info("No prints logged yet. Use the form above to start building the calibration history.")
    else:
        rows = []
        any_alert = False
        for code in expiry_codes:
            d = drift.get(code)
            if d is None:
                rows.append({
                    "Expiry":        code,
                    "AXW now":       f"{axw_now_bps[code]:.1f}bp",
                    "Recent prints": 0,
                    "Mean Δ":        "—",
                    "Status":        "— no prints —",
                    "Recommend":     "—",
                })
                continue
            n, mean_d, alert = d["n_prints"], d["mean_delta_bps"], d["alert"]
            if alert:
                status = f"🔴 {alert}"
                any_alert = True
            elif abs(mean_d) > cal.DRIFT_TOLERANCE_BPS:
                status = f"🟡 mean {mean_d:+.1f}bp ({n} prints)"
            else:
                status = f"🟢 within {cal.DRIFT_TOLERANCE_BPS}bp ({n} prints)"
            rows.append({
                "Expiry":        code,
                "AXW now":       f"{axw_now_bps[code]:.1f}bp",
                "Recent prints": n,
                "Mean Δ":        f"{mean_d:+.1f}bp",
                "Status":        status,
                "Recommend":     f"{d['rec_borrow_bps']:.1f}bp" if alert else "—",
            })
        st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")

        if any_alert:
            st.warning(
                "One or more expiries have AXW drift exceeding tolerance. "
                "Consider overriding the green Excel borrow cell to the **Recommend** value, "
                "or flag the expiry on your run."
            )

    # ── Trader-style analysis ──────────────────────────────────────────────
    if not df.empty:
        st.markdown("#### 🧠 Analysis")

        market_line = cal.market_summary(df, axw_now_bps, expiry_codes)
        st.markdown(f"_{market_line}_")

        narratives = cal.narrate_drift(df, axw_now_bps, expiry_codes)
        if narratives:
            for code in expiry_codes:
                if code not in narratives:
                    continue
                d = drift.get(code, {})
                # Match the same colour band as the dashboard
                if d.get("alert"):
                    icon = "🔴"
                elif abs(d.get("mean_delta_bps", 0.0)) > cal.DRIFT_TOLERANCE_BPS:
                    icon = "🟡"
                else:
                    icon = "🟢"
                st.markdown(f"{icon} **{code}** — {narratives[code]}")

    # ── Recent prints log ──────────────────────────────────────────────────
    if not df.empty:
        st.markdown("#### Recent prints")
        f1, f2, f3 = st.columns([2, 1, 1])
        with f1:
            sel_expiries = st.multiselect(
                "Filter by expiry", expiry_codes, default=expiry_codes,
                key="cal_filter_exp",
            )
        with f2:
            window_choice = st.selectbox(
                "Window", [1, 7, 30, 90, 365], index=2,
                format_func=lambda d: f"{d}d", key="cal_window",
            )
        with f3:
            st.metric("Total logged", len(df))

        view = cal.read_log(since_days=window_choice)
        view = view[view["expiry"].isin(sel_expiries)].head(100).copy()
        view["axw_now_bps"]     = view["expiry"].map(axw_now_bps)
        view["delta_today_bps"] = view["implied_borrow_bps"] - view["axw_now_bps"]

        display = view[[
            "timestamp", "expiry", "strike", "combo_mid", "es_at_trade",
            "notional_mm", "implied_borrow_bps", "axw_borrow_bps",
            "axw_now_bps", "delta_today_bps", "source",
        ]].rename(columns={
            "timestamp":          "Time",
            "expiry":             "Exp",
            "strike":             "K",
            "combo_mid":          "Combo",
            "es_at_trade":        "ES@trade",
            "notional_mm":        "mm",
            "implied_borrow_bps": "Impl bp",
            "axw_borrow_bps":     "AXW@trade",
            "axw_now_bps":        "AXW now",
            "delta_today_bps":    "Δ today",
            "source":             "Source",
        })
        st.dataframe(display, hide_index=True, width="stretch")


# ---------------------------------------------------------------------------
# Tab 4 — Quote Run (generate your own runs in broker format)
# ---------------------------------------------------------------------------

def render_run(snap):
    from pricer.ladder import build_ladder
    from pricer import math as m
    import pandas as pd

    defs = _load_defaults()
    expiry_codes = list(snap.expiries.keys()) if snap else []

    c1, c2, c3, c4 = st.columns([1, 1, 1, 2])
    with c1:
        default_width = defs.get("width", {}).get("dealer_to_dealer", 25)
        width_bps = st.slider(
            "Default width (bps of forward)", 5, 100, int(default_width), 1, key="run_width",
            help="Default for all expiries. Each expiry can be overridden individually below.",
        )
    with c2:
        strike_mode = st.radio(
            "Strike for points quote",
            ["ATM forward (round 100)", "ES level (round 25)"],
            key="run_strike_mode",
            help="Choose what strike to use for the points-format quote.",
        )
    with c3:
        live_es = float(snap.es_future) if snap else 0.0
        use_es_override = st.checkbox(
            "Override ES level",
            value=False,
            key="run_use_es_override",
            help="Tick to reprice the whole run against a specific ES level "
                 "(e.g. matching a broker print at 7329f).",
        )
        es_override = st.number_input(
            "ES futures level",
            value=live_es,
            step=0.25,
            format="%.2f",
            key="run_es_override",
            disabled=not use_es_override,
        )
    with c4:
        per_expiry_defaults: dict[str, int] = defs.get("per_expiry_widths", {}) or {}
        bcol1, bcol2 = st.columns(2)
        with bcol1:
            if st.button("Set all = default", help="Set every per-expiry slider to the default width above"):
                for code in expiry_codes:
                    st.session_state[f"run_width_{code}"] = int(width_bps)
                st.rerun()
        with bcol2:
            if st.button("Reset to per-expiry", help="Reset each slider to its YAML-configured default"):
                for code in expiry_codes:
                    if code in per_expiry_defaults:
                        st.session_state[f"run_width_{code}"] = int(per_expiry_defaults[code])
                st.rerun()
        st.caption(
            "**% of spot** and **points** formats. Per-expiry sliders persist across refreshes."
        )

    if not snap:
        st.info("Waiting for Bloomberg connection…")
        return

    es_used = float(es_override) if use_es_override else float(snap.es_future)
    es_arg = float(es_override) if use_es_override else None
    spot_used = snap.spot + (es_used - snap.es_future)

    now_str = datetime.now(_NY).strftime("%H:%M:%S")
    header_es = (
        f"ES {es_used:,.2f}"
        + (f" (LIVE {snap.es_future:,.2f})" if use_es_override else "")
    )

    # ---- Per-expiry sliders + run output (side by side) -------------------
    slider_col, run_col = st.columns([1, 3])

    per_expiry_widths: dict[str, int] = {}
    with slider_col:
        st.markdown("**Per-expiry width**")

        # Auto-apply per-expiry YAML defaults on first load AND whenever the
        # YAML changes. Within a stable YAML version, user adjustments persist
        # across the 5-second auto-refresh.
        import hashlib
        yaml_sig = hashlib.md5(
            str(sorted(per_expiry_defaults.items())).encode()
        ).hexdigest()
        if st.session_state.get("_run_widths_yaml_sig") != yaml_sig:
            for code in expiry_codes:
                st.session_state[f"run_width_{code}"] = int(
                    per_expiry_defaults.get(code, width_bps)
                )
            st.session_state["_run_widths_yaml_sig"] = yaml_sig

        for code in expiry_codes:
            key = f"run_width_{code}"
            if key not in st.session_state:
                st.session_state[key] = int(per_expiry_defaults.get(code, width_bps))
            per_expiry_widths[code] = st.slider(
                code,
                min_value=5, max_value=100, step=1,
                key=key,
                format="%dbp",
            )

    # ---- Build all calculations using per-expiry widths -------------------
    rows = []
    pct_floats = []
    text_pct = [
        f"SPX Combo Run — {now_str} NY",
        f"{header_es}  SPX {spot_used:,.2f}  basis {snap.es_basis:+.2f}",
        "",
    ]
    text_pts = [
        f"SPX Combo Run (points) — {now_str} NY",
        f"{header_es}  SPX {spot_used:,.2f}  basis {snap.es_basis:+.2f}",
        "",
    ]

    for code in expiry_codes:
        try:
            w = float(per_expiry_widths[code])
            half_spread = (w / 2) / 10000

            res = build_ladder(snap, code, w, es_override=es_arg)
            F, S, T, r, b, q, DF = res["F"], res["spot_eff"], res["T"], res["r"], res["b"], res["q"], res["DF"]

            FoverS = F / S
            FoS_bid = FoverS * (1 - half_spread)
            FoS_ask = FoverS * (1 + half_spread)

            if strike_mode.startswith("ATM"):
                K = round(F / 100) * 100
            else:
                K = round(es_used / 25) * 25

            mid = m.combo_mid(F, K, r, T)
            eps = m.edge_per_side(w, F, r, T)
            bid_pts, ask_pts = mid - eps, mid + eps

            rows.append({
                "Expiry":   code,
                "Width":    f"{int(w)}bp",
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
            pct_floats.append((code, FoS_bid * 100, FoS_ask * 100))

            text_pct.append(
                f"SPX {code} combo  {FoS_bid*100:.4f}% / {FoS_ask*100:.4f}%  "
                f"({es_used:,.0f}f)"
            )
            text_pts.append(
                f"SPX {code} combo K={K:,.0f}  {bid_pts:+.2f} / {ask_pts:+.2f}  "
                f"({es_used:,.0f}f)"
            )
        except Exception as exc:
            rows.append({"Expiry": code, "F": f"ERR: {exc}"})

    # Flash detection: compare current pct values to previous render
    _prev = st.session_state.get("_run_prev_pct", {})
    _curr = {code: (bid, ask) for code, bid, ask in pct_floats}

    def _num_class(code: str, side: str, val: float) -> str:
        prev_bid, prev_ask = _prev.get(code, (val, val))
        prev_val = prev_bid if side == "bid" else prev_ask
        if val > prev_val + 1e-6:
            return "num-up"
        if val < prev_val - 1e-6:
            return "num-down"
        return ""

    # Build styled % of spot HTML block
    hdr1 = text_pct[0]
    hdr2 = text_pct[1]
    pct_html_lines = []
    for code, bid_pct, ask_pct in pct_floats:
        bid_cls = _num_class(code, "bid", bid_pct)
        ask_cls = _num_class(code, "ask", ask_pct)
        bid_str = f"{bid_pct:.4f}%"
        ask_str = f"{ask_pct:.4f}%"
        pct_html_lines.append(
            f"SPX {code} combo &nbsp; "
            f"<span class='{bid_cls}' style='font-size:1.25rem;font-weight:600;'>{bid_str}</span>"
            f" / "
            f"<span class='{ask_cls}' style='font-size:1.25rem;font-weight:600;'>{ask_str}</span>"
            f" &nbsp;<span style='color:#888;'>({es_used:,.0f}f)</span>"
        )

    pct_html = (
        f"<div style='font-family:monospace;font-size:1rem;padding:1rem 1.2rem;"
        f"background:#0e1117;color:#fafafa;border-radius:6px;line-height:2.4;'>"
        f"<div style='font-weight:bold;margin-bottom:0.1rem;'>{hdr1}</div>"
        f"<div style='color:#aaa;font-size:0.85rem;margin-bottom:0.6rem;'>{hdr2}</div>"
        + "<br>".join(pct_html_lines)
        + "</div>"
    )

    st.session_state["_run_prev_pct"] = _curr

    # Right side of the slider/run row: the styled % of spot block
    with run_col:
        st.markdown("**% of spot**")
        st.markdown(pct_html, unsafe_allow_html=True)

    # Full-width: the run dataframe (now includes per-expiry width column)
    st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")

    # Copy blocks below, side by side
    cc1, cc2 = st.columns(2)
    with cc1:
        st.markdown("**Copy — % of spot**")
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

    # Value-staleness detection: warn if SPX hasn't moved across reads
    if snap:
        prev_spx        = st.session_state.get("_prev_spx")
        frozen_since    = st.session_state.get("_spx_frozen_since")
        now_local       = datetime.now()

        if prev_spx is not None and abs(snap.spot - prev_spx) < 0.001:
            if frozen_since is None:
                st.session_state["_spx_frozen_since"] = now_local
            frozen_secs = (now_local - (frozen_since or now_local)).total_seconds()
            st.session_state["_spx_frozen"] = frozen_secs > 60
        else:
            st.session_state["_spx_frozen_since"] = None
            st.session_state["_spx_frozen"] = False

        st.session_state["_prev_spx"] = snap.spot
    else:
        st.session_state["_spx_frozen"] = False

    data_ok = render_top_bar(snap, err, tick)

    if st.session_state.get("_spx_frozen"):
        frozen_secs = (datetime.now() - st.session_state["_spx_frozen_since"]).total_seconds()
        st.warning(
            f"SPX has been frozen at {snap.spot:,.2f} for {frozen_secs:.0f}s — "
            "Bloomberg RTD may have stalled. "
            "Press **Ctrl+Alt+F9** in Excel, or click **Fix BDP** above."
        )

    st.divider()

    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        ["Run", "Ladder", "Revcon", "Reverse Borrow", "Agent Deployment"]
    )

    with tab1:
        if snap:
            render_run(snap)

    with tab2:
        if data_ok and snap:
            render_ladder(snap)
        elif not data_ok:
            st.error("Stale data — refresh Bloomberg.")

    with tab3:
        if snap:
            render_revcon(snap)

    with tab4:
        if snap:
            render_reverse(snap)

    with tab5:
        render_agent_deployment(snap)


if __name__ == "__main__":
    main()

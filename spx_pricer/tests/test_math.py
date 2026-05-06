"""
Unit tests for pricer/math.py — pure formula correctness.
All expected values derived analytically from the documented formulas.
"""

import math
import pytest
from pricer.math import (
    discount_factor,
    pv_divs,
    forward,
    combo_mid,
    edge_per_side,
    ladder_strikes,
    revcon_pnl,
    reverse_borrow,
    sensitivity_sofr,
    sensitivity_div,
    SPX_MULTIPLIER,
)


# ---------------------------------------------------------------------------
# Building blocks
# ---------------------------------------------------------------------------

class TestDiscountFactor:
    def test_zero_time(self):
        assert discount_factor(0.036, 0.0) == pytest.approx(1.0)

    def test_known_value(self):
        # exp(-0.036 * 1.0) = 0.96464
        assert discount_factor(0.036, 1.0) == pytest.approx(math.exp(-0.036), rel=1e-9)

    def test_dec26_approx(self):
        # r=3.63%, T=0.65y → DF ≈ 0.9767
        df = discount_factor(0.0363, 0.65)
        assert 0.976 < df < 0.977


class TestPvDivs:
    def test_zero_yield(self):
        assert pv_divs(7109, 0.0, 0.65) == 0.0

    def test_zero_time(self):
        assert pv_divs(7109, 0.0125, 0.0) == 0.0

    def test_linear(self):
        # S×q×T = 7109 × 0.0125 × 0.65 = 57.76
        assert pv_divs(7109, 0.0125, 0.65) == pytest.approx(57.763, rel=1e-4)


class TestForward:
    def test_zero_rates_zero_divs(self):
        # F = S when r=b=q=0
        assert forward(7109, 0.0, 0.0, 0.0, 1.0) == pytest.approx(7109.0)

    def test_higher_rate_raises_forward(self):
        f1 = forward(7109, 0.0125, 0.036, 0.0012, 0.65)
        f2 = forward(7109, 0.0125, 0.040, 0.0012, 0.65)
        assert f2 > f1

    def test_higher_divs_lower_forward(self):
        f1 = forward(7109, 0.0125, 0.036, 0.0012, 0.65)
        f2 = forward(7109, 0.0200, 0.036, 0.0012, 0.65)
        assert f2 < f1

    def test_higher_borrow_raises_forward(self):
        f1 = forward(7109, 0.0125, 0.036, 0.0012, 0.65)
        f2 = forward(7109, 0.0125, 0.036, 0.0050, 0.65)
        assert f2 > f1


class TestComboMid:
    def test_atm_near_zero_when_fwd_equals_strike(self):
        # If K = F, combo ≈ 0 (exactly zero before discounting; DF × 0 = 0)
        mid = combo_mid(F=7200.0, K=7200.0, r=0.036, T=0.65)
        assert mid == pytest.approx(0.0, abs=1e-10)

    def test_itm_call_positive(self):
        mid = combo_mid(F=7221, K=7000, r=0.0363, T=0.65)
        assert mid > 0

    def test_otm_call_negative(self):
        mid = combo_mid(F=7221, K=7400, r=0.0363, T=0.65)
        assert mid < 0

    def test_step_equals_df_per_100pts(self):
        """Mid step between K and K+100 equals DF × 100."""
        F, r, T = 7221.0, 0.0363, 0.65
        mid_7200 = combo_mid(F, 7200, r, T)
        mid_7300 = combo_mid(F, 7300, r, T)
        step = mid_7200 - mid_7300
        df = discount_factor(r, T)
        assert step == pytest.approx(df * 100, rel=1e-9)

    def test_matches_pdf_dec26_structure(self):
        """
        Forward backed out from PDF Dec26 table: F=7221.26, DF=0.9765.
        Using those, mid at K=7200 should match PDF's 20.76 within $0.50.
        """
        F = 7221.26
        r, T = 0.0363, 0.65
        mid = combo_mid(F, 7200, r, T)
        # PDF shows 20.76 — but that was computed with slightly different F
        # Our F=7225 from stated inputs gives a different mid.
        # Here we use the table-implied F, so it should match.
        assert mid == pytest.approx(20.76, abs=0.50)


class TestEdgePerSide:
    def test_wider_bps_gives_wider_spread(self):
        e1 = edge_per_side(10, 7221, 0.0363, 0.65)
        e2 = edge_per_side(15, 7221, 0.0363, 0.65)
        assert e2 > e1

    def test_bid_below_ask(self):
        eps = edge_per_side(10, 7221, 0.0363, 0.65)
        mid = combo_mid(7221, 7200, 0.0363, 0.65)
        assert mid - eps < mid < mid + eps

    def test_spread_constant_across_strikes(self):
        """The full bid-ask spread is the same for all strikes (only F, DF matter)."""
        F, r, T = 7221.0, 0.0363, 0.65
        eps = edge_per_side(10, F, r, T)
        for K in [7000, 7100, 7200, 7300, 7400]:
            mid = combo_mid(F, K, r, T)
            spread = (mid + eps) - (mid - eps)
            assert spread == pytest.approx(2 * eps, rel=1e-9)


class TestLadderStrikes:
    def test_seven_strikes(self):
        assert len(ladder_strikes(7221.26)) == 7

    def test_centred_on_atm(self):
        strikes = ladder_strikes(7221.26)
        # ATM = 7200 (nearest 100)
        assert 7200 in strikes

    def test_symmetric_around_atm(self):
        strikes = ladder_strikes(7221.26)
        atm = 7200.0
        assert atm - 300 in strikes
        assert atm + 300 in strikes

    def test_spacing(self):
        strikes = ladder_strikes(7221.26, spacing=100.0)
        diffs = [strikes[i+1] - strikes[i] for i in range(len(strikes)-1)]
        assert all(d == pytest.approx(100.0) for d in diffs)


# ---------------------------------------------------------------------------
# Revcon P&L — tests against PDF §6 exact worked example
# ---------------------------------------------------------------------------

class TestRevconPnl:
    """PDF §6: K=7200, combo=20, r=3.63%, T=0.65, ES=7160, basis=40, REV, 100 lots, spot=7109."""

    INPUTS = dict(
        K=7200.0, combo=20.0, r=0.0363, T=0.65,
        es_future=7160.0, basis=40.0, side="REV",
        lots=100, spot=7109.0,
    )

    def test_f_combo(self):
        res = revcon_pnl(**self.INPUTS)
        assert res["f_combo"] == pytest.approx(7220.48, abs=0.02)

    def test_f_hedge(self):
        res = revcon_pnl(**self.INPUTS)
        assert res["f_hedge"] == pytest.approx(7200.00, abs=0.01)

    def test_edge_pts(self):
        res = revcon_pnl(**self.INPUTS)
        assert res["edge_pts"] == pytest.approx(-20.48, abs=0.02)

    def test_pnl_dollars(self):
        # PDF shows −$204,800 but rounds the intermediate 20.4775 → 20.48 before multiplying.
        # Our implementation carries full precision → −$204,775. Tolerance wiened to ±$30.
        res = revcon_pnl(**self.INPUTS)
        assert res["pnl_dollars"] == pytest.approx(-204800, abs=30)

    def test_verdict_pass(self):
        res = revcon_pnl(**self.INPUTS)
        assert res["verdict"] == "PASS"

    def test_bps_ann(self):
        # bps_ann = (-20.48 / 7109) / 0.65 × 10000 ≈ −44.3 bps
        res = revcon_pnl(**self.INPUTS)
        assert res["bps_ann"] == pytest.approx(-44.3, abs=0.5)

    def test_conv_flips_sign(self):
        """CONV is the mirror: edge sign flips, P&L flips."""
        inp = {**self.INPUTS, "side": "CONV"}
        res = revcon_pnl(**inp)
        assert res["edge_pts"] == pytest.approx(20.48, abs=0.02)
        assert res["pnl_dollars"] == pytest.approx(204800, abs=30)
        assert res["verdict"] == "LIFT"

    def test_lift_verdict(self):
        """Negative combo (ITM on put side) → F_combo < F_hedge → positive edge → LIFT."""
        # F_combo = 7200 + (−20) × exp(rT) ≈ 7179.5 < F_hedge = 7200 → LIFT
        inp = {**self.INPUTS, "combo": -20.0}
        res = revcon_pnl(**inp)
        assert res["verdict"] == "LIFT"

    def test_flat_verdict(self):
        """Near-zero edge → FLAT."""
        # Find combo that gives ~0 bps annualised
        # bps_ann = 0 when edge = 0, i.e. F_hedge = F_combo
        # F_hedge = 7200, so combo s.t. K + combo × exp(rT) = 7200
        # combo = (7200 - 7200) / exp(rT) = 0
        inp = {**self.INPUTS, "combo": 0.0}
        res = revcon_pnl(**inp)
        assert res["verdict"] == "FLAT"


# ---------------------------------------------------------------------------
# Reverse-implied borrow
# ---------------------------------------------------------------------------

class TestReverseBorrow:
    def test_roundtrip(self):
        """Forward borrow in → reverse → same borrow back out."""
        spot, div_yield, r, T = 7109.0, 0.0125, 0.0363, 0.65
        b_in = 0.0012  # 12 bps
        F = forward(spot, div_yield, r, b_in, T)
        mid = combo_mid(F, 7200, r, T)
        result = reverse_borrow(7200, mid, r, T, spot, div_yield)
        assert result["borrow_bps"] == pytest.approx(b_in * 10000, abs=0.1)

    def test_higher_combo_implies_higher_borrow(self):
        """A higher combo mid implies a richer forward → higher implied borrow."""
        kwargs = dict(K=7200, r=0.0363, T=0.65, spot=7109, div_yield=0.0125)
        r1 = reverse_borrow(**kwargs, combo_mid_market=20.0)
        r2 = reverse_borrow(**kwargs, combo_mid_market=25.0)
        assert r2["borrow_bps"] > r1["borrow_bps"]


# ---------------------------------------------------------------------------
# Sensitivities (sign checks only — approximations)
# ---------------------------------------------------------------------------

class TestSensitivities:
    def test_sofr_sensitivity_positive(self):
        """Higher SOFR → higher forward → higher combo. Sensitivity > 0."""
        s = sensitivity_sofr(0.65, 7221, 0.0363)
        assert s > 0

    def test_div_sensitivity_negative(self):
        """Higher divs → lower forward → lower combo. Sensitivity < 0."""
        s = sensitivity_div(0.65, 7109, 0.0363)
        assert s < 0

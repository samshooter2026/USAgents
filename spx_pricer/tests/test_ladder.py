"""
Tests for the ladder module.
Uses the forward implied from the PDF Dec26 table (F=7221.26, DF=0.9765)
rather than the stated-but-rounded inputs, so combo mids match within ±$0.50.
"""

import pytest
from datetime import date, datetime
from unittest.mock import patch, MagicMock

from pricer.math import combo_mid, forward, discount_factor


# PDF Dec26 table-implied parameters
SPOT = 7109.0
R = 0.0363
B = 0.0012
Q = 0.0125
T = 0.65
F_TABLE = 7221.26   # backed out from PDF table mids


# ---------------------------------------------------------------------------
# Structure tests
# ---------------------------------------------------------------------------

class TestLadderStructure:
    def _ladder(self, F=F_TABLE, width_bps=10.0, r=R, T=T):
        from pricer.math import ladder_strikes, edge_per_side
        strikes = ladder_strikes(F)
        eps = edge_per_side(width_bps, F, r, T)
        return [
            {"strike": K, "bid": combo_mid(F, K, r, T) - eps,
             "mid": combo_mid(F, K, r, T),
             "ask": combo_mid(F, K, r, T) + eps}
            for K in strikes
        ]

    def test_seven_rows(self):
        assert len(self._ladder()) == 7

    def test_bid_lt_mid_lt_ask(self):
        for row in self._ladder():
            assert row["bid"] < row["mid"] < row["ask"]

    def test_constant_spread(self):
        rows = self._ladder()
        spreads = [r["ask"] - r["bid"] for r in rows]
        assert all(abs(s - spreads[0]) < 0.01 for s in spreads)

    def test_strike_step_100(self):
        from pricer.math import ladder_strikes
        strikes = ladder_strikes(F_TABLE)
        diffs = [strikes[i+1] - strikes[i] for i in range(len(strikes)-1)]
        assert all(d == pytest.approx(100.0) for d in diffs)

    def test_atm_mid_smallest_absolute(self):
        """The ATM combo (nearest strike to F) should have the smallest |mid|."""
        rows = self._ladder()
        atm_row = min(rows, key=lambda r: abs(r["mid"]))
        assert atm_row["strike"] == 7200.0

    def test_mid_step_equals_df_times_100(self):
        """Mid decreases by DF×100 for each 100pt increase in strike."""
        rows = self._ladder()
        df = discount_factor(R, T)
        for i in range(len(rows) - 1):
            step = rows[i]["mid"] - rows[i+1]["mid"]
            assert step == pytest.approx(df * 100, rel=1e-6)


# ---------------------------------------------------------------------------
# Price proximity to PDF table values
# ---------------------------------------------------------------------------

PDF_DEC26 = [
    (7000, 213.20, 216.06, 218.92),
    (7100, 115.55, 118.41, 121.27),
    (7200,  17.90,  20.76,  23.62),
    (7300, -79.74, -76.88, -74.02),
    (7400, -177.39, -174.53, -171.67),
]


@pytest.mark.parametrize("K, bid_pdf, mid_pdf, ask_pdf", PDF_DEC26)
def test_dec26_mid_within_tolerance(K, bid_pdf, mid_pdf, ask_pdf):
    """
    Combo mid from formula matches PDF table to within ±$0.50.
    Uses table-implied F=7221.26 (not PDF's stated '≈7216' which is rounded).
    """
    mid = combo_mid(F_TABLE, K, R, T)
    assert mid == pytest.approx(mid_pdf, abs=0.50)


PDF_DEC30 = [
    (7700, 158.85, 163.91, 168.97),
    (7800,  73.35,  78.41,  83.47),
    (7900, -12.15,  -7.09,  -2.03),
    (8000, -97.64, -92.58, -87.52),
    (8100, -183.14, -178.08, -173.02),
]

# Dec30 table-implied: mid step = 85.49 → DF = 0.8549; F from K=7900: F = 7900 + (-7.09/0.8549)
R30 = 0.0336
T30 = 4.66
# Back out F from K=7900 mid = -7.09, DF = exp(-0.0336*4.66) = 0.8551
import math as _math
_DF30 = _math.exp(-R30 * T30)
F30 = 7900 + (-7.09 / _DF30)


@pytest.mark.parametrize("K, bid_pdf, mid_pdf, ask_pdf", PDF_DEC30)
def test_dec30_mid_within_tolerance(K, bid_pdf, mid_pdf, ask_pdf):
    """Combo mid matches PDF Dec30 table within ±$0.50 using table-implied F."""
    mid = combo_mid(F30, K, R30, T30)
    assert mid == pytest.approx(mid_pdf, abs=0.50)

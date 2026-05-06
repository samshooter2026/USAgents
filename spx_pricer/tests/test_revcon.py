"""Integration test for revcon module — uses PDF §6 worked example."""

import pytest
from unittest.mock import MagicMock
from datetime import date, datetime

from pricer.math import revcon_pnl


def test_revcon_pdf_example_exact():
    """PDF §6 example: K=7200, combo=20, r=3.63%, T=0.65, ES=7160+40, REV, 100 lots."""
    result = revcon_pnl(
        K=7200.0, combo=20.0, r=0.0363, T=0.65,
        es_future=7160.0, basis=40.0, side="REV",
        lots=100, spot=7109.0,
    )
    assert result["f_combo"] == pytest.approx(7220.48, abs=0.02)
    assert result["f_hedge"] == pytest.approx(7200.00, abs=0.01)
    assert result["edge_pts"] == pytest.approx(-20.48, abs=0.02)
    # PDF rounds intermediate 20.4775→20.48 before multiplying; we carry full precision.
    # Full-precision result is −$204,775. Tolerance ±$30 covers the rounding discrepancy.
    assert result["pnl_dollars"] == pytest.approx(-204800, abs=30)
    assert result["verdict"] == "PASS"

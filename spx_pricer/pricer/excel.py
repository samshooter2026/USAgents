"""
xlwings I/O layer — reads live Bloomberg values from an open Excel workbook.

Never use openpyxl or pandas.read_excel here. Bloomberg formulas re-evaluate
only inside the live Excel session; xlwings is the only bridge that works.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

import yaml

logger = logging.getLogger(__name__)


class StaleSnapshotError(Exception):
    """Raised when a required cell is #N/A, #NAME?, None, or otherwise unusable."""


@dataclass(frozen=True)
class MarketSnapshot:
    timestamp: datetime
    spot: float
    sofr: dict[float, float]           # tenor_years -> rate (decimal)
    div_yield_curve: dict[str, float]  # expiry_code -> annual div yield (decimal)
    borrow: dict[str, float]           # expiry_code -> borrow spread (decimal)
    es_future: float
    es_basis: float
    expiries: dict[str, date]          # expiry_code -> settlement date


# ---------------------------------------------------------------------------
# Low-level helpers
# ---------------------------------------------------------------------------

def _load_config(path: str | Path) -> dict:
    with open(path) as f:
        return yaml.safe_load(f)


def _get_workbook(path: str):
    """Attach to an already-open Excel session. Never opens a new file."""
    import xlwings as xw
    try:
        wb = xw.Book(path)
        return wb, wb.app
    except Exception as exc:
        raise ConnectionError(
            f"Cannot connect to '{path}'.\n"
            "Open the workbook in Excel with Bloomberg showing 'Connected' and try again.\n"
            f"Detail: {exc}"
        ) from exc


def _check_value(val: Any, ref: str) -> None:
    if val is None:
        raise StaleSnapshotError(
            f"{ref} is empty (None) — Bloomberg may not have refreshed yet."
        )
    if isinstance(val, str) and val.startswith("#"):
        raise StaleSnapshotError(
            f"{ref} = '{val}' — Bloomberg formula error. "
            f"Check the formula in that cell."
        )


def _read_cell(sheet, cell: str) -> Any:
    val = sheet.range(cell).value
    _check_value(val, f"{sheet.name}!{cell}")
    return val


def _read_range_flat(sheet, rng: str) -> list[Any]:
    """Read a range and return a flat list (handles single-column 2D lists)."""
    raw = sheet.range(rng).value
    if raw is None:
        return []
    if not isinstance(raw, (list, tuple)):
        return [raw]
    flat: list[Any] = []
    for item in raw:
        if isinstance(item, (list, tuple)):
            flat.extend(item)
        else:
            flat.append(item)
    return flat


def _to_decimal(val: float, ref: str) -> float:
    """
    Detect whether Excel stores the rate as percent (e.g. 3.60) or decimal (0.036).
    Rule: if abs(val) > 1.0, treat as percent and divide by 100.
    """
    if abs(val) > 1.0:
        logger.debug("Rate at %s = %.4f looks like percent; dividing by 100.", ref, val)
        return val / 100.0
    return val


def _bps_to_decimal(val: float, ref: str) -> float:
    """
    Convert a borrow value stored in basis-points in Excel (e.g. 12 → 0.0012).
    Borrow is always entered as bps in the BBG Data sheet yellow cells.
    """
    return val / 10000.0


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def pick_sofr(sofr: dict[float, float], T: float) -> tuple[float, float]:
    """
    Select SOFR rate for tenor T.
    Default rule: nearest tenor >= T; fall back to closest if none available above.
    Returns (matched_tenor, rate_decimal).
    """
    above = {t: r for t, r in sofr.items() if t >= T}
    if above:
        tenor = min(above, key=lambda t: t - T)
    else:
        tenor = min(sofr, key=lambda t: abs(t - T))
    rate = sofr[tenor]
    logger.info(
        "SOFR selected: tenor=%.4fy (T=%.4fy) rate=%.4f (%.2f%%)",
        tenor, T, rate, rate * 100,
    )
    return tenor, rate


def read_snapshot(config_path: str | Path) -> MarketSnapshot:
    """
    Read a live MarketSnapshot from the open Bloomberg-connected workbook.
    Forces a full recalculation before reading any cell.
    Raises StaleSnapshotError on any #N/A or empty cell.
    Raises ConnectionError if Excel/workbook is not reachable.
    """
    cfg = _load_config(config_path)
    wb, app = _get_workbook(cfg["workbook"])

    # Force Bloomberg formulas to flush
    try:
        app.calculate()
        if hasattr(app.api, "CalculateFull"):
            app.api.CalculateFull()
    except Exception as exc:
        logger.warning("CalculateFull skipped: %s", exc)

    sheets = {sh.name: sh for sh in wb.sheets}

    def sh(name: str):
        if name not in sheets:
            avail = list(sheets)
            raise KeyError(f"Sheet '{name}' not found. Available sheets: {avail}")
        return sheets[name]

    # Spot
    sc = cfg["spot"]
    spot = float(_read_cell(sh(sc["sheet"]), sc["cell"]))

    # SOFR curve
    sofc = cfg["sofr_curve"]
    sofr_raw = _read_range_flat(sh(sofc["sheet"]), sofc["range"])
    tenors: list[float] = sofc["tenors_years"]
    sofr: dict[float, float] = {}
    for tenor, val in zip(tenors, sofr_raw):
        ref = f"{sofc['sheet']}!{sofc['range']}[{tenor}y]"
        _check_value(val, ref)
        sofr[float(tenor)] = _to_decimal(float(val), ref)

    # Dividend strip → implied yield
    dc = cfg["div_strip"]
    div_raw = _read_range_flat(sh(dc["sheet"]), dc["range"])
    div_labels: list[str] = dc.get(
        "expiry_labels", [str(y) for y in dc.get("years", [])]
    )
    div_yield_curve: dict[str, float] = {}
    for label, val in zip(div_labels, div_raw):
        ref = f"{dc['sheet']}!{dc['range']}[{label}]"
        _check_value(val, ref)
        div_yield_curve[str(label)] = _to_decimal(float(val), ref)

    # Borrow curve — stored as bps in Excel (e.g. 12 = 12bps), convert to decimal
    bc = cfg["borrow_curve"]
    borrow_raw = _read_range_flat(sh(bc["sheet"]), bc["range"])
    borrow_labels: list[str] = bc["expiry_labels"]
    borrow: dict[str, float] = {}
    for label, val in zip(borrow_labels, borrow_raw):
        ref = f"{bc['sheet']}!{bc['range']}[{label}]"
        _check_value(val, ref)
        borrow[str(label)] = _bps_to_decimal(float(val), ref)

    # ES future
    esc = cfg["es_future"]
    es_future = float(_read_cell(sh(esc["sheet"]), esc["cell"]))

    # ES basis
    ebc = cfg["es_basis"]
    es_basis = float(_read_cell(sh(ebc["sheet"]), ebc["cell"]))

    # Expiries from config (static — settlement dates don't change)
    expiries: dict[str, date] = {}
    for entry in cfg["expiries"]:
        expiries[str(entry["code"])] = date.fromisoformat(str(entry["date"]))

    snap = MarketSnapshot(
        timestamp=datetime.now(),
        spot=spot,
        sofr=sofr,
        div_yield_curve=div_yield_curve,
        borrow=borrow,
        es_future=es_future,
        es_basis=es_basis,
        expiries=expiries,
    )
    logger.info(
        "Snapshot OK: spot=%.2f es=%.2f basis=%.2f sofr(1Y)=%.4f",
        snap.spot, snap.es_future, snap.es_basis, snap.sofr.get(1.0, 0),
    )
    return snap

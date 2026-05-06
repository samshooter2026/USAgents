# SPX Combo & Revcon Pricer

Live two-way pricer for SPX combos and revcons. Reads from Bloomberg via Excel.

## Prerequisites

- Python 3.11+
- Excel open with `Combos_RevCon_Pricer.xlsx`
- Bloomberg Excel add-in showing **Connected** in the Excel ribbon

## One-time setup

```
pip install -r requirements.txt
```

**Windows only:** Excel and Python must be the same bitness (both 64-bit is standard).

## Phase 0 — Map your Excel cells (run once)

With Excel open and Bloomberg connected:

```
python discover.py
```

Follow the prompts. It will scan every sheet for keywords and ask you to confirm
each cell. At the end it writes `config/excel_map.yaml`.

Then verify the live snapshot matches your Excel:

```
python -m pricer.snapshot
```

## Run the Streamlit UI

```
streamlit run app.py
```

Opens in your browser. Three tabs: Ladder, Revcon, Reverse Borrow.

## CLI commands

```
python -m pricer.cli ladder  --expiry DEC26 --width-bps 10
python -m pricer.cli revcon  --expiry DEC26 --strike 7200 --combo 20 \
                             --es 7160 --basis 40 --side REV --lots 100
python -m pricer.cli reverse --expiry DEC26 --strike 7200 --combo-mid 22.5
```

## Run tests

```
pytest tests/ -v
```

## Troubleshooting

**"Cannot connect to workbook"**
→ Open the .xlsx in Excel. Bloomberg must show Connected.

**"#N/A" or stale data**
→ Force a Bloomberg recalc: in Excel, press Ctrl+Alt+F9.
  If still stale, check your Bloomberg session.

**Bloomberg formulas not refreshing**
→ The pricer calls `CalculateFull()` on every read. If that fails, manually
  press `Ctrl+Alt+F9` in Excel then re-run.

## Config files

- `config/excel_map.yaml` — cell mappings (written by discover.py)
- `config/defaults.yaml`  — widths, thresholds, flags (edit freely)

## Quote log

Every quote shown in the Streamlit UI is appended to `logs/quotes.parquet`.
Use this for end-of-day calibration review.

# Current Working State

- Branch: `main` (synced to `origin/main` after publish).
- Workspace: `annuity_model`
- **Altair profit waterfall** (tip of `main`): signed walking bridge (green up / red down / blue totals from zero); SPIA left-to-right undiscounted level benefit → single premium; Term undiscounted claims → net PV. Replaces stacked `st.bar_chart` faux waterfall.

# Key Logic Implemented

## Profit decomposition (`pricing_ui.py`)

- `_build_profit_waterfall_chart_df`: cumulative `start`/`end` per row; change rows keep **signed** `delta` matching the table.
- `_altair_profit_waterfall_chart`: floating `mark_bar` with `y` / `y2`, zero reference rule, tooltips (Step, delta, from, to).
- Table unchanged: all components including single premium / net PV.

## Streamlit Pricing Run (unchanged from prior)

- `pricing_run_form_state`: `PRICING_RUN_NUMBER_INPUT_KEYS`, `_pricing_run_numeric_seeds`, `run_number_input` session-state hygiene.

## Tests

- `tests/test_pricing_ui_profit_decomposition.py` — waterfall walk geometry, SPIA/Term row reconciliation, labels.

# Validation Completed

- `python -m pytest tests/parity/ -v` — 19 passed.
- `python -m pytest tests/ -v` — 114 passed.

# Handoff Notes

- **Parity**: No changes to `pricing_projection`, ALM, or Excel generators; parity contract untouched.
- **New UI dependency**: Uses existing `altair`; chart renders via `st.altair_chart`.
- **If regressions**: Compare tooltips “Step ($)” to table “Amount ($)” for each component; first/last blue pillars should align with anchor and modeled premium on the y-axis.

# Immediate Next Steps

- None required unless product adds new decomposition rows — extend `_build_profit_decomposition_rows` and ensure new steps are ordered correctly for the waterfall.

# Unresolved Bugs / Pending Calculations

- None known from this change.

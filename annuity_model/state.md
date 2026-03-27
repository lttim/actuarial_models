# Current Working State

- Branch: `main` (ahead of `origin/main` with local commits from this chat).
- Workspace: `annuity_model`
- Latest commits from this chat:
  - `8bd1c52` — moved results charts into `Pricing Run`, removed separate `Charts` page, updated `state.md`.
  - `7493216` — fixed `run_m_mode` Streamlit widget state conflict warning.

# Key Logic Implemented

## Python Architecture Decisions

- Consolidated run-result visualization directly into `_render_run_and_results()` via `_render_pricing_run_charts()`.
- Removed chart-specific navigation architecture:
  - deleted `"charts"` from section labels/order
  - removed `_render_charts_section()`
  - removed `main()` route branch for `page == "charts"`
- Kept profit decomposition renderer as shared logic, now invoked from the run page flow.
- Removed mortality radio default-index pattern that conflicted with Session State:
  - `st.radio(..., key="run_m_mode")` now relies on normalized session value only
  - explicit `index=...` computation was removed to eliminate Streamlit warning.
- Updated run chart for benefits:
  - replaced point-in-time `PV benefits` chart with **Cumulative PV benefits** (`np.cumsum(expected_benefit_cashflows * discount_factors)`).
- Refactored profit decomposition into product-aware logic via `_build_profit_decomposition_rows(...)`:
  - `TERM_LIFE`: decomposition now uses claims/premium economics only (undiscounted expected claims, discounting effect, policyholder premium PV funding, net PV).
  - `SPIA`: retained mortality/discounting decomposition but renamed indexation line to a broader **Benefit design effect (e.g., indexation)**.
  - fallback path added for future products (Whole Life / Variable Annuity scaffolds) to avoid misleading SPIA-specific labels.

## Data Structures and UI Mapping

- `st.session_state` remains the source of truth for run inputs and selected modes.
- Charts still require successful run outputs (`pricing_res`, `pricing_contract`) before rendering.
- Existing use of `pandas.DataFrame` for chart feed tables and run output display remains unchanged.
- Expense assumptions continue to come from `st.session_state["pricing_excel_context"]["expenses"]` when available.

## Actuarial Modeling Assumptions Finalized

- No actuarial engine formula changes were made.
- No changes to disinvestment/reinvestment ordering, epsilon tie-break policy, or parity tolerances.
- Prior decision captured for Term Life UX: default monthly premium should be backsolved at issue under current assumptions.

# Immediate Next Steps

- Manual UI sanity check:
  1. confirm no `run_m_mode` default-vs-session warning appears
  2. confirm charts appear under `Pricing Run` only after a successful run
  3. confirm `Charts` sidebar section is absent
  4. confirm Term waterfall does not show SPIA-only language (e.g., indexation option cost).
- Run full test gates before release/push:
  - `pytest tests/parity/ -v`
  - `pytest tests/ -v`

# Unresolved Bugs / Pending Calculations

- No confirmed failing tests in this chat (tests not re-run after the final widget fix commit).
- Added regression tests in `tests/test_pricing_ui_profit_decomposition.py` for:
  - Term decomposition label hygiene (no indexation language)
  - Term decomposition reconciliation to net PV
  - SPIA decomposition labeling using product design effect wording.
- No known parity discrepancies introduced by the UI/session-state changes, but parity gates should still be rerun per repo rules.

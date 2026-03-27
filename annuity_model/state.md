# Current Working State

- Branch: `main`
- Workspace: `annuity_model`
- Recent changes in this chat were applied in `pricing_ui.py` (with existing unrelated `state.md` edits already present before this handoff update).
- Latest verified gates for this handoff:
  - `python -m pytest tests/parity/ -v` (pass)
  - `python -m pytest tests/ -v` (pass)

# Key Logic Implemented

## Python Architecture Decisions

- Relocated run-result charts into `Pricing Run` by adding `_render_pricing_run_charts()`:
  - `Cumulative PV benefits` line chart (replacing point-in-time PV benefit chart)
  - `Economic reserve` line chart
  - `Profit decomposition` waterfall/table (reused existing renderer)
- Hooked chart rendering directly under the existing month-by-month projection table in `_render_run_and_results()`.
- Removed dedicated chart navigation and rendering:
  - Deleted `"charts"` from `SECTION_LABELS` and `SECTION_ORDER`
  - Removed `_render_charts_section()`
  - Removed `main()` route branch for `page == "charts"`
- Deleted old generic chart block (`_render_charts`) that included additional visuals no longer requested.

## Data Structures and UI Mapping

- Kept `st.session_state` behavior unchanged.
- Chart rendering remains gated by existing successful run state:
  - charts only render when both `pricing_res` and `pricing_contract` exist
  - therefore charts appear only after clicking `Run pricing` and obtaining results
- Reused expense assumptions from `st.session_state["pricing_excel_context"]["expenses"]` for profit decomposition.

## Product/UI Behavior Decisions (Current)

- Mortality input:
  - Remove the Streamlit warning text: `The widget with key "run_m_mode" was created with a default value but also had its value set via the Session State API.`
- Term Life monthly premium default:
  - Default should be **backsolved premium at issue** (using current assumptions and selected benefit structure), because this keeps first-run behavior internally consistent with pricing logic and avoids arbitrary fixed default values.
  - UX expectation: when product changes to `Term Life`, auto-populate monthly premium with this backsolved value unless user has manually overridden premium in the current session.
- Charts:
  - Replace point-in-time `PV benefits` display with `Cumulative PV benefits` to better align with total discounted outgo interpretation over projection horizon.

## Actuarial Modeling Assumptions Finalized

- No actuarial formula changes were introduced in this chat.
- No changes to disinvestment/reinvestment logic, epsilon tie-break behavior, or parity tolerances.
- Changes are UI presentation/routing only.

# Immediate Next Steps

- Manual UI sanity check in Streamlit:
  1. Run pricing and confirm charts appear below the monthly projection table in `Pricing Run`.
  2. Confirm `Charts` section is no longer present in sidebar.
  3. Confirm no charts appear before first successful run.
- If desired, commit this change set in `pricing_ui.py` after review.

# Unresolved Bugs / Pending Calculations

- No failing tests at handoff.
- No known parity discrepancies introduced by this chat’s changes.
- Residual risk: this task did not include workbook (`openpyxl`) chart layout changes; scope was Streamlit UI chart placement/removal.

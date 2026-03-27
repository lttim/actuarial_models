# Handoff — SPIA Annuity Model (`annuity_model`)

## Tip of `main`

- **`f765189`** — Altair signed profit waterfall (green / red bridges, blue totals from zero).
- Recent chain: `741120d` Streamlit entry `pricing_ui` import; `57dc32a` stacked waterfall + run-form numeric seeds; `f765189` true waterfall.

## Repo / run

- **Git root:** `Code_Sandbox` (monorepo). **Product code:** `annuity_model/`.
- **Streamlit Cloud:** root `streamlit_app.py` prepends `annuity_model` to `sys.path` and runs `pricing_ui.main()`.
- **Local UI:** `cd annuity_model` → `streamlit run pricing_ui.py`.

## What’s in flight (UI)

| Area | Files | Notes |
|------|--------|--------|
| Profit decomposition | `pricing_ui.py` | `_build_profit_decomposition_rows`, `_build_profit_waterfall_chart_df`, `_altair_profit_waterfall_chart`; table matches signed `delta` in tooltips. |
| Pricing Run session keys | `pricing_run_form_state.py`, `pricing_ui.py` | `PRICING_RUN_NUMBER_INPUT_KEYS` + `_pricing_run_numeric_seeds` avoid Streamlit “default + Session State” warnings on `run_number_input`. |

## Parity / release gates

- **Do not change** calculation engines or Excel generators without contract updates + parity tests (`docs/model_parity_contract.md`).
- Before merge: `pytest tests/parity/ -v` (0.00 discrepancy), `pytest tests/ -v`.

## Validation (this workspace)

- `python -m pytest tests/parity/ -v` — **19** passed.
- `python -m pytest tests/ -v` — **114** passed.

## Suggested next steps (optional)

- Extend `_build_profit_decomposition_rows` + waterfall order when adding products or bridge lines.
- Apply `run_number_input` / seed pattern to ALM or What-if tabs if min-value first-paint issues appear.
- Document `run_alm_projection_from_pricing_result` in `AGENTS.md` if operators need it.

## Open issues

- None recorded from the waterfall / Streamlit workstream.

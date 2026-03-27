# Current Working State

- Branch: `main` (ahead of `origin/main`; push when ready).
- Workspace: `annuity_model`
- Latest completed commit: `685a78f` — Centralize ALM pricing dispatch; neutralize SPIA-only Excel copy.
- Prior related commits: `f266f24` (Term what-if / ALM UI routing + `tests/test_pricing_ui_what_if_term.py`), `22bac2c` (Term monthly premium default when missing).
- Working tree: clean after handoff commit (this file).

# Key Logic Implemented

## Term what-if and product-aware UI (`pricing_ui.py`)

- What-if branches by `pricing_product_type`: Term uses `compute_what_if_term_shocked_pricing` / `tp.price_term_life_level_monthly`; SPIA unchanged (index regime + MC + ALM overlay).
- Regression: Term no longer hits SPIA pricer with mismatched `index_levels_payment` shape (e.g. 240 vs 540 months).
- ALM “single MC path” repricing gated in `build_alm_pricing_for_mc_scenario` (SPIA + `SPIAContract` only).
- Diagnostics JSON: Term what-if export allowed without MC tail blocks when product is Term.
- Term what-if tolerates missing `expenses` in context (zero stub).

## Engine: unified ALM entrypoint (`pricing_projection.py`)

- `run_alm_projection_from_pricing_result(...)` dispatches SPIA vs Term (`lazy import term_projection` to avoid import cycle).
- Streamlit `_run_alm_from_session_pricing` delegates here.

## Excel / tooling copy

- `alm_excel_ladder.py`: field-guide strings say “liability” instead of “SPIA” where the grid is product-agnostic.
- `test_dashboard.py`: titles use “Model unit tests” (not SPIA-only).

## Tests

- `tests/test_pricing_ui_what_if_term.py` — Term what-if + ALM MC skip for Term.
- `tests/test_pricing_projection.py` — `test_run_alm_projection_from_pricing_result_dispatches_spia_and_term`.

## Parity / invariants

- No changes to disinvestment epsilon, tie-break, or Excel formula generators in this thread.
- Python-vs-Excel parity for ALM/Term workbooks unchanged by these edits.

# Validation Completed

- `python -m pytest tests/parity/ -v` — 18 passed.
- `python -m pytest tests/ -v` — 107 passed.

# Immediate Next Steps

- Push `main` (multiple commits ahead of `origin`) and open PR if desired.
- Quick manual checks: Term pricing run → What-if tab (no shape error); SPIA What-if identity at baseline dials; optional ALM run for Term vs SPIA.

# Unresolved Bugs / Pending Calculations

- No known failing tests.
- Optional: document `run_alm_projection_from_pricing_result` in `AGENTS.md` / parity checklist (not done here).

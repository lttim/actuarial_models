# Current Working State

- Branch: `main` (ahead of `origin/main`; push when ready).
- Workspace: `annuity_model`
- Latest completed commit: `41e3251` — Term Excel workbook: formula TermProjection + curve/qx sheets (+ product-specific recalc download filenames).
- Prior related commits: `3d89cae` (handoff state refresh), `685a78f` (ALM pricing dispatch; SPIA-neutral Excel copy), `f266f24` / `22bac2c` (Term what-if / monthly premium defaults).
- Working tree: clean; `model output/` (generated xlsx) remains untracked.

# Key Logic Implemented

## Term Excel recalculation workbook (`build_term_excel_workbook.py`)

- **TermProjection** is formula-driven from **Inputs** (issue age, DB, premium, term, horizon, spread), **ZeroCurve** (log-linear $\ln(\mathrm{DF})$ interpolation + endpoint extrapolation matching `YieldCurve.discount_factors`), and **QxTable** (`MortalityTableQx`) or **MortalMonthly** (RP2014+MP: per-month $q_x$ exported; survival chain still formula).
- **ModelCheck** aggregates columns **F / I / J** (claims, discount, PV net). Full Excel **CalculateFull** shows differences ~0 vs Python snapshot at export.
- Per-product download names in `product_registry.py` (e.g. `spia_recalc_model.xlsx`, `term_life_recalc_model.xlsx`).

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

- `tests/parity/test_term_parity.py` — Term workbook formula wiring + ModelCheck snapshot vs Python PV totals.
- `tests/test_product_registry.py` — product-specific `recalc_workbook_filename` expectations.
- `tests/test_pricing_ui_what_if_term.py` — Term what-if + ALM MC skip for Term.
- `tests/test_pricing_projection.py` — `test_run_alm_projection_from_pricing_result_dispatches_spia_and_term`.

## Parity / invariants

- SPIA ALM ladder / disinvestment parity unchanged by Term Excel export work.
- Term workbook is separate builder from SPIA (`product_excel.build_product_workbook`).

# Validation Completed

- `python -m pytest tests/parity/ -v` — 18 passed.
- `python -m pytest tests/ -v` — 107 passed.
- Optional: Excel COM **CalculateFull** on `model output/term_life_recalc_model.xlsx` — ModelCheck column D ~0 (needs `pywin32`).

# Immediate Next Steps

- Push `main` when ready.
- If extending Term horizon beyond exported **MortalMonthly** rows for RP/MP runs, re-download workbook from Streamlit.

# Unresolved Bugs / Pending Calculations

- No known failing tests.
- Optional: document `run_alm_projection_from_pricing_result` in `AGENTS.md` / parity checklist.

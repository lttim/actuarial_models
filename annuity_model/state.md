# Current Working State

- Branch: `main` (ahead of `origin/main`; push when ready).
- Workspace: `annuity_model`
- Latest completed commit (code): `acfa999` — Excel recalc parity: SPIA strips Dashboard/Runbook; Term aligns with SPIA naming (`Liabilities`, `YieldCurve`, `MonthlyCurve`) + optional ALM in download; shared `recalc_excel_shared.py`. Tip of `main` includes a follow-up commit refreshing this `state.md`.
- Prior related commits: `41e3251` (Term TermProjection + filenames), `3d89cae` (handoff), `685a78f` (ALM dispatch / SPIA-neutral copy).
- Working tree: clean after commit; `model output/` (generated xlsx) remains untracked.

# Key Logic Implemented

## Shared recalc conventions (`recalc_excel_shared.py`)

- Canonical sheet names and helpers for **YieldCurve** (nodes table) and **MonthlyCurve** (log-linear DF components matching `YieldCurve.discount_factors`).
- SPIA `build_pricing_excel_workbook` delegates monthly-curve formulas here; intended extension point for future products.

## SPIA Excel recalculation workbook (`build_pricing_excel_workbook.py`)

- **No** `Dashboard` or `Python_Runbook` tabs.
- ALM projection writer takes `n_months` + `y_last_row` (reusable from Term).
- **ModelCheck**: optional `pricing_rows`, `sheet_title`, `subtitle` for product-specific metrics.

## Term Excel recalculation workbook (`build_term_excel_workbook.py`)

- **Liabilities** sheet (replaces `TermProjection`): discount from **MonthlyCurve**; summary **X4–X9** aligned with SPIA-style checks; **C / O / S** columns satisfy ALM ladder (`ExpTotalCF`, discount).
- **YieldCurve** + **MonthlyCurve** (not inline DF); **Inputs** uses **B6** payments/year and **B9** spread (matches `ALM_Engine`).
- Optional **ALM_Engine** / **ALM_Projection** / field guide when snapshot + assumptions passed (via `product_excel.build_product_workbook`).
- **ModelCheck**: five Term rows (claims, premiums, net, APV, annuity-style factor) + optional ALM block.

## Product router (`product_excel.py`)

- Term path forwards `alm_snapshot` and `alm_assumptions` like SPIA.

## Streamlit (`pricing_ui.py`)

- Download help text no longer mentions Dashboard ALM charts.

## Tests

- `tests/parity/test_term_parity.py` — Liabilities wiring, ModelCheck vs Python, ALM sheets when snapshot passed.
- `tests/test_excel_workbook.py` — asserts no Dashboard/Python_Runbook; ALM + ModelCheck smoke (renamed test).

## Parity / invariants

- SPIA ALM ladder / disinvestment parity unchanged (same formulas; tabs removed only).
- Term workbook still separate builder; ALM uses same ladder code paths when embedded.

# Validation Completed

- `python -m pytest tests/parity/ -v` — 19 passed.
- `python -m pytest tests/ -v` — 108 passed.

# Immediate Next Steps

- Push `main` when ready.
- Re-download Term/SPIA recalc xlsx from Streamlit after UI runs if you rely on embedded ALM caches.

# Unresolved Bugs / Pending Calculations

- No known failing tests.
- Optional: document `run_alm_projection_from_pricing_result` in `AGENTS.md` / parity checklist.

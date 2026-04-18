# RILA parity contract (Python ↔ Excel)

**Version:** 1.0  
**Applies to:** `rila_projection.py` and `build_rila_excel_workbook.py` / liability sheet formulas.

This addendum sits beside `model_parity_contract.md` (SPIA ALM ladder). RILA parity covers **liability pricing and cashflows** on the liability sheet plus `ModelCheck` aggregates. When a RILA workbook includes `ALM_Projection`, SPIA ALM parity rules apply to those sheets unchanged.

## Month and index indexing

- Policy months are **1-based** in the liability grid: row `month = 1..n_months`.
- Scenario index: `L[0]` is `s0` from CSV month 0; for `j >= 1`, `L[j]` is the level for CSV month `j` (end of policy month `j`), matching `load_index_scenario_monthly_csv` / SPIA `levels_payment[j-1]`.

## Crediting month

- For `month` a multiple of 12 and `month >= 12`, segment raw return uses `L[month] / L[month - 12] - 1` (Excel and Python must use the same pair).

## Tolerances (liability sheet)

The table below is generated from `annuity_model/parity_constants.py` by
`scripts/render_parity_contract.py`. Edit the constants module, then run
`python -m annuity_model.scripts.render_parity_contract` to refresh; CI verifies
the docs are in sync via `--check`.

<!-- BEGIN GENERATED tolerances -->
| Variable | Tolerance | Units | Notes |
|----------|-----------|-------|-------|
| PV(benefit) cell | `1e-04` (`parity_constants.RILA_PV_TOL`) | USD | ModelCheck B5 |
| PV(expenses) cell | `1e-04` (`parity_constants.RILA_PV_TOL`) | USD | ModelCheck B6 |
| PV(total) cell | `1e-04` (`parity_constants.RILA_PV_TOL`) | USD | ModelCheck B7 |
| Single premium cell | `1e-04` (`parity_constants.RILA_PV_TOL`) | USD | ModelCheck B8 |
| Account value path | `1e-06` (`parity_constants.RILA_AV_TOL`) | USD | Per month |
| ModelCheck snapshot | `0.0 (exact)` (`parity_constants.MODELCHECK_TOL`) | USD | Exact match required |
<!-- END GENERATED tolerances -->

## Fee

- Monthly multiplicative fee: `AV *= (1 - fee_annual/12)` with `fee_annual` from Inputs.

## ModelCheck

- Surplus difference on embedded ALM (if present) follows SPIA checklist: **0.00** on golden scenarios after full recalc.

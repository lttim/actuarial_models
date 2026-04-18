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

| Variable | Absolute tolerance |
|----------|---------------------|
| Expected claim cashflow (month) | 1e-4 |
| Account value (end of month, per $ premium basis pre-scale) | 1e-10 |
| Discount factor | 1e-10 |
| PV claims / PV expenses / actuarial PV | 1e-4 |
| `ModelCheck` pricing differences | 0.00 (exact match to snapshot at export) |

## Fee

- Monthly multiplicative fee: `AV *= (1 - fee_annual/12)` with `fee_annual` from Inputs.

## ModelCheck

- Surplus difference on embedded ALM (if present) follows SPIA checklist: **0.00** on golden scenarios after full recalc.

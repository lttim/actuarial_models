# Portfolio parity addendum

This document augments [`model_parity_contract.md`](model_parity_contract.md)
for the **portfolio** surface.

## Python invariants

1. **Union grid** -- every `LiabilityPath` aggregated in a portfolio run must
   use the standard monthly `times_years = (1..N)/12` grid. Horizons differ only
   by `N`; shorter paths are zero-padded on the right.
2. **Rollup sum** -- `sum_t rollups_by_product_type[t].cf == portfolio_total.cf`
   elementwise within `PORTFOLIO_ROLLUP_TOL` (`parity_constants.py`).
3. **Single-policy degeneracy** -- one policy reproduces that policy’s path as
   the total and as its type rollup.

## Excel (portfolio workbook)

- **Static validation** -- `validate_workbook_or_raise` before save (same rule
  as per-product builders).
- **PolicyCashflows** -- one column per policy (monthly liability CF literals).
- **LiabilityAggregate.total_cf** -- Excel ``SUM`` over the policy columns on
  ``PolicyCashflows`` for that row (formula-linked rollup).
- **ModelCheck** -- (1) ``SUM(LiabilityAggregate!type_cols) -
  LiabilityAggregate!total_cf`` must be zero after recalc; (2) per-month
  ``excel_total_cf - python_total_cf`` (columns D/E) must be zero after recalc
  when the workbook is consistent. Parity tests assert formula wiring and
  Python snapshot agreement per `AGENTS.md`.

## Goldens

- `tests/parity/golden/portfolio/portfolio_5policy.json` -- refresh only with
  `UPDATE_GOLDEN_PORTFOLIO=1 pytest …/test_portfolio_golden.py`.
- SME lite baseline includes a `"portfolio"` block; refresh with
  `UPDATE_GOLDEN_SME=1` when deliberate scenario changes warrant it.

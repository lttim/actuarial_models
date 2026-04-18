# Glossary

One-paragraph definitions for every actuarial / engineering term used in the
codebase. New terms added to the engine MUST land here in the same PR.

## Products

**SPIA -- Single Premium Immediate Annuity.** A contract where the policyholder
pays a single premium up front and receives periodic income for life starting
immediately. Modelled in `pricing_projection.py`. Cashflows: monthly benefits
weighted by survival probability `_{t}p_x`, expense scale `(1 + infl)^t`.

**Term Life.** Level-premium pure-mortality contract paying a death benefit for
a fixed term. Modelled in `term_projection.py`. Liability cashflow is
expected-claims minus expected-premiums per month.

**RILA -- Registered Index-Linked Annuity.** Single-premium, deferred annuity
whose account value tracks an equity index with a participation rate, cap, and
buffer/floor on segment crediting. Modelled in `rila_projection.py`.

## Pricing & cashflow primitives

**q_x.** One-year mortality probability for a person aged exactly `x` -- the
fundamental input to any longevity-linked liability.

**_{t}p_x.** Probability that a person aged `x` survives to age `x + t`.
Computed as `prod_{k=0..t-1} (1 - q_{x+k})`.

**Yield curve.** A `YieldCurve` (see `pricing_projection.py`) is a
log-linear-on-zero-rates interpolator producing continuously-compounded
discount factors `DF(t) = exp(-z(t) * t)`. Spread `s` shifts the zero rate
additively before discounting.

**Discount factor (DF).** Present-value scaling factor; `DF(0) = 1` and
`DF(t) > 0` strictly. Tolerance: `parity_constants.TOL_DF`.

## ALM ladder

**ALM -- Asset-Liability Management.** The end-of-month projection that holds
assets (cash + Treasury ladder) against the liability cashflow path. Source of
truth for all asset-side numbers: `pricing_projection.run_alm_projection`.

**Bucket.** One bond-tenor sleeve in the asset ladder (e.g. `1Y`, `5Y`, `20Y`)
plus an explicit `Cash` bucket. Each has a target weight, current face, and
remaining tenor.

**t_rem.** Remaining time-to-maturity (years) of a bond bucket. Decremented by
`dt = 1/12` each month. **Critical invariant:** `t_rem` is never compared raw
across buckets; an epsilon-adjusted argsort key is always used (see
"Tie-break").

**Face / MV / dmv.** Face = par. MV = market value = `face * DF(t_rem)`. dmv =
market-value change for a reinvest / disinvest action; converted to face via
`face = dmv / DF`.

**Disinvest.** When end-of-month cash is short of the liability, the ladder
sells bonds shortest-first until cash >= 0 (or the residual is borrowed,
governed by `borrowing_policy`).

**Reinvest (pro_rata).** When excess cash exists AND at least one bond bucket
matures in the current month, surplus cash is reinvested across underweight
buckets proportional to their gap-to-target. See
`tests/parity/excel_formula_sim.py :: excel_reinvest_pro_rata`.

**Tie-break.** When two bucket `t_rem` values are within `5e-10` of each other,
the lower-indexed bucket sells first. The `5e-10` threshold is half the
inter-bucket epsilon (`EXCEL_DISINVEST_EPSILON = 1e-9`); both values live in
`parity_constants.py`. See `docs/model_parity_contract.md` section 2.

**P2P -- Point-to-Point (RILA).** A segment whose return is computed as
`L[end] / L[start] - 1` over the segment window (typically 12 months for an
annual P2P).

**Cap / Floor / Participation (RILA).** Crediting transformation:
`credited = clip(participation * raw, floor, cap)`.

**Account value (AV).** RILA's balance: starts at single premium, scales by
`(1 + credited)` at each segment boundary, and pays a monthly multiplicative
fee `AV *= (1 - fee_annual / 12)`.

## Excel pipeline

**ModelCheck sheet.** A workbook tab whose B-column cells (B5..B12) hold
Python-computed reference values that the Excel formulas must match exactly
(`MODELCHECK_TOL = 0.0`). The "exact" gate is what catches Excel-formula bugs.

**Liability sheet.** Per-product tab containing the month-by-month projected
cashflows. Column letters live in `liability_layouts.LIABILITY_LAYOUTS` --
**never** hardcode `"S"` or `"M"` outside that registry.

**Validator -- `excel_workbook_validator.py`.** Static (AST-free, parser-based)
check of every formula cell: function arity, range syntax, sheet existence,
known function name, and the structural rule that RILA's ALM_Projection sheet
references the liability column from `LIABILITY_LAYOUTS["rila"]`, not SPIA's.

**OOXML / data_only.** openpyxl's `data_only=True` reads the *cached* value
from the last save instead of re-running formulas. Always recalculate the
workbook (open in Excel or call a recalc engine) before reading values for
parity assertions; relying on data_only silently masks formula bugs.

## Process / governance

**Parity gate.** The `tests/parity/` suite. Must be 100% green to merge.

**Validator gate.** `validate_workbook_or_raise(wb)` invoked immediately
before every `wb.save(...)` -- enforced by convention today, by AST meta-test
in P4.

**Parity contract.** `docs/model_parity_contract.md` (SPIA + ALM) and
`docs/rila_parity_contract.md` (RILA). Both have a generated tolerance block
sourced from `parity_constants.py`.

**Parity trace.** A CSV produced by `scripts/parity_trace.py` containing
month-by-month python-vs-excel state for visual diff in Excel or pandas.
Cited in `runbooks/investigate_parity_break.md`.

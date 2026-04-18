# Runbook: investigate a parity break

## When to run this

* `tests/parity/` failed locally or in CI.
* `ModelCheck` cells differ from Python by any non-zero amount.
* A previously-green parity test starts failing after an unrelated change.

## Triage flow (5 minutes)

```
parity test failed
   │
   ├── Failure < TOL_DOLLAR  ──► Floating-point ordering issue. Almost always
   │                            disinvest tie-break. Go to step 3.
   │
   ├── Failure == single bucket / month ──► Index drift. Go to step 4.
   │
   ├── Failure cumulative across months ──► Carry-state bug (cash, t_rem,
   │                                        or borrowing balance). Go to step 5.
   │
   └── Failure on first month only ──► Initialisation mismatch. Compare
                                       initial-state code in pricing engine
                                       vs Excel "Inputs" sheet.
```

## Procedure

### 1. Pin the failure to a specific scenario

```bash
cd annuity_model
python -m pytest tests/parity -q --tb=short
```

Identify the failing test name and parameter set (e.g.
`test_alm_parity::test_disinvest_excess_cash_pro_rata[delta=0]`).

### 2. Run the parity trace for the same scenario

```bash
python scripts/parity_trace.py --steps 60 \
    --output traces/break_$(date +%Y%m%d_%H%M).csv
```

This writes a CSV with `month, py_*, xl_*, diff_*` columns. Open in Excel
or Pandas. The first month with `|diff_*| > parity_constants.TOL_DOLLAR`
is your "moment of divergence".

If the trace is silent ("All metrics within tolerance") then the failure
is in a code path the trace does not cover -- extend `_trace()` in
`scripts/parity_trace.py` for the affected metric, or write a focused
unit test against the failing assertion.

### 3. Disinvestment tie-break check

If divergence is < `TOL_DOLLAR` at the first month a bond bucket matures:

```python
from parity_constants import EXCEL_DISINVEST_EPSILON, EXCEL_DISINVEST_THRESHOLD
assert EXCEL_DISINVEST_THRESHOLD < EXCEL_DISINVEST_EPSILON / 2
```

Then in the failing scenario, instrument `pricing_projection.py` near the
disinvest argsort to print the per-bucket epsilon-adjusted key. Compare to
`excel_formula_sim.excel_disinvest_shortest_first` for the same inputs.
The fix is **never** to widen the tolerance; the fix is to align the
tie-break logic.

### 4. Index / column-letter drift

```python
from liability_layouts import liability_layout_for
print(liability_layout_for(failing_product_code))
```

If the validator is green but parity is red on the ALM_Projection sheet,
the builder is reading the right column but writing the wrong one (or
vice versa). Grep for the literal column letter: it MUST appear only in
`liability_layouts.py`.

### 5. Carry-state bug

Cumulative drift means the engine is silently dropping a delta each
month. Common culprits:

* Borrowing balance accrual rate differs between Python and Excel.
* `t_rem` decremented in Python but reset to nominal in Excel (or vice
  versa) for buckets that crossed maturity mid-month.
* Reinvestment fires in one engine but not the other due to a different
  `gap_sum` epsilon -- check `parity_constants.REINVEST_GAP_EPS` and
  `REINVEST_XSR_EPS`.

Add a focused regression test in `tests/parity/test_alm_parity.py` that
isolates the discovered bug *before* fixing it; that way the fix is
guarded forever.

## Tolerance change is not a fix

The parity contract tolerances (`parity_constants.TOL_DOLLAR`,
`MODELCHECK_TOL`, etc.) are immutable except via:

1. PR that updates `parity_constants.py`.
2. Re-rendered docs (`python scripts/render_parity_contract.py`).
3. CODEOWNERS approval.
4. Entry in `docs/model_change_log.md` with parity-trace before/after.

Anything else is an automatic merge block.

## Related

* [debug_validator_failure.md](debug_validator_failure.md)
* [docs/model_parity_contract.md](../model_parity_contract.md) -- the contract.
* [docs/parity_test_checklist.md](../parity_test_checklist.md) -- release
  checklist that calls this runbook.

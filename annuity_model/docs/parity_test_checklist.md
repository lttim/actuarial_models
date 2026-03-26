# Parity Test Checklist

Use this checklist before every merge to main and every product release.

---

## Before every merge (PR gate)

- [ ] `pytest tests/parity/ -v` — all pass, zero failures
- [ ] `pytest tests/ -v` — all pass, zero failures
- [ ] No new ad-hoc tolerances added to code without updating `docs/model_parity_contract.md`
- [ ] If any formula-generating file changed: confirm the regenerated workbook
      `ModelCheck` sheet shows **0.00** surplus discrepancy

---

## Before every product release

- [ ] All merge checklist items above
- [ ] Golden scenario suite passes with **|diff| ≤ 0.01** on surplus at every month
- [ ] Tie-break regression test passes (disinvestment sells from lowest-indexed bucket only)
- [ ] No `data_only=True` Excel reads used for validation (must use recalculated values)
- [ ] Step-level comparison (not just month 60 surplus) shows ≤ 1e-4 at every month

---

## After any change to epsilon / tolerance / ordering logic

- [ ] Boundary test added: prove no double-trigger at the new threshold
- [ ] `docs/model_parity_contract.md` epsilon table updated
- [ ] All golden scenarios re-run and results captured
- [ ] Old failing test (if fixing a bug) preserved as a regression test with `@pytest.mark.regression`

---

## Red flags — stop and investigate if you see these

- Final surplus matches but intermediate monthly values differ → offsetting errors
- Python sells from a higher-indexed bucket when a lower-indexed has equal tenor → tie-break failure
- Two bond faces both reduced in the same disinvestment step → double-sell bug
- Excel cached value (`data_only=True`) differs from full-recalculation value → stale cache
- `t_rem` argsort produces different order on repeated runs → raw float ordering (add epsilon)

---

## Regression test naming convention

```python
# Tests that fix a specific historical bug
@pytest.mark.regression
def test_disinvest_no_double_sell_equal_tenor_buckets():
    """Regression: 2026-03 — double-sell when two bonds had t_rem differing by ~1e-16."""
    ...
```

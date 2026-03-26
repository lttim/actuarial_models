# Parity Test Checklist

Use this checklist before every merge to main and every product release.

---

## Before every merge (PR gate)

- [ ] `pytest tests/parity/ -v` — all pass, zero failures
- [ ] `pytest tests/ -v` — all pass, zero failures
- [ ] No new ad-hoc tolerances added to code without updating `docs/model_parity_contract.md`
- [ ] If any formula-generating file changed: confirm regenerated workbook reconciliation
      sheet shows **0.00** difference

---

## Before every product release

- [ ] All merge checklist items above
- [ ] Golden scenario suite passes with `|diff| ≤ [tolerance]` on primary output at every step
- [ ] Tie-break regression test passes (selection applies to lowest-indexed item only)
- [ ] No `data_only=True` Excel reads used for validation
- [ ] Step-level comparison (not just final output) shows ≤ tolerance at every step

---

## After any change to epsilon / tolerance / ordering logic

- [ ] Boundary test added: prove no double-selection at the new threshold
- [ ] `docs/model_parity_contract.md` epsilon table updated
- [ ] All golden scenarios re-run and results captured
- [ ] Old failing test preserved as `@pytest.mark.regression`

---

## Red flags — stop and investigate if you see these

- Final output matches but intermediate values differ → offsetting errors
- Engine A selects higher-indexed item when lower-indexed has equal sort key → tie-break failure
- Two items both modified in same selection step when only one should be → double-selection bug
- Excel cached value differs from full-recalculation value → stale cache
- Sort order changes on repeated runs → raw float ordering (add epsilon)

---

## Regression test naming convention

```python
@pytest.mark.regression
def test_[description_of_bug]():
    """Regression: [DATE] — [one-line description of what was broken].
    [Optional: what input conditions trigger it.]
    [Optional: what the fix was.]
    """
    ...
```

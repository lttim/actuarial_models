# Agent Instructions — SPIA Annuity Model

This project implements a SPIA (Single Premium Immediate Annuity) pricing and ALM projection
engine with two synchronised calculation engines: Python and Excel.

## Non-negotiable parity requirement

Every code change must maintain zero-discrepancy between the Python ALM engine and the
generated Excel workbook. Read and follow `docs/model_parity_contract.md` before modifying
any calculation logic.

## Before completing any task

```
pytest tests/parity/ -v         # must all pass
pytest tests/ -v                # must all pass
```

If a task changes Excel-generating code, also verify the `ModelCheck` sheet in a regenerated
workbook shows 0.00 surplus difference.

## Key files

| File | Purpose |
|------|---------|
| `pricing_projection.py` | Python ALM engine (source of truth for calculation logic) |
| `alm_excel_ladder.py` | Excel formula generator — must match Python exactly |
| `build_pricing_excel_workbook.py` | Workbook builder and OOXML cache injector |
| `docs/model_parity_contract.md` | Parity contract: tolerances, tie-break, epsilon policies |
| `docs/parity_test_checklist.md` | Release gate checklist |
| `tests/parity/test_alm_parity.py` | Parity regression tests |

## Critical rules

1. Any change to disinvestment/reinvestment ordering logic requires a new parity test.
2. Never change epsilon values without updating the contract and adding a boundary test.
3. Never rely on raw floating-point comparison of `t_rem` values for ordering — always use epsilon tie-breaking.
4. Every bug fixed must have a permanent regression test capturing the exact scenario.
5. Step-level reconciliation (monthly state) not just final surplus.

## Parity kit for future products

See `../actuarial_parity_kit/` for a reusable governance template to carry forward to new
actuarial product repos. Copy that directory into any new repo and adapt the test fixtures.

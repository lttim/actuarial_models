# Agent Instructions — [PRODUCT NAME] Model

<!-- Replace [PRODUCT NAME] with your product (e.g. ULSG, Whole Life, DI). -->

This project implements a [PRODUCT TYPE] pricing and projection engine with two
synchronised calculation engines: Python and Excel.

## Non-negotiable parity requirement

Every code change must maintain zero-discrepancy between the Python engine and the generated
Excel workbook. Read and follow `docs/model_parity_contract.md` before modifying any
calculation logic.

## Before completing any task

```
pytest tests/parity/ -v         # must all pass
pytest tests/ -v                # must all pass
```

If a task changes Excel-generating code, also verify the reconciliation sheet in a
regenerated workbook shows 0.00 difference.

## Key files

| File | Purpose |
|------|---------|
| `[engine].py` | Python calculation engine (source of truth) |
| `[excel_generator].py` | Excel formula generator — must match Python exactly |
| `[workbook_builder].py` | Workbook builder |
| `docs/model_parity_contract.md` | Parity contract: tolerances, tie-break, epsilon policies |
| `docs/parity_test_checklist.md` | Release gate checklist |
| `tests/parity/test_parity.py` | Parity regression tests |

## Critical rules

1. Any change to ordering/selection logic requires a new parity test.
2. Never change epsilon values without updating the contract and adding a boundary test.
3. Never use raw floating-point comparison for ordering accumulated values — use epsilon.
4. Every bug fixed must produce a permanent `@pytest.mark.regression` test.
5. Step-level reconciliation (periodic state), not just final output.

## Reuse guidance

This repo was set up using the `actuarial_parity_kit` template. When starting a new
product, copy that kit rather than starting from scratch.

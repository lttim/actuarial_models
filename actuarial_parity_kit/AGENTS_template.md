# Agent Instructions — [PRODUCT NAME] Model

<!-- Replace [PRODUCT NAME] with your product (e.g. ULSG, Whole Life, DI). -->

This project implements a [PRODUCT TYPE] pricing and projection engine with two
synchronised calculation engines: Python and Excel.

## Non-negotiable parity requirement

Every code change must maintain zero-discrepancy between the Python engine and the generated
Excel workbook. Read and follow `docs/model_parity_contract.md` before modifying any
calculation logic.

## Before completing any task -- canonical gates

> This block is the *single source of truth* for the "what do I run before
> claiming a task is done?" question. Any other doc that needs to state these
> gates (root `AGENTS.md`, `actuarial-parity.mdc`, `parity_test_checklist.md`,
> `CONTRIBUTING.md`) MUST link here instead of restating.

```bash
# 1. Parity gate (blocks any merge on failure)
python -m pytest tests/parity -q

# 2. Full unit-test gate
python -m pytest -q

# 3. End-to-end smoke (every implemented product + full Excel validator)
python scripts/deep_smoke.py

# 4. Tolerance contract is in sync with parity_constants.py
python scripts/render_parity_contract.py --check
```

All four must exit 0.

If a task changes Excel-generating code, ALSO run the static workbook
validator and verify the `ModelCheck` links point to the expected formula
summary rows. Avoid headless desktop spreadsheet automation in automated
gates; it is not reliable across macOS and sandboxed agent environments.
Provide a short runbook for that flow at `docs/runbooks/regenerate_excel_cache.md`.

If parity fails, follow `docs/runbooks/investigate_parity_break.md`.
If the Excel validator fails, follow `docs/runbooks/debug_validator_failure.md`.
**Never widen a tolerance to make a test pass.** Tolerance changes route
through `parity_constants.py` plus `model_change_log.md`.

## Key files

| File | Purpose |
|------|---------|
| `[engine].py` | Python calculation engine (source of truth) |
| `[excel_generator].py` | Excel formula generator — must match Python exactly |
| `[workbook_builder].py` | Workbook builder + OOXML cache injector |
| `parity_constants.py` | Single source of truth for all numerical tolerances |
| `model_change_log.md` | Human-readable log of every tolerance / formula change |
| `docs/model_parity_contract.md` | Parity contract: tolerances, tie-break, epsilon policies |
| `docs/parity_test_checklist.md` | Release gate checklist |
| `tests/parity/test_parity.py` | Parity regression tests |

## Critical rules

1. Any change to ordering/selection logic requires a new parity test.
2. Never change epsilon values without updating the contract and adding a boundary test.
3. Never use raw floating-point comparison for ordering accumulated values — use epsilon.
4. Every bug fixed must produce a permanent `@pytest.mark.regression` test.
5. Step-level reconciliation (periodic state), not just final output.
6. **Excel formulas must pass static validation before saving.** Every workbook
   builder must call `excel_workbook_validator.validate_workbook_or_raise(wb)`
   immediately before `wb.save(...)`. The validator should check:
   - Balanced parentheses and string quotes (catches f-string concatenation bugs).
   - Correct argument counts for every known function (no `IF(cond, value)`
     without an explicit false branch).
   - No embedded Excel error literals (`#REF!`, `#NAME?`, `#DIV/0!`, ...).
   - No trailing empty arguments (`IF(a, b, )` / `IFERROR(x, )`).
   - Cross-sheet column resolution (every `Sheet!Col` reference, including those
     embedded in `INDIRECT(...)`, must resolve to a column that has data on
     the target sheet -- catches silent reconciliation bugs where a column
     letter from one product is reused in another).

## Reuse guidance

This template lives in `actuarial_parity_kit/`. When starting a new product, copy
that directory rather than starting from scratch, then customise the table
above and run all four canonical gates from a fresh checkout to confirm the
scaffold is wired correctly before adding any actuarial code.

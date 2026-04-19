# Agent Instructions — SPIA Annuity Model

This project implements SPIA, Term Life, and RILA (accumulation) pricing with ALM projection
and two synchronised calculation engines where applicable: Python and Excel.

> **AI coding agents:** read [`docs/AI_AGENT_PREFLIGHT.md`](docs/AI_AGENT_PREFLIGHT.md)
> *first*. It contains the source-of-truth map, decision tree, and hard
> rules an autonomous agent must follow on this repo. This file
> (`AGENTS.md`) remains the canonical owner of the four-gate list below;
> the pre-flight doc just routes agents to the right runbook before they
> get there.

## Non-negotiable parity requirement

Every code change must maintain zero-discrepancy between the Python ALM engine and the
generated Excel workbook. Read and follow `docs/model_parity_contract.md` before modifying
any calculation logic.

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

# 3. End-to-end smoke (3 products + full Excel validator)
python scripts/deep_smoke.py

# 4. Tolerance contract is in sync with parity_constants.py
python scripts/render_parity_contract.py --check
```

All four must exit 0.

Two always-on gates inside gate (2) above are worth calling out explicitly
because they catch the bug classes the parity engine cannot:

* **`tests/ui/test_apptest_full_workflow.py`** — runs Streamlit's
  `AppTest` harness end-to-end against `pricing_ui.py` AND
  `streamlit_app.py` (the Streamlit Cloud entry). Asserts every
  sidebar section renders without exceptions, that clicking
  "Run pricing" populates `st.session_state['pricing_res']` for
  every implemented product, and that the Excel download bytes pass
  strict `excel_workbook_validator`. If you touch `pricing_ui.py`,
  any product adapter, or the workbook builder dispatch, this gate
  is your first signal.
* **`tests/parity/test_excel_recalc_per_product.py`** — per-product
  Excel↔Python recalc gate. The "always-on" layer asserts the
  Python literals the builder bakes into ModelCheck column B equal
  the engine within `MODELCHECK_TOL` for every product (no skip);
  the LibreOffice layer (skipped without `soffice` on PATH; runs in
  CI parity-gate) actually recalculates the workbook through the
  reference spreadsheet engine. The pre-existing
  `tests/parity/test_runtime_excel_recalc.py` is the SPIA-only
  ancestor — the per-product file supersedes it.

If a task changes Excel-generating code, ALSO open the regenerated workbook
in Excel (or run `recalc_excel_shared.recalculate_workbook`) and verify the
`ModelCheck` sheet shows 0.00 difference -- the parity tests load values
from openpyxl and CANNOT see formula bugs that only surface on Excel
recalc. Step-by-step recipe in
[docs/runbooks/regenerate_excel_cache.md](docs/runbooks/regenerate_excel_cache.md).

If parity fails, follow
[docs/runbooks/investigate_parity_break.md](docs/runbooks/investigate_parity_break.md).
If the validator fails, follow
[docs/runbooks/debug_validator_failure.md](docs/runbooks/debug_validator_failure.md).
**Never widen a tolerance to make a test pass.** Tolerance changes route
through `parity_constants.py` plus `model_change_log.md`.

## Key files

| File | Purpose |
|------|---------|
| `pricing_projection.py` | Python ALM engine (source of truth for calculation logic) |
| `alm_excel_ladder.py` | Excel formula generator — must match Python exactly |
| `build_pricing_excel_workbook.py` | Workbook builder and OOXML cache injector |
| `docs/model_parity_contract.md` | Parity contract: tolerances, tie-break, epsilon policies |
| `docs/parity_test_checklist.md` | Release gate checklist |
| `tests/parity/test_alm_parity.py` | Parity regression tests |
| `rila_projection.py` | RILA liability / crediting (Python) |
| `build_rila_excel_workbook.py` | RILA Excel workbook + `ModelCheck` |
| `docs/rila_product_spec.md` | RILA v1 product definition |
| `docs/rila_parity_contract.md` | RILA Python ↔ Excel parity addendum |
| `tests/parity/test_rila_parity.py` | RILA parity tests |

## Critical rules

1. Any change to disinvestment/reinvestment ordering logic requires a new parity test.
2. Never change epsilon values without updating the contract and adding a boundary test.
3. Never rely on raw floating-point comparison of `t_rem` values for ordering — always use epsilon tie-breaking.
4. Every bug fixed must have a permanent regression test capturing the exact scenario.
5. Step-level reconciliation (monthly state) not just final surplus.
6. **Excel formulas must pass static validation before saving.** Every workbook builder must
   call `excel_workbook_validator.validate_workbook_or_raise(wb)` immediately before
   `wb.save(...)`. The validator now checks both *syntax* and *cross-sheet semantics*:
   - Balanced parentheses and string quotes (catches f-string concatenation bugs).
   - Correct argument counts for every known function (no `IF(cond, value)` without an
     explicit false branch — Excel will repair the file with "Removed Records: Formula
     from /xl/worksheets/sheetN.xml part" and replace cells with `#DIV/0!`).
   - No embedded Excel error literals (`#REF!`, `#NAME?`, `#DIV/0!`, ...) inside formula
     bodies.
   - No trailing empty arguments (`IF(a, b, )` / `IFERROR(x, )`) — those are almost
     always an f-string that lost its substitution. Write `,""` or `,0` explicitly.
   - **Cross-sheet column resolution.** Every `Sheet!Col` reference *and* every
     `Sheet!Col` literal embedded in `INDIRECT(...)` must resolve to a column that
     actually has data on the target sheet. This catches the class of silent
     reconciliation bug where a SPIA-style `Liabilities!S` reference is generated
     from a RILA workbook (RILA puts `ExpTotalCF` in column M, not S). Excel
     coerces the missing column to zero in `SUMPRODUCT`/`INDEX`, so without this
     check the failure only shows up as drift in `ModelCheck` after Excel recalcs.
   When introducing a new built-in function in any builder, register its arity in
   `excel_workbook_validator.FUNCTION_ARITIES`. When introducing a new product, the
   shared ALM helper `_write_alm_projection_sheet` accepts `liability_total_col`
   (default `"S"` for SPIA/Term) and `liability_discount_col` (default `"O"`) — pass
   the column letters that match your product's `Liabilities` layout (e.g. RILA
   uses `liability_total_col="M"`). The end-to-end gate
   `tests/test_excel_export_validation.py` builds workbooks for every implemented
   product (SPIA, Term, RILA) and runs the validator on them; that test must
   always pass.

## Parity kit for future products

See `../actuarial_parity_kit/` for a reusable governance template to carry forward to new
actuarial product repos. Copy that directory into any new repo and adapt the test fixtures.

<!--
P0 hardening 2026-04: this template makes the four canonical gates from
annuity_model/AGENTS.md visible on every PR. Reviewers should not approve
until all four boxes are checked. Run `just preflight` for the one-liner
that exercises all four locally.
-->

## Summary

<!-- 1-3 sentences. What does this PR change and why? Link the issue if any. -->

## Products affected

<!-- Tick every product whose calculation, UI, or Excel surface is touched.
     If you ticked anything other than "none", add the relevant tests/parity/
     and tests/test_excel_export_validation.py runs to your local checks. -->

- [ ] SPIA
- [ ] Term Life
- [ ] RILA
- [ ] None (docs / tooling / CI only)

## Surface affected

- [ ] Engine math (`pricing_projection.py`, `term_projection.py`, `rila_projection.py`)
- [ ] Excel builder (`build_*_excel_workbook.py`, `alm_excel_ladder.py`, `excel_workbook_validator.py`)
- [ ] Streamlit UI (`pricing_ui.py`, `streamlit_app.py`, `pricing_run_form_state.py`)
- [ ] Parity / test infrastructure
- [ ] CI / governance (`.github/`, `.pre-commit-config.yaml`, `Justfile`, `pyproject.toml`)
- [ ] Docs only

## Canonical gates

> Source of truth: [`annuity_model/AGENTS.md`](annuity_model/AGENTS.md#before-completing-any-task----canonical-gates).
> One-liner that runs all four: `just preflight`.

- [ ] **Parity gate**: `cd annuity_model && python -m pytest tests/parity -q` -- exit 0
- [ ] **Full unit-test suite**: `cd annuity_model && python -m pytest -q` -- exit 0
- [ ] **End-to-end smoke**: `cd annuity_model && python scripts/deep_smoke.py` -- exit 0
- [ ] **Tolerance docs in sync**: `cd annuity_model && python scripts/render_parity_contract.py --check` -- exit 0

## Excel safety (only if Excel-generating code changed)

- [ ] Regenerated workbook opened in Excel/LibreOffice and `ModelCheck` shows 0.00 difference
  (or the `tests/parity/test_runtime_excel_recalc.py` LibreOffice gate passes locally)
- [ ] No new Excel function used without registering its arity in
  [`excel_workbook_validator.FUNCTION_ARITIES`](annuity_model/excel_workbook_validator.py)
- [ ] No SPIA-style `Liabilities!S` reference emitted from a non-SPIA builder
  (cross-sheet validator catches this; double-check if you added a new builder)

## Tolerance / parity-constant changes

- [ ] `parity_constants.py` was NOT modified, **OR**
- [ ] `parity_constants.py` was modified AND
  [`annuity_model/docs/model_change_log.md`](annuity_model/docs/model_change_log.md)
  has a new entry with PR link, parity-trace before/after, and justification.
  (CI enforces this via `scripts/check_parity_constants_changelog.py`.)

## New product or scaffolding work

- [ ] New product is registered in `product_registry._PRODUCT_ADAPTERS` AND
  `product_excel.@register_builder` AND has a liability-path converter
  registered via `liability_dispatch.register_liability_path_converter`
  (the `tests/test_meta_invariants.py` and
  `tests/test_builder_registry_invariants.py` invariants enforce this)
- [ ] Per-product `_PRICING_METRIC_FORMATTERS` entry added (no SPIA fallback)
- [ ] Streamlit UI section wires every widget value to the typed engine
  contract (no hard-coded literals; see Term wiring regression test)

## Reviewer notes

<!-- Anything the reviewer should focus on, anything explicitly out of scope,
     known follow-ups, etc. -->

# Current Working State

- Branch: `main`
- Latest pushed commit: `6e785a5` on `origin/main`
- Status at handoff: hard rename/deprecation completed; canonical naming is now `pricing_*` (legacy `spia_*` files removed).
- Validation status: `python -m pytest tests/parity/ -v` passed (18/18), `python -m pytest tests/ -v` passed (95/95), and lints were clean on edited files.

# Key Logic Implemented

## Python Architecture Decisions

- Product orchestration is centralized in `product_registry.py` using:
  - `ProductType` enum
  - adapter protocol + concrete adapters (SPIA and TERM_LIFE)
  - product capability/config helpers (UI behavior, mortality options/defaults, metric schema, export naming)
- Product workbook dispatch is centralized in `product_excel.py`, routing per product type to the correct workbook spec/builder.
- Canonical modules and entrypoints are now:
  - `pricing_projection.py` (core pricing/ALM engine)
  - `build_pricing_excel_workbook.py` (pricing workbook builder)
  - `pricing_ui.py` and `run_pricing_ui.bat` (UI launch)
- Term product logic is isolated in:
  - `term_projection.py` (deterministic term pricing + liability path adapter)
  - `build_term_excel_workbook.py` (Term Excel path with ModelCheck reconciliation)

## UI/Data Structure Decisions

- Product-specific UI inputs/labels/defaults are centralized in registry helpers rather than hardcoded in UI.
- Pricing output metrics are product-driven via `get_pricing_metrics(...)` and rendered generically in the UI.
- Export filenames are product-driven via `get_product_ui_config(...)`.
- `pandas.DataFrame` remains the core table structure for:
  - projection table display/export
  - workbook snapshot/reconciliation helper tables
  - parity-facing tabular checkpoints

## Actuarial Modeling Assumptions Finalized

- Term variant implemented: **20-year level term life**.
- Premium structure: **level monthly premium**.
- Benefit timing: **end of policy year of death (EOY)**.
- Term release scope: deterministic path only (**no Monte Carlo**).
- Product capability behavior:
  - Term hides Economic Scenario and Monte Carlo controls.
- Mortality defaults:
  - Term defaults to **US SSA 2015 period** (sex-specific), with options/default selected through registry.
- Parity invariants remain enforced (tie-break epsilon policy, tolerance policy, step-level parity).

# Immediate Next Steps

- Optional cleanup: rename SPIA-specific function/test symbol names (e.g., `price_spia_*`) to neutral naming if desired; currently they remain for continuity while files are pricing-agnostic.
- Optional product expansion: implement `WHOLE_LIFE` / `VARIABLE_ANNUITY` adapters and workbook paths using the now-centralized product config seams.
- Optional release governance: if preparing release notes, document naming migration (`spia_*` file removal) and new canonical launch path (`run_pricing_ui.bat`).

# Unresolved Bugs / Pending Calculations

- No known failing tests or active parity discrepancies at handoff.
- No unresolved runtime bug identified during the latest migration and validation passes.
- Pending calculations: none required for current scope; next work is feature expansion/refinement rather than defect remediation.

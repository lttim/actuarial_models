# Handoff — SPIA / Term / RILA model (`annuity_model`)

## Tip of `main`

- **`16f7a04`** — repo: cross-platform Git hygiene + macOS launchers.
- `5bb89a5` — RILA + Excel formula validator + ALM export column fix
  (validator perf: 68 s → 0.4 s on RILA + ALM via per-formula strip cache and
  template dedup).
- `ae6f545` — earlier handoff snapshot.

Single source of truth: `Code_Sandbox/.git`. There is no longer a nested
`.git` inside `annuity_model/`.

## Cross-platform setup

- **macOS / Linux**: `cd annuity_model && ./bootstrap_macos.sh` creates
  `.venv`, installs `requirements.txt`, runs the full regression suite.
- **Windows**: `py -3 -m venv .venv && .venv\Scripts\Activate.ps1 && pip install -r requirements.txt`,
  then `.\run_pricing_ui.bat`.
- See `MACOS_HANDOFF.md` at the repo root for the full step-by-step.

## Verification gates (must all be green)

```
pytest tests/ tests/parity/ -q   # 150 passed
pytest tests/parity/ -v          # 21 passed (incl. RILA)
```

Last full run on Windows (Python 3.14, openpyxl latest): **150 passed in
~10–12 s**.

## What's in flight

| Area | Files | Notes |
|------|-------|-------|
| Excel formula validator | `excel_workbook_validator.py`, `tests/test_excel_export_validation.py` | Statically checks balanced parens, function arities, `#REF!`/`#NAME?` literals, trailing empty args, and **cross-sheet column existence** (incl. inside `INDIRECT(...)`). Cached strip + template dedup keeps a 75 000-formula RILA workbook under 1 s. |
| RILA (accumulation) | `rila_projection.py`, `build_rila_excel_workbook.py`, `product_registry.py`, `pricing_ui.py` | Annual point-to-point crediting + Excel `ModelCheck`. RILA puts `ExpTotalCF` in column **M**; the shared ALM helpers must be called with `liability_total_col="M"` (SPIA / Term default `"S"`). |
| ALM export | `alm_excel_ladder.py`, `build_pricing_excel_workbook.py` | `_alm_liability_pv_cell_formula` and `_write_alm_projection_sheet` now accept `liability_total_col` / `liability_discount_col`; `INDEX(Liabilities!$<col>:$<col>, ...)` resolves to a populated column for every product. |
| Profit decomposition (UI) | `pricing_ui.py` | Signed Altair waterfall — `_build_profit_decomposition_rows`, `_build_profit_waterfall_chart_df`, `_altair_profit_waterfall_chart`. |
| Pricing Run session keys | `pricing_run_form_state.py`, `pricing_ui.py` | `PRICING_RUN_NUMBER_INPUT_KEYS` + `_pricing_run_numeric_seeds` to dodge Streamlit "default + Session State" warnings. |

## Parity / release gates (do not skip)

- `docs/model_parity_contract.md` — SPIA/ALM tolerances, tie-break, epsilon policy.
- `docs/rila_parity_contract.md` — RILA Python ↔ Excel addendum.
- `docs/rila_product_spec.md` — RILA v1 product definition.
- `pytest tests/parity/ -v` must show 0.00 discrepancy before merge.
- If a builder changed Excel formulas: regenerate a workbook, open in Excel,
  confirm `ModelCheck` shows 0.00 and that no "Removed Records: Formula …"
  repair dialog appears.

## Critical invariants — never break these

1. **Single Git store.** `Code_Sandbox/.git` only.
2. **Python = Excel.** Step-level reconciliation, not just final surplus.
3. **Mandatory `validate_workbook_or_raise(wb)`** before every `wb.save(...)`.
4. **Right `liability_total_col` per product** when calling shared ALM helpers.
5. **Tie-break / epsilon policy** as in `actuarial-parity.mdc`.
6. **Launcher pairs**: every `.bat` ships with a matching `.sh` in the same
   commit. `.gitattributes` enforces LF / CRLF / +x.

## Suggested next steps (optional)

- Smoke-launch on the M5 MacBook Air and capture timings of
  `pytest tests/ tests/parity/` — Apple Silicon should run the suite faster
  than the Windows desktop.
- If a new product is added, also add `liability_total_col` guidance and a
  workbook to the `tests/test_excel_export_validation.py` end-to-end gate.
- Consider extracting `_strip_strings_and_brackets` and the template cache
  into a tiny utility module if a second tool ever needs static formula
  parsing.

## Open issues

- None recorded. Validator is green on every committed builder; cross-platform
  bootstrap is documented in `MACOS_HANDOFF.md` and `bootstrap_macos.sh`.

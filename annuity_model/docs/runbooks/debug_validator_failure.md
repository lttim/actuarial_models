# Runbook: debug an `excel_workbook_validator` failure

## When to run this

* `validate_workbook_or_raise(wb)` raised `ExcelWorkbookValidationError`.
* CI's parity-gate failed with `excel_workbook_validator` in the traceback.
* A new builder change exploded a previously-green workbook.

## Triage flow

```
ExcelWorkbookValidationError
   │
   ├── "Unknown function FOO" ─────────────► Add FOO to FUNCTION_ARITY in
   │                                         excel_workbook_validator.py if it
   │                                         is a real Excel function. Otherwise
   │                                         the formula has a typo.
   │
   ├── "Wrong arity for FOO: got N, expects M" ──► Builder bug; count the
   │                                         arguments in the formula string.
   │
   ├── "Unknown sheet name 'Foo'" ────────► Rename or create the sheet.
   │                                         Common cause: builder wrote
   │                                         "Liability" but reader expects
   │                                         "Liability_Cashflows".
   │
   ├── "Bad range syntax 'A1:B'" ──────────► Builder is omitting the end-row.
   │                                         Almost always a column-letter
   │                                         drift -- check
   │                                         liability_layouts.py.
   │
   ├── "RILA ALM_Projection references col S" ──► **MERGE BLOCKER.** Builder
   │                                         is using SPIA's column letter for
   │                                         RILA's liability. Fix:
   │                                         liability_layout_for("rila") ->
   │                                         column M.
   │
   └── "Cell A1 holds non-formula but starts with ="
                                          ─► openpyxl wrote the formula as a
                                             literal string; force-cast in the
                                             builder via ws[coord] = "=...".
```

## Procedure

1. **Capture the failing message verbatim.** The validator's exception
   includes sheet, cell, and the offending fragment. Do not paraphrase.

2. **Open the workbook with `data_only=False`.** The validator inspects
   formulas, not values; you must read the formula text:
   ```python
   from openpyxl import load_workbook
   wb = load_workbook("RILA.xlsx", data_only=False)
   print(wb["Liability_Cashflows"]["S5"].value)
   ```

3. **Cross-check the column registry.**
   ```python
   from liability_layouts import liability_layout_for
   layout = liability_layout_for("rila")
   print(layout)
   ```
   The validator's "RILA references col S" failure is a layout-vs-builder
   drift; the fix is always in the builder, never in the validator.

4. **Re-run validator manually with the workbook in hand:**
   ```python
   from excel_workbook_validator import validate_workbook
   issues = validate_workbook(wb)
   for issue in issues:
       print(issue)
   ```
   `validate_workbook` returns the issue list; `validate_workbook_or_raise`
   raises on the first.

5. **Fix the builder, regenerate, re-validate, re-test:**
   ```bash
   python -m pytest tests/parity -q
   python scripts/deep_smoke.py
   ```

## Failure-mode -> fix table

| Failure                                  | Likely root cause                              | Fix lives in                          |
|------------------------------------------|------------------------------------------------|---------------------------------------|
| Unknown function `XLOOKUP`               | New Excel-365 function not registered          | `excel_workbook_validator.FUNCTION_ARITY` |
| Wrong arity for `IF`                     | Builder stitched too few/many args             | builder string template               |
| Unknown sheet name                       | Builder typo or sheet renamed                  | `build_*_excel_workbook.py`           |
| RILA col S violation                     | Hardcoded `"S"` literal somewhere              | `liability_layouts.py` + builder      |
| Reference outside sheet range            | Off-by-one in `n_months` loop                  | builder loop bound                    |
| `=text` left as a string                 | openpyxl coerced from non-string               | force `ws[coord] = "=" + expr`        |
| Cycle detected (rare)                    | Self-referential formula                       | builder logic bug                     |

## Pitfalls

* **The validator does not run formulas.** It only checks structure. A
  workbook can pass the validator and still produce wrong numbers; that
  failure mode is caught by the parity tests, not the validator.
* **The validator is the ONLY thing standing between you and a bad workbook
  in production.** Never `try/except`-swallow it. Never weaken
  `MODELCHECK_TOL`.

## Related

* [investigate_parity_break.md](investigate_parity_break.md) -- if validator
  is green but parity is red.
* `.cursor/rules/excel-formula-safety.mdc` -- the canonical formula-safety
  rules.

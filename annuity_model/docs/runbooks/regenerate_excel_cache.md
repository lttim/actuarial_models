# Runbook: inspect generated Excel workbooks

## When to run this

* You generated a workbook with `build_*_excel_workbook.py` and a parity test
  reports bad `ModelCheck` wiring.
* `ModelCheck` cells point at missing or product-inappropriate formula rows.
* Diff between two consecutive runs shows formula or workbook-structure drift.

## Why this happens

Automated gates validate generated workbooks without launching a desktop
spreadsheet app. `openpyxl` is used for workbook structure and formulas, and
the parity suite checks Python snapshot values plus `ModelCheck` formula links.
Do not use `data_only=True` as an automated gate for freshly generated files:
formula caches may be missing or stale.

## Procedure

1. **Confirm the workbook has formulas** (not values):
   ```bash
   python - <<'PY'
   from openpyxl import load_workbook
   wb = load_workbook("SPIA.xlsx", data_only=False)
   ws = wb["ModelCheck"]
   for row in ws.iter_rows(min_row=5, max_row=12, max_col=2):
       for cell in row:
           print(cell.coordinate, repr(cell.value))
   PY
   ```
   Expect strings starting with `=` for the B column. If you see numbers, the
   workbook was never formula-built; this runbook does not apply.

2. **Validate formulas and cross-sheet references**:
   ```python
   from openpyxl import load_workbook
   from excel_workbook_validator import validate_workbook_or_raise
   wb = load_workbook("SPIA.xlsx", data_only=False)
   validate_workbook_or_raise(wb)
   ```

3. **Check ModelCheck links**:
   `tests/parity/test_excel_recalc_per_product.py` asserts the Python
   snapshot cells match the engine and the formula cells link to canonical
   `Liabilities!X*` summary rows. Run it directly when debugging workbook
   drift:
   ```bash
   python -m pytest tests/parity/test_excel_recalc_per_product.py -q
   ```

## Pitfalls

* **Never** read with `data_only=True` from a workbook that was not
  recalculated by an end-user spreadsheet app. The values are not what the
  formulas would produce; they are whatever was cached at last save.
* **The Streamlit UI's "Run pricing" button** uses `data_only=False` plus
  Python snapshots, so it is unaffected by formula-cache absence.

## Related

* [debug_validator_failure.md](debug_validator_failure.md) -- if validation fails.
* [investigate_parity_break.md](investigate_parity_break.md) -- if parity values
  don't match Python.

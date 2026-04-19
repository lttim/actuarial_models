# Runtime Excel recalc parity gate

## Status

**Active** since 2026-04-18 (P0 hardening, this PR).

The gate is implemented as
[`tests/parity/test_runtime_excel_recalc.py`](../../tests/parity/test_runtime_excel_recalc.py)
plus the helper [`excel_runtime_recalc.py`](../../excel_runtime_recalc.py).
It builds a small SPIA workbook and recalculates it via
**LibreOffice headless** (`soffice --headless --calc --convert-to xlsx`),
then asserts the cached `ModelCheck` cells match the Python pricing result
within `parity_constants.MODELCHECK_TOL`.

## Why LibreOffice instead of a pure-Python evaluator?

Pure-Python alternatives all have caveats:

* `xlcalculator==0.5.0` transitively pins `yearfrac<2`, incompatible with
  the project's `numpy==2.4.4` lock. Parked since 2026-04 (history below).
* `formulas==1.3.4` installs cleanly but takes >3 minutes to load and
  recalculate the SPIA workbook because it builds the full dependency
  graph for thousands of cells. Unusable in CI on every PR.
* `pycel` has narrower function coverage and a similar perf profile.

LibreOffice is the same engine real users open these workbooks in, install
size on Linux is ~250 MB via apt, and a SPIA recalc completes in <5
seconds.

## Local install

* **macOS:**
  ```bash
  brew install --cask libreoffice
  ```
  (or download from https://www.libreoffice.org/download/)
* **Linux (Debian/Ubuntu):**
  ```bash
  sudo apt-get install -y libreoffice-calc
  ```
* **Windows:**
  ```cmd
  winget install TheDocumentFoundation.LibreOffice
  ```

After installing, ensure `soffice` is on `PATH`. As a fallback, set
`$LIBREOFFICE_SOFFICE` to its absolute path; the helper will pick it up.

If LibreOffice is **not** installed locally, the test
`test_modelcheck_cells_recalc_to_python_values` skips with a clear
install hint -- contributors are NOT blocked. CI installs it on the
parity-gate runner so the gate is enforced upstream.

## CI

The parity-gate workflow ([`.github/workflows/parity-gate.yml`](../../../.github/workflows/parity-gate.yml))
installs `libreoffice-calc` on Ubuntu before running the parity suite, so
the recalc test executes on every PR. The parity-gate workflow is also
listed in [`.github/branch-protection.json`](../../../.github/branch-protection.json)
as a required status check.

## What this gate catches

The static parity stack already covers a lot:

* `tests/parity/` -- Python pricing vs. a Python re-implementation of the
  Excel formula evaluation order (`excel_formula_sim.py`).
* `tests/test_excel_export_validation.py` -- per-product AST validator
  that parses every emitted formula.
* `tests/test_validator_invariants.py` -- meta-test that every
  `wb.save(...)` is preceded by `validate_workbook_or_raise(...)`.
* `scripts/deep_smoke.py` -- builds + validates every product workbook on
  every CI matrix entry.

The gap that remained -- and that this gate now closes -- is the case
where the *emitted formula string* differs from what `excel_formula_sim`
expects AND Excel itself would compute a different answer. The static
validator catches structural breaks; this runtime gate catches semantic
ones.

## Audit trail (most recent first)

* **2026-04-18: restored.** Switched to LibreOffice headless via
  `excel_runtime_recalc.py`. Test re-enabled. Quarterly review cadence
  retired (this is the steady-state).
* 2026-04-18: parked. `xlcalculator==0.5.0` still pulls `yearfrac<2`,
  conflicts with `numpy==2.4.4`. Next review due 2026-Q2.

## Related

* [`excel_runtime_recalc.py`](../../excel_runtime_recalc.py) -- LibreOffice
  helper module.
* [`tests/parity/test_runtime_excel_recalc.py`](../../tests/parity/test_runtime_excel_recalc.py)
  -- the gate.
* [`scripts/parity_trace.py`](../../scripts/parity_trace.py) -- uses the
  same helper to populate the Excel side of the diff trace when soffice
  is available; emits NaN otherwise so the gap is visible.

# Runtime Excel recalc parity gate

## Status

**Active in CI / opt-in on macOS local machines** since 2026-05-03.

The gate is implemented as
[`tests/parity/test_excel_recalc_per_product.py`](../../tests/parity/test_excel_recalc_per_product.py)
plus the helper [`excel_runtime_recalc.py`](../../excel_runtime_recalc.py).
It builds per-product workbooks and, where enabled, recalculates them via
**LibreOffice headless** (`soffice --headless --calc --convert-to xlsx`),
then asserts the cached `ModelCheck` cells match the Python pricing result
within `parity_constants.MODELCHECK_TOL`.

On macOS developer machines this runtime layer is skipped by default because
LibreOffice can surface desktop crash/reopen dialogs when automated headless
Calc exits unexpectedly. To opt in for a single controlled local run:

```bash
ANNUITY_MODEL_ENABLE_MACOS_LIBREOFFICE_RECALC=1 \
  annuity_model/.venv/bin/python -m pytest annuity_model/tests/parity/test_excel_recalc_per_product.py -q
```

The helper serializes all LibreOffice recalc invocations across local
processes via a host-wide lock and passes a per-run
`-env:UserInstallation=file://...` profile. This prevents the concurrent
subagent/pytest collision that can make macOS LibreOffice unstable.

## Why LibreOffice instead of a pure-Python evaluator?

Pure-Python alternatives all have caveats:

* `xlcalculator==0.5.0` transitively pins `yearfrac<2`, incompatible with
  the project's `numpy==2.4.4` lock. Parked since 2026-04 (history below).
* `formulas==1.3.4` installs cleanly but takes >3 minutes to load and
  recalculate generated workbooks because it builds the full dependency
  graph for thousands of cells. Unusable in CI on every PR.
* `pycel` has narrower function coverage and a similar perf profile.

LibreOffice remains the pragmatic CI answer: it is the same engine many users
open these workbooks in, install size on Linux is ~250 MB via apt, and a SPIA
recalc completes quickly on the Ubuntu parity runner.

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

If LibreOffice is **not** installed locally, or if the machine is macOS and
`ANNUITY_MODEL_ENABLE_MACOS_LIBREOFFICE_RECALC` is not set, the runtime recalc
tests skip with a clear message. Contributors are NOT blocked. CI installs
LibreOffice on Linux so the runtime gate is enforced upstream without macOS
desktop dialogs.

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

* **2026-05-03: macOS guarded.** Added host-wide serialization, explicit
  LibreOffice user-profile isolation, and a macOS local opt-in guard to avoid
  desktop crash/reopen dialogs during parallel agent or pytest runs.
* **2026-04-18: restored.** Switched to LibreOffice headless via
  `excel_runtime_recalc.py`. Test re-enabled. Quarterly review cadence
  retired (this is the steady-state).
* 2026-04-18: parked. `xlcalculator==0.5.0` still pulls `yearfrac<2`,
  conflicts with `numpy==2.4.4`. Next review due 2026-Q2.

## Related

* [`excel_runtime_recalc.py`](../../excel_runtime_recalc.py) -- LibreOffice
  helper module.
* [`tests/parity/test_excel_recalc_per_product.py`](../../tests/parity/test_excel_recalc_per_product.py)
  -- the gate.
* [`scripts/parity_trace.py`](../../scripts/parity_trace.py) -- uses the
  same helper to populate the Excel side of the diff trace when soffice
  is available; emits NaN otherwise so the gap is visible.

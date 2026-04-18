# Runtime Excel recalc parity gate (parked)

## Status

**Parked** since 2026-04-18.

The dev dep `xlcalculator==0.5.0` is commented out in
[`requirements-dev.txt`](../../requirements-dev.txt) and
[`tests/parity/test_runtime_excel_recalc.py`](../../tests/parity/test_runtime_excel_recalc.py)
self-skips via `pytest.importorskip("xlcalculator")`. The full parity gate
(static formula simulation in `tests/parity/excel_formula_sim.py` plus the
AST-walking validator in `excel_workbook_validator.py`) is unaffected.

## Why it was parked

`xlcalculator==0.5.0` (latest on PyPI as of 2026-04) transitively depends on
`yearfrac<2`, which is incompatible with the pinned `numpy==2.4.4` in
[`requirements.lock`](../../requirements.lock). The first push of the P0-P4
hardening commit failed `ci.yml` and `docs.yml` at the dependency-install step
with `ResolutionImpossible`. Locally the conflict was masked because the dev
venv was created before `xlcalculator` was added, so `pip install` had never
been re-run against the dev requirements.

## Restore plan (in priority order)

1. **Wait for an upstream `xlcalculator` release** that drops or relaxes the
   `yearfrac` dependency. Check [PyPI](https://pypi.org/project/xlcalculator/)
   periodically; track upstream issue at
   `https://github.com/bradbase/xlcalculator/issues`.
2. **Switch to a different pure-Python Excel evaluator** that does not pull
   `yearfrac<2`. Candidates:
   - [`formulas`](https://pypi.org/project/formulas/) (Vincenzo Arcidiacono).
     Larger surface; supports many more functions. Pulls `numpy`, no
     `yearfrac` cap last we checked.
   - [`pycel`](https://pypi.org/project/pycel/). Smaller, focused on
     compilation; less coverage of date/time functions.
3. **Vendor a `yearfrac` fork** that lifts the `numpy<2` cap. Only do this if
   options 1 and 2 are both blocked -- vendoring numerical code is a parity
   risk and adds maintenance burden.

## How to restore

When one of the above unblocks:

1. Re-add the dep to [`requirements-dev.txt`](../../requirements-dev.txt):

   ```text
   xlcalculator==<new version>     # OR
   formulas==<version>              # plus refactor of the test
   ```

2. Recreate the dev venv from scratch and regenerate the lockfile:

   ```bash
   rm -rf .venv
   python3.12 -m venv .venv
   ./.venv/bin/python -m pip install --upgrade pip
   ./.venv/bin/python -m pip install -r requirements.txt -r requirements-dev.txt
   ./.venv/bin/python -m pip freeze | grep -v '^-e ' | sort > requirements.lock
   ```

3. Run the gate locally:

   ```bash
   ./.venv/bin/python -m pytest tests/parity/test_runtime_excel_recalc.py -v
   ```

4. Update this runbook to **Active** and add a CHANGELOG entry under `Fixed`.

## Why we still have a parity story without it

The runtime recalc gate was a P4 *belt-and-braces* check on top of the static
parity stack. The remaining layers all run unconditionally:

- `tests/parity/` -- Python pricing vs. a Python re-implementation of the
  Excel formula evaluation order (`excel_formula_sim.py`).
- `tests/test_excel_export_validation.py` -- per-product AST validator that
  parses every emitted formula and checks references, range shapes, and
  cross-sheet wiring.
- `tests/test_validator_invariants.py` -- meta-test that every `wb.save(...)`
  is preceded by `validate_workbook_or_raise(...)` and every implemented
  product has a `LIABILITY_LAYOUTS` entry.
- `scripts/deep_smoke.py` -- builds + validates every product workbook on
  every CI matrix entry.

The gap closed only by the recalc gate is "the builder emits a formula string
that differs from what `excel_formula_sim` expects, and Excel itself would
disagree." That is real but narrow; the AST validator catches structural
breaks, and parity tolerances at `MODELCHECK_TOL` catch numerical drift if it
ever materialises in a real workbook.

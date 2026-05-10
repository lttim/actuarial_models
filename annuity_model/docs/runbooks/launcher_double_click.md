# Runbook: "I double-clicked `run_pricing_ui.command` and got an error"

This runbook covers the Finder / Explorer double-click path. For the standard
`streamlit run src/annuity_model/pricing_ui.py` developer flow, see the project README instead.

## Symptom catalog

| What you see                                                                              | Most likely cause                                        | Jump to     |
| ----------------------------------------------------------------------------------------- | -------------------------------------------------------- | ----------- |
| `TypeError: dataclass() got an unexpected keyword argument 'slots'`                       | Launcher picked up a Python < 3.10 (e.g. macOS 3.9)      | [P-OLD]     |
| `[ERROR] $PY is Python 3.X.Y; this project requires >= 3.11`                              | Launcher refused an old interpreter (working as intended) | [P-OLD]     |
| `[ERROR] streamlit is not importable ... Refusing to pip install ... (PEP 668)`           | No project `.venv`; running against system Python        | [NO-VENV]   |
| `[ERROR] No usable Python interpreter on PATH`                                            | No Python at all                                         | [NO-PY]     |
| `[ERROR] Failed to import annuity_model.pricing_ui ...`                                   | Module-level regression broke import                     | [BAD-IMPORT] |
| Terminal window closes the moment the script ends                                         | (Should not happen anymore -- file a bug)                | [BUG]       |

## Triage flow

1. Look at the line above `[ERROR]` in the Terminal window. The launcher
   keeps the window open with a `Press Return to close...` prompt on any
   non-zero exit.
2. Match the symptom to the table.
3. Apply the fix.
4. Re-run by double-clicking again.

## [P-OLD] Wrong / too-old Python

The launcher prefers, in order:

1. `./.venv/bin/python` (or `.venv\Scripts\python.exe` on Windows)
2. The currently-active virtualenv (`$VIRTUAL_ENV`)
3. `python3` / `python` on `PATH`

If options 1 and 2 are missing it falls back to (3), which on stock macOS is
the Xcode CommandLineTools 3.9. The fix is to create the project venv with a
modern Python:

```bash
cd annuity_model
python3.12 -m venv ./.venv
./.venv/bin/python -m pip install -r requirements.txt
```

The minimum supported version lives in `annuity_model/pyproject.toml` under
`[project].requires-python`. `tests/test_launcher_invariants.py` enforces
that the launchers reference the same value.

## [NO-VENV] No project `.venv`

Same fix as [P-OLD]: create `./.venv` with a Python that satisfies
`requires-python`. The launcher will not `pip install` into a system Python
(would violate PEP 668 on macOS / Debian-derived distros).

## [NO-PY] No Python at all

Install Python 3.11+:

- macOS: `brew install python@3.12` *or* the official installer from
  <https://www.python.org/downloads/macos/>.
- Windows: <https://www.python.org/downloads/> -- on the first installer
  screen, enable **Add python.exe to PATH**.

Then follow [NO-VENV].

## [BAD-IMPORT] `pricing_ui` won't import

The launcher imports `pricing_ui` *before* launching Streamlit specifically
to surface this error early. Re-run with full traceback:

```bash
cd annuity_model
PYTHONPATH=src ./.venv/bin/python -c "import annuity_model.pricing_ui"
```

Common causes:

- A `requirements.txt` change wasn't installed: rerun
  `./.venv/bin/python -m pip install -r requirements.txt`.
- A code-level regression at module load. Bisect with `git log -- pricing_ui.py`
  or the modules it imports (`product_registry`, `liability_layouts`, ...).

## [BUG] Terminal closes immediately

`run_pricing_ui.command` is required to keep Terminal open on non-zero exit
(`tests/test_launcher_invariants.py::test_command_launcher_holds_terminal_on_error`).
If it doesn't, the file has been edited without re-running pre-commit or the
test suite. Restore the `read -r` block at the bottom of the file and add a
regression note to `docs/CHANGELOG.md`.

## How this is enforced going forward

- `[project].requires-python` in `annuity_model/pyproject.toml` -- single
  source of truth.
- `tests/test_launcher_invariants.py` -- 16 checks: pyproject pin, launcher
  pin alignment, presence of each required guard, executable bits, and an
  end-to-end `--self-check` run under a stripped `PATH`.
- Pre-commit hook `launcher-invariants` -- runs the meta-tests when
  `pyproject.toml` or any launcher file changes.
- CI step `Launcher self-check` in `.github/workflows/ci.yml` -- runs the
  shell launcher on macOS / Linux runners and the batch launcher on Windows
  on every PR, in a freshly built `.venv`.

If a future regression slips past **all** of those, add the missed gap as a
new check in `tests/test_launcher_invariants.py` first; the fix to the
launcher comes after.

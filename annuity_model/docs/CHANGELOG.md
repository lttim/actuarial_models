# Changelog

All notable changes to `annuity_model` are documented here. The format is based
on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Parity-impacting changes (tolerances, mortality tables, curve construction,
ALM rules, RILA crediting) MUST also be logged in
[model_change_log.md](model_change_log.md).

## [Unreleased]

### Fixed
- CI dependency-resolution failure on `main` after the P0-P4 hardening commit:
  `xlcalculator==0.5.0` (dev dep) transitively required `yearfrac<2`, which
  conflicted with the pinned `numpy==2.4.4` in `requirements.lock`. Both
  `ci.yml` and `docs.yml` failed at the install step on first push. Locally
  the conflict was masked because the `.venv` predated the `xlcalculator`
  addition. Resolution: parked `xlcalculator` (commented out in
  `requirements-dev.txt`); the corresponding parity test
  `tests/parity/test_runtime_excel_recalc.py` already self-skips via
  `pytest.importorskip`. See
  `docs/runbooks/runtime_excel_recalc_gate.md` for the restore plan.
- `.github/workflows/docs.yml` now installs from `requirements.lock` instead
  of the loose `requirements.txt + requirements-dev.txt`, so docs builds are
  reproducible and immune to upstream transitive-dep drift (e.g. the
  `contourpy 1.2.0` backtrack that broke the first run).
- `run_pricing_ui.command` (Finder double-click) crashed on macOS systems
  whose default `python3` was Python 3.9 with a stray `streamlit` install in
  `~/Library/Python/3.9/site-packages`: `liability_layouts.py` uses
  `dataclass(slots=True)` (3.10+), so the app died at import with
  `TypeError: dataclass() got an unexpected keyword argument 'slots'`, and
  Terminal then closed the window before the user could read the trace.
  See `docs/runbooks/launcher_double_click.md` for the new triage flow.

### Added
- `[project].requires-python = ">=3.11"` in `annuity_model/pyproject.toml` --
  single source of truth for the minimum supported interpreter, mirrored by
  the launchers and CI matrix.
- `run_pricing_ui.sh` / `run_pricing_ui.bat` now (1) prefer the project
  `.venv`, (2) refuse interpreters older than `requires-python`, (3) refuse
  to `pip install` into a system Python (PEP 668), (4) import-smoke
  `pricing_ui` itself before launching streamlit, and (5) support
  `--self-check` for CI/pre-commit gating.
- `run_pricing_ui.command` now `read`s before exit on non-zero status so the
  user sees the error before macOS closes the Terminal window.
- `tests/test_launcher_invariants.py` -- meta-test suite that locks the
  launcher contract (version alignment, required guards, executable bits,
  end-to-end self-check under a stripped PATH).
- CI step `Launcher self-check` runs `./run_pricing_ui.sh --self-check` (or
  `run_pricing_ui.bat --self-check` on Windows) under a clean shell on every
  matrix entry.
- Pre-commit hook `launcher-invariants` re-runs the meta-tests whenever
  `pyproject.toml` or any launcher file changes.
- `docs/runbooks/launcher_double_click.md` -- triage guide for "I
  double-clicked the .command and got an error".
- `tests/test_validator_invariants.py` -- AST-walking meta-tests that
  enforce the `validate_workbook_or_raise` precedes-`wb.save` invariant
  and the layout-coverage invariant.
- `tests/test_property_invariants.py` -- Hypothesis property tests for
  SPIA single-premium positivity / monotonicity, RILA cap/floor bounds and
  monotonicity, and Term zero-q_x invariant.
- `tests/test_perf_baselines.py` -- pytest-benchmark gates for SPIA / RILA
  workbook builds and validator wall-time.
- `tests/parity/test_runtime_excel_recalc.py` -- xlcalculator-based Excel
  recalc parity gate (skipped when xlcalculator is not installed).
- `_observability.py` -- optional OpenTelemetry hook decorator.
- `Dockerfile` and `Justfile` -- reproducible builds and one-line task
  recipes (`just bootstrap`, `just test`, `just smoke`, ...).
- `.github/workflows/security.yml` -- weekly pip-audit + bandit + gitleaks.
- `.github/workflows/docs.yml` -- mkdocs build + GitHub Pages deploy.
- `parity_constants.py` -- single source of truth for all parity tolerances
  and Excel-formula epsilons.
- `scripts/render_parity_contract.py` -- regenerates the tolerance tables in
  `docs/model_parity_contract.md` and `docs/rila_parity_contract.md` from
  `parity_constants.py`. CI verifies via `--check`.
- `scripts/parity_trace.py` -- python-vs-excel CSV trace for parity-debug
  workflows; cited from `docs/runbooks/investigate_parity_break.md`.
- `liability_layouts.py` -- registry of Excel column letters per product,
  killing the `S`/`M` magic-letter trap.
- `liability_dispatch.py` -- replaces `isinstance` chains in
  `pricing_projection.run_alm_projection_from_pricing_result`.
- `excel_builder_helpers.py` -- public surface for shared builder utilities,
  closing `_private` cross-imports between SPIA / Term / RILA builders.
- `_logging.py` -- structured logging module replacing `print()` calls in the
  engine.
- `annuity_model/__init__.py` -- public API surface (engines, Excel pipeline,
  registry, logging).
- `annuity_model/ui/` -- decomposition target for `pricing_ui.py`; see
  `ui/MIGRATION.md`.
- `docs/glossary.md` -- one-paragraph definitions for every actuarial /
  engineering term used in the codebase.
- `docs/runbooks/{regenerate_excel_cache,debug_validator_failure,investigate_parity_break,release}.md`.
- `docs/CODEOWNERS_RATIONALE.md`.
- Top-level `README.md` and `CONTRIBUTING.md`.
- `pyproject.toml` with tool stanzas for ruff, black, mypy, pytest, coverage.
- `.pre-commit-config.yaml` with ruff/black/mypy gates plus import-smoke hooks.
- `.github/workflows/{ci,parity-gate}.yml`, `.github/CODEOWNERS`,
  `.github/dependabot.yml`.
- `requirements.lock` -- frozen transitive dependency tree.
- New parity tests for the pro-rata reinvestment path.
- `scripts/deep_smoke.py` (promoted from `.smoke/deep_smoke.py`).

### Changed
- Pinned all runtime dependencies in both `requirements.txt` files.
- RILA parity test tolerance tightened from `1e-3` to `1e-4` to match
  `docs/rila_parity_contract.md`.
- `parity_test_checklist.md` reconciled to `1e-4` (was `0.01`).
- `product_registry.get_product_adapter` and `get_pricing_metrics` refactored
  to dictionary dispatch (no behaviour change).
- Builder modules now read column letters from `liability_layout_for(<code>)`
  instead of hard-coded `"S"` / `"M"` literals.

### Removed
- `pytest` and `pytest-html` from runtime `requirements.txt` (now dev-only
  via `requirements-dev.txt`).

### Fixed
- Stale "column S" comment in `build_term_excel_workbook.py`.
- `bootstrap_macos.sh` import-smoke step missing `import importlib.util`.

## [0.1.0] - 2026-03-15
- Initial parity-gated release of SPIA / Term Life / RILA pricing engine and
  Excel workbook generator.

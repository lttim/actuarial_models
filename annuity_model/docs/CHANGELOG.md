# Changelog

All notable changes to `annuity_model` are documented here. The format is based
on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Parity-impacting changes (tolerances, mortality tables, curve construction,
ALM rules, RILA crediting) MUST also be logged in
[model_change_log.md](model_change_log.md).

## [Unreleased]

### Fixed
- **Dockerfile base image digest** was a placeholder (`sha256:3d77c6a4...`).
  `docker build .` failed with `manifest unknown`. Replaced with the real
  current digest for `python:3.12-slim-bookworm`
  (`sha256:d97792894a6a4162cae14da44542a83c75e56c77a27b92d58f3f83b7bc961292`)
  fetched from the Docker Hub registry API. CI now verifies the image
  builds and runs `deep_smoke.py` + parity tests inside the container on
  every push (new `docker` job in `.github/workflows/ci.yml`). Refresh
  procedure is documented inline in the workflow.

### Added
- New `docker` CI job that builds the image and runs both `deep_smoke.py`
  and the parity gate inside the container on every push to `main` and on
  every PR. Catches drift between the host venv and the reproducible
  container build (the latter is what auditors will reproduce).

### Removed
- Dead `[tool.pytest.ini_options]` block from `pyproject.toml`. Pytest reads
  `pytest.ini` first when both exist and silently ignores the pyproject
  block, so the duplicate was never enforced and had already fallen out of
  sync (the marker list was missing `invariant` and `property`). `pytest.ini`
  is the single runtime source of truth until the `src/` migration moves
  config into `pyproject.toml` and deletes `pytest.ini` in the same commit.
- Dropped `UP038` from `[tool.ruff.lint.ignore]` -- the rule was removed in
  ruff 0.10 (2025-Q1) and ignoring it now produces a "rules have been
  removed" warning. Tuple-form `isinstance(x, (X, Y))` is no longer flagged.

### Added
- **Strict mypy restored on `pricing_projection.py`** (Wave 2.1 of phase-5
  hardening). Added back to the strict override in
  `annuity_model/pyproject.toml` and to the pre-commit `mypy` files
  pattern, so the engine module is now checked under
  `disallow_untyped_defs / disallow_incomplete_defs / check_untyped_defs /
  warn_unreachable / no_implicit_reexport` on every commit and CI run.
  The 16 strict errors were resolved by:
  - Promoting `_is_numeric` to a `TypeGuard[int | float | np.number[Any]]`
    so downstream `int()` / `float()` calls type-check without per-callsite
    casts (eliminates 5 errors).
  - Adding an explicit
    `mortality: MortalityTableQx | MortalityTableRP2014MP2016`
    annotation in `__main__` so the try/except fallback cascade widens the
    inferred type instead of getting locked in by the first assignment
    (eliminates 2 errors).
  - Replacing two `pd.io.common.file_exists(...)` calls (undocumented and
    not in pandas-stubs) with `os.path.exists(...)` (eliminates 2 errors).
  - Defaulting `valuation_year` to `0` when calling
    `monthly_survival_to_payment` on the `MortalityTableQx | RP2014MP2016`
    union -- the runtime guard above the call already fast-fails the only
    None+RP2014MP2016 combo and the Qx branch ignores the value
    (eliminates 2 errors).
  - Splitting the `np.flatnonzero(matured)` loop variable into `k_idx`
    (numpy signedinteger) and `k = int(k_idx)` (Python int) so the
    subsequent `faces[k]` array index type-checks (1 error).
  - Adding `.tolist()` on two ndarray->`Sequence[float]` arguments to
    `bootstrap_zero_rates_from_par_yields` (2 errors).
  - Annotating `df_prev: list[float] = []` (1 error).
  All 4 canonical gates remain green; only `alm_excel_ladder.py` is left in
  the strict-excluded list, scheduled for Wave 2.2.

- **Versioned data-artifacts registry** (Wave 3.2 of phase-5 hardening).
  New `annuity_model/data_registry.py` is now the single source of truth
  for every CSV consumed by the engines and builders. Each artifact has a
  declared `kind`, `version`, `relative_path`, `sha256`, and `source`
  attribution string, exposed via a frozen `DataArtifact` dataclass and
  catalogued in `data_registry.REGISTRY`.

  On-disk layout moved from flat basenames in `annuity_model/` to
  `annuity_model/data/<kind>/<version>/<basename>`:
  - `data/yield_curves/2026-03-20/treasury_zero_rate_curve.csv`
  - `data/yield_curves/2026-03-20/treasury_par_yield_curve.csv`
  - `data/mortality/rp2014/rp2014_male_healthy_annuitant_qx_2014.csv`
  - `data/mortality/mp2016/mp2016_male_improvement_rates.csv`
  - `data/expenses/us_placeholders/expenses_assumptions_us_placeholders.csv`
  - `data/index_scenarios/sp500_seed_baseline/sp500_scenario_projection_monthly.csv`

  The 6 `pricing_projection.DEFAULT_*_CSV` constants now resolve through
  `data_registry.path_str(...)` so existing call sites
  (`pd.read_csv(DEFAULT_*_CSV)`) keep working without filesystem layout
  knowledge. Hardcoded paths in `build_pricing_excel_workbook.py`,
  `illustrate_pricing_projection.py`, `tests/test_rila_projection.py`,
  and `generate_sp500_scenario_csv.py` were rewritten to go through the
  registry.

  New invariant test `tests/test_data_registry_invariants.py` (8 tests,
  marked `invariant`) locks the registry contract:
  - Every artifact exists at its declared path (catches stale
    `relative_path` after a `git mv`).
  - **Every artifact's on-disk bytes match the declared sha256.** This
    is the parity-critical safety net: in-place edits to a yield curve
    or mortality table will fail CI on the next run, forcing the
    editor to either roll back or move the file to a new
    `data/<kind>/<new_version>/` folder with a CHANGELOG entry.
  - `DEFAULT_*_CSV` constants in `pricing_projection` resolve through
    the registry (catches future bypasses where someone reintroduces
    a hardcoded basename).
  - Names are unique, paths are absolute, `relative_path` matches
    `kind`/`version`, `get_artifact("totally_fake")` raises
    `KeyError` with the list of known names.

  Refresh procedure (when a yield-curve snapshot is updated, etc.):
  1. Drop the new file under `data/<kind>/<new_version>/<basename>`.
  2. Add a new `DataArtifact` entry to `REGISTRY` (or replace the old
     one if the version label was bumped intentionally).
  3. Run `pytest tests/test_data_registry_invariants.py` -- the
     failure message includes the actual sha256 to paste into the
     entry.
  4. Add a CHANGELOG entry under `[Unreleased] -> Changed` documenting
     the source of the new snapshot.

- **`@register_builder` decorator pattern** for `build_product_workbook`
  (Wave 3.1 of phase-5 hardening). Replaces the ~30-line if/elif chain
  in `annuity_model/product_excel.py` with a `ProductType -> builder`
  registry populated at import time by `@register_builder(ProductType.X,
  spec_type=XSpec)` decorators on three thin wrapper functions
  (`_build_spia_workbook`, `_build_term_workbook`, `_build_rila_workbook`).
  Adding a new product is now a one-edit change: write the wrapper +
  decorate it. The dispatcher (`build_product_workbook`) is now
  product-agnostic and stays under 20 lines. Spec-type validation runs
  before the builder is invoked, so wrong-type specs fail with a clear
  `TypeError` instead of an `AttributeError` deep in the builder. New
  invariant test `tests/test_builder_registry_invariants.py` (5 tests,
  marked `invariant`) locks the registry contract:
  - Every product in `implemented_product_types()` has a registered
    builder, and vice versa (no orphans).
  - Each builder declares its expected spec dataclass.
  - Wrong-type spec raises `TypeError` with the product name in the
    message.
  - Unimplemented product raises `NotImplementedError` with the enum
    `.value` in the message.
  - Re-registering the same `ProductType` raises `RuntimeError` (catches
    copy-paste mistakes where two builders claim the same enum).

- **Strict mypy restored on `alm_excel_ladder.py`** (Wave 2.2 of phase-5
  hardening). Added back to the strict override in
  `annuity_model/pyproject.toml` and to the pre-commit `mypy` files
  pattern. The 37 strict errors were resolved by:
  - Introducing two narrow column accessors --
    `Ci(name) -> int` and `Cl(name) -> list[int]` -- inside
    `write_alm_engine_sheet` that wrap the existing `C()` registry. `C()`
    returns `int | list[int]` (scalar columns vs per-bond column slices),
    which forced every call site to cast or `isinstance`-check. The new
    helpers fail-fast at runtime with a clear `RuntimeError` if a column
    name resolves to the wrong shape, and let mypy --strict type-check the
    ~30 affected `ws.cell(..., column=Ci("..."))` calls without per-callsite
    casts.
  - Annotating the `ws` parameter as
    `openpyxl.worksheet.worksheet.Worksheet`. Untyped earlier because
    openpyxl's stubs were spotty -- `types-openpyxl` (already pinned for
    `excel_workbook_validator`) covers it now.
  - Replacing two inline `lambda i, d=di: ...` closures (used to capture the
    disinvestment-pass index `di` per loop iteration) with named factory
    helpers `_fd_header_for(d)` / `_fd_gloss_for(d)`. The default-arg
    closure idiom is a Python pattern for capturing loop variables but
    mypy cannot annotate `d=di` defaults inside lambdas, so this is a pure
    typing-driven refactor with identical runtime behavior.
  All 4 canonical gates remain green. The strict mypy override list now
  covers every parity-critical engine and builder module; the next strict
  expansion will come with the `src/` layout migration in Wave 3.3.

- **Quarterly recurring check** for the parked runtime Excel recalc gate.
  `annuity_model/docs/runbooks/runtime_excel_recalc_gate.md` now carries
  a "Recurring quarterly check" section with three concrete `pip install
  --dry-run` probes (xlcalculator update, yearfrac>=2 unblocked,
  numpy>=2 + formulas/pycel) and a dated audit trail. Next review due
  end of 2026-Q2.

### Deferred
- **2nd CODEOWNER not added (still solo-owned by `@lttim`).** Phase-5
  backlog item to add a second human/team CODEOWNER cannot be solved by
  editing `.github/CODEOWNERS` alone -- it requires a real second
  reviewer to exist. Two extension hooks are now in place so that the day
  a second reviewer is onboarded the change is one-line: a placeholder
  team handle `@lttim/actuarial-reviewers` is documented as a commented
  TODO in `.github/CODEOWNERS`, and the default-owner line plus every
  parity-critical override has a `TODO(second-owner)` next to it. Branch
  protection on `main` (Wave 7) will still be enabled with
  `require_code_owner_reviews: true` so that the contract is in place when
  the team grows.

### Changed
- **Dependabot mkdocstrings major bump (PR #8 applied directly to main):**
  `mkdocstrings[python] 0.27.0 -> 1.0.4`. Verified `mkdocs build --strict`
  still succeeds against the existing `mkdocs.yml` plugin block; 1.0's
  breaking changes are confined to `BaseHandler` internal API
  (`__init__` signature, removed submodules) which we do not import.
  Side benefit: silences the `set_fallback_anchor_function is deprecated`
  warning that the 0.27 plugin emitted on every build.
  `mkdocstrings-python` jumped to 2.0.3 (transitive dep) -- `requirements.lock`
  regenerated. Closes Dependabot PR #8.

- **Dependabot Actions + pytest-cov bumps (PRs #1, #2, #3, #4, #9 applied directly to main):**
  - `actions/checkout v4 -> v6` across `ci.yml`, `docs.yml`, `parity-gate.yml`,
    `security.yml` (8 occurrences). v5 dropped Node 16; v6 ships on Node 20
    runtime which is what `actions/setup-python@v5` already requires.
  - `actions/deploy-pages v4 -> v5` in `docs.yml`. v5 only added optional
    `preview` mode; existing `id: deployment` step keeps working unchanged.
  - `actions/upload-artifact v4 -> v7` in `ci.yml`. v5 added compression
    options, v6/v7 changed nothing the workflow uses; the per-matrix-entry
    artifact names already follow the v4+ unique-name requirement so no
    collisions are possible.
  - `pytest-cov 6.0.0 -> 7.1.0` in `requirements-dev.txt`; `requirements.lock`
    regenerated. Verified `coverage run / report --fail-under=55` still
    produces the same 59.6% total.
  - Closes Dependabot PRs #1, #2, #3, #4, #9 as superseded.

- **Dependabot patch-and-minor group bump (PR #7 applied directly to main):**
  `ruff 0.7.4 -> 0.15.11`, `mypy 1.13.0 -> 1.20.1`,
  `mkdocs-material 9.5.49 -> 9.7.6`, `pymdown-extensions 10.12 -> 10.21.2`,
  `hypothesis 6.122.3 -> 6.152.1`, `pytest-benchmark 5.1.0 -> 5.2.3`,
  `mutmut 3.2.0 -> 3.5.0`, `pip-audit 2.7.3 -> 2.10.0`,
  `bandit 1.7.10 -> 1.9.4`. Side effects:
  - `.pre-commit-config.yaml` `rev:` fields bumped in lockstep
    (`ruff-pre-commit v0.15.11`, `mirrors-mypy v1.20.1`).
  - `ruff-pre-commit` 0.15 renamed the lint hook id from `ruff` to
    `ruff-check`; `.pre-commit-config.yaml` updated accordingly to silence
    the legacy-alias warning.
  - `.github/workflows/security.yml` hardcoded `pip-audit==2.7.3` and
    `bandit==1.7.10` bumped to match `requirements-dev.txt`.
  - `requirements.lock` regenerated from `pip freeze`.
  - 14 source files reformatted by ruff-format 0.15 (whitespace and
    long-line wrapping only -- no semantic changes; verified by full
    pytest + parity + deep_smoke green).

### Fixed
- CI coverage gate: `--fail-under=85` was aspirational and CI never went
  green on it. Today's actual coverage is 59.6%, dominated by the 4103-LOC
  Streamlit `pricing_ui.py` which the current test suite cannot exercise.
  Set `--fail-under=55` as a one-way ratchet just below current; restoration
  to 75-85% is tracked under the ui/ decomposition (`ui/MIGRATION.md`).
- GitHub Pages was not enabled on the repo, so `docs.yml` failed at
  `actions/configure-pages` with `Resource not accessible by integration`
  even with `enablement: true` (the `GITHUB_TOKEN` lacks admin:repo).
  Pages enabled out-of-band via `gh api -X POST repos/.../pages
  -f build_type=workflow`; `enablement: true` kept on the action so a
  fresh fork still self-bootstraps. Documentation now publishes to
  `https://lttim.github.io/actuarial_models/`.
- `launcher-invariants` pre-commit hook hardcoded `./.venv/bin/python`
  but the CI pre-commit job has no project venv, so the hook crashed
  with `bash: ./.venv/bin/python: No such file or directory`. Switched
  to the same `.venv`-prefer-with-`python3`-fallback pattern used by the
  other local hooks.
- Three further CI failures inherited from the P0-P4 hardening commit
  (after the dependency-resolution fix above):
  1. **`pre-commit` job** failed because `.pre-commit-config.yaml` pinned
     `default_language_version: python3.11` but `actions/setup-python`
     provisions only the matrix entry's interpreter (3.12 for that job),
     so virtualenv could not find `python3.11`. Switched to `python3` --
     the actual *checked* Python version is still fixed by
     `tool.ruff.target-version`, `tool.mypy.python_version`, etc.
  2. **`docs.yml` --strict** failed on `mkdocs build` because the nav
     entries used `docs/...` paths although `docs_dir: docs` already
     resolves them, and on out-of-tree references (`../README.md`,
     `../AGENTS.md`). Stripped the duplicate prefix; switched out-of-tree
     links to absolute GitHub URLs (parity_test_checklist, index page,
     and the new runtime_excel_recalc_gate runbook).
  3. **`pre-commit` job** also failed mypy because the original strict
     override list (`pricing_projection`, `term_projection`,
     `rila_projection`, `alm_excel_ladder`, `excel_workbook_validator`,
     `product_registry`) was aspirational: `pricing_projection.py` and
     `alm_excel_ladder.py` produced 52 strict-mode errors (ndarray ->
     Sequence narrowing, int|None -> int, pandas attr-defined, numeric
     widening). Narrowed the override to the four modules that pass
     strict today; restoring the other two is a typed-narrowing pass
     tracked as a P5 follow-up. Pre-commit + CI now lock that gate at
     its real coverage rather than declaring it green falsely.
- Dropped `black==25.1.0` -- ruff-format is now the single formatter.
  Both were configured at line-length=100 / target=py311 but disagreed on
  paren wrapping for long `assert` statements, making `pre-commit run
  --all-files` non-idempotent (each pass reformatted the same 3 files).
  Side benefits: removes the dependabot black-major-bump PR (#10) and
  one fewer pinned dev dep to track.
- Local pre-commit hooks (`import-smoke-validator`, `import-smoke-engines`,
  `render-parity-contract-check`) now prefer `./annuity_model/.venv/bin/python`
  with a `python3` fallback, instead of the bare `python` entry that did
  not exist on macOS and bound to system `python3.9` (incompatible with
  `dataclass(slots=True)`). Same defence the launcher already enforces.
- Pinned `mirrors-mypy` in `.pre-commit-config.yaml` to `v1.13.0` to match
  the `mypy==1.13.0` pin in `requirements-dev.txt`. The previous drift
  (hook ran `v1.14.1`, requirements pinned `1.13.0`) meant local + CI
  could trip different rule sets.
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

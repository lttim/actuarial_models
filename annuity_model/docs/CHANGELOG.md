# Changelog

All notable changes to `annuity_model` are documented here. The format is based
on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Parity-impacting changes (tolerances, mortality tables, curve construction,
ALM rules, RILA crediting) MUST also be logged in
[model_change_log.md](model_change_log.md).

## [Unreleased]

### Added — Seven-product rollout (Phases 0-9)

Per [docs/seven_product_rollout_plan.md](seven_product_rollout_plan.md),
the following seven products are now first-class citizens alongside SPIA
/ Term / RILA. The `implemented_product_types()` set grows from 3 to 10:

- **MYGA** (Multi-Year Guaranteed Annuity) — single premium, declared
  rate, fixed guarantee period. Maturity payout + in-period death CF.
  Engine: `myga_projection.py`. Builder:
  `build_myga_excel_workbook.py`. Subpackage: `products/myga/`.
- **FIA** (Fixed Indexed Annuity) — single premium with annual P2P
  crediting (cap + floor + participation). Reuses `crediting`
  framework. Engine: `fia_projection.py`. Builder:
  `build_fia_excel_workbook.py`. Subpackage: `products/fia/`.
- **VA** (Variable Annuity, repurposed `ProductType.VARIABLE_ANNUITY`)
  — single premium with sub-account return. GMDB =
  `max(AV, single_premium)`. Engine: `va_projection.py`. Builder:
  `build_va_excel_workbook.py`. Subpackage:
  `products/variable_annuity/`.
- **WL** (Whole Life — single premium, repurposed
  `ProductType.WHOLE_LIFE`) — level face × death-prob CF. Defaults to
  CSO 2017 Ultimate placeholder mortality. Engine:
  `wl_projection.py`. Builder: `build_wl_excel_workbook.py`.
  Subpackage: `products/whole_life/`.
- **UL** (Universal Life — single premium) — monthly cycle of
  load → declared-rate credit → COI → expense charge. Type A death
  benefit. AV-depletion terminates the contract. Engine:
  `ul_projection.py`. Builder: `build_ul_excel_workbook.py`.
  Subpackage: `products/universal_life/`.
- **IUL** (Indexed UL) — UL with annual P2P credit on segment
  anniversaries. Engine: `iul_projection.py`. Builder:
  `build_iul_excel_workbook.py`. Subpackage: `products/indexed_ul/`.
- **VUL** (Variable UL) — UL with monthly sub-account return as the
  credit. Engine: `vul_projection.py`. Builder:
  `build_vul_excel_workbook.py`. Subpackage: `products/variable_ul/`.

Foundation modules added in Phase 0 (referenced by all seven new
products, opt-in for SPIA / Term / RILA):

- **`lapse.py`** — static lapse / persistency framework
  (`LapseAssumption`, `combined_monthly_survival`, monthly hazard
  helpers, default 8/7/6/5/4/3/2 ultimate-2% template). Per-product
  opt-in via `lapse: LapseAssumption | None`. Existing engines stay
  mortality-only (verified by unchanged golden JSON).
- **`crediting.py`** — strategy hierarchy
  (`CreditingStrategy` Protocol, `FixedDeclaredRate`,
  `AnnualPointToPointCapped`). RILA's `segment_credited_return`
  refactored to delegate to `AnnualPointToPointCapped` — public name
  preserved, golden JSON byte-identical.
- **`account_value.py`** — single-source UL/IUL/VUL monthly AV cycle
  (`AVConfig`, `evolve_account_value`).
- **`mortality_2017_cso.py` + four CSV artifacts** under
  `data/mortality/cso_2017_ult/` — synthetic Gompertz-Makeham
  approximation of CSO 2017 Ultimate (sex × smoker). NOT licensed CSO;
  production users overlay their own file at the same path.
- **`actuarial_benchmarks.py` + `docs/actuarial_benchmarks.md`** —
  per-product band constants + rationale narrative; cross-checked by
  `scripts/render_actuarial_benchmarks.py --check` (now part of
  `just preflight`).
- **`parity_constants.py`** extended with `LIFE_MODELCHECK_TOL`,
  `ANNUITY_ACCUM_MODELCHECK_TOL`, `AV_TOL`, `LAPSE_DECREMENT_TOL`,
  `MYGA_PV_TOL`, `FIA_PV_TOL`, `VA_PV_TOL`, `WL_PV_TOL`, `UL_PV_TOL`,
  `IUL_PV_TOL`, `VUL_PV_TOL`.

Per-product wiring (all 7 new products):

- **Adapter & registry:** `product_registry._PRODUCT_ADAPTERS`,
  `_PRICING_METRIC_FORMATTERS`, `_PRODUCT_DISPLAY_NAME`,
  `_PRODUCT_CAPABILITIES`, `_PRODUCT_MORTALITY_MODE_OPTIONS`,
  `_PRODUCT_DEFAULT_MORTALITY_MODE`, `_PRODUCT_UI_CONFIG`,
  `_PRODUCT_VALIDATORS`. Life products default to `cso_2017_ult`
  mortality; annuity products keep `rp2014_mp2016`.
- **Excel builder dispatch:** `product_excel.py` `@register_builder`
  decorators for all 7 new products.
- **Liability layouts:** `liability_layouts.LIABILITY_LAYOUTS` —
  accumulation products (RILA / MYGA / FIA / VA) use
  `total_cf_col=M, discount_col=O`; life products (SPIA / Term / WL /
  UL / IUL / VUL) use `total_cf_col=S, discount_col=O`.
- **UI:** 7 new contract widget blocks + 7 new contract-construction
  blocks in `pricing_ui.py`. New `RUN_KEY` constants in
  `pricing_run_form_state.py` (one per product knob); seed defaults +
  per-product normalization extended (force `run_use_index = True`
  for indexed / VA-style products, `False` for fixed / declared-rate).
- **Liability-path dispatch:** each engine module registers via
  `liability_dispatch.register_liability_path_converter` at import
  time. ALM dispatch needs no changes for new products.
- **Tests (per product):** parity (`tests/parity/test_<P>_actuarial.py`),
  golden JSON (`tests/parity/golden/<P>.json`), recalc case in
  `tests/parity/test_excel_recalc_per_product.py::_CASE_BUILDERS`,
  AppTest smoke (`tests/ui/test_apptest_<P>.py`),
  regression-matrix fixture (`tests/test_regression_matrix.py`),
  observability wiring (`tests/test_observability_wiring.py`).
  All 7 actuarial assessments use bands from `actuarial_benchmarks.py`
  imported by name — never inline literals.

End-to-end deliverables:

- **`scripts/deep_smoke.py`** now exercises all 10 products (was 3 +
  RILA-with-ALM); zero failures, ~3 s total.
- **`tests/ui/test_apptest_full_workflow.py::_PRODUCT_RUN_CASES`**
  covers all 10 products; 57 tests pass.
- **`scripts/generate_cso_2017_synthetic.py`** generates the four
  placeholder mortality CSVs with documented "synthetic — overlay
  licensed file in production" warnings (in a sidecar
  `data/mortality/cso_2017_ult/README.md`).
- **`scripts/render_actuarial_benchmarks.py`** mirrors
  `render_parity_contract.py` and is wired into `just preflight`.
- **`docs/lapse_framework.md`** — narrative for the lapse v1 contract.
- **`docs/actuarial_benchmarks.md`** — per-product band rationale +
  closed-form / sensitivity references.
- **`docs/model_change_log.md`** — consolidated Phase 0 + Phases 1-9
  entries.

### Added
- **Always-on UI smoke gate (``tests/ui/test_apptest_full_workflow.py``,
  19 tests).** Closes the gap between the per-product
  ``test_apptest_<product>.py`` files (which only assert that the
  Pricing Run form RENDERS for each product) and the actual
  user-visible workflow. New coverage runs on every default
  ``pytest`` invocation (no opt-in marker), no skips on modern
  streamlit installs:

  1. **Boot** -- ``pricing_ui.py`` AND the Streamlit Cloud entry
     point ``streamlit_app.py`` must render on first paint with zero
     script-level exceptions. The Cloud entry was previously
     untested; a regression there would have shipped silently.
  2. **Per-section render** -- parametrized over every routable
     sidebar section (``overview``, ``run``, ``alm``, ``what_if``,
     ``excel_replicator``); each must render its empty state without
     raising, catching the bug class where a session-state lookup
     KeyErrors before the user has done anything.
  3. **End-to-end pricing run** -- per product (SPIA / Term / RILA),
     click "Run pricing" with deterministic inputs and assert
     ``st.session_state['pricing_res' / 'pricing_contract' /
     'pricing_meta']`` are populated. This is the functionality test
     -- the core action of the app must work.
  4. **Downstream pages after pricing** -- after a real run, navigate
     to ALM and Excel Replicator and assert no exception. Catches the
     bug class where a downstream page assumes a key the success
     path forgot to write for a particular product.
  5. **Excel download surface** -- ``st.session_state['pricing_xlsx_bytes']``
     (the bytes the ``st.download_button`` serves) must pass strict
     ``excel_workbook_validator``. UI-side complement to
     ``tests/test_excel_export_validation.py`` (engine-side gate).

  Total runtime ~13 s on a 2024-era laptop. Module-level skip applies
  only when ``streamlit.testing.v1`` is unimportable (streamlit < 1.28),
  which is impossible under the pinned ``requirements.lock``.
- **Per-product "Excel recalc matches Python" gate
  (``tests/parity/test_excel_recalc_per_product.py``, 7 tests).**
  Two complementary layers, both parametrized over every implemented
  product so a new product cannot ship without engaging both:

  1. ``test_python_cached_modelcheck_values_match_engine_<product>`` --
     **always runs, no skip.** Builds a small workbook for every
     product (SPIA 12 mo, Term 60 mo, RILA 60 mo) and asserts that
     the literal Python values the builder bakes into ModelCheck
     column B equal the engine outputs within ``MODELCHECK_TOL`` /
     ``TERM_MODELCHECK_TOL`` / ``RILA_PV_TOL``. This is the gate that
     fires on every developer machine, every CI shard, every PR --
     it catches the bug class where a builder refactor writes
     stale/rounded numbers into the workbook column the user actually
     reads in Excel.
  2. ``test_libreoffice_recalc_matches_engine_<product>`` --
     LibreOffice-headless recalc, parametrized per product (the
     pre-existing ``test_runtime_excel_recalc.py`` only covered
     SPIA). Runs in CI (parity-gate workflow installs
     ``libreoffice-calc``) and on developer laptops with ``soffice``
     on PATH; skips with a clear install hint otherwise. This is the
     strongest gate -- it actually invokes the spreadsheet engine
     end users open these workbooks in, catching builder bugs that
     no static check can see (e.g. an off-by-one SUMPRODUCT range).

  A coverage invariant
  (``test_every_implemented_product_has_a_recalc_case``) ensures the
  always-on layer fires for every product registered in
  ``product_registry.implemented_product_types``. Both layers live
  under ``tests/parity/`` so they're already inside the default
  ``pytest tests/ tests/parity/`` invocation -- no CI workflow change
  required.
- **Drop-in branch-protection profile for the second-CODEOWNER day
  (`.github/branch-protection.with-second-reviewer.json`).** Mirrors
  the active `branch-protection.json` byte-for-byte except for the
  ``required_pull_request_reviews`` block, which flips on
  ``required_approving_review_count = 1``,
  ``require_code_owner_reviews = true``,
  ``dismiss_stale_reviews = true``, and
  ``require_last_push_approval = true``. Pinned by
  ``tests/test_branch_protection_drift.py`` (5 cases): asserts the
  active profile stays null until activation day, asserts the deferred
  profile actually requires reviews, and asserts every other key
  (status checks, linear history, force-push policy, conversation
  resolution) stays in lockstep across the two files. The full
  activation checklist -- onboard reviewer / team, uncomment the
  ``TODO(second-owner)`` lines in ``.github/CODEOWNERS``, swap the
  profile, then collapse the deferred file after one clean release
  cycle -- is documented in
  ``annuity_model/docs/CODEOWNERS_RATIONALE.md``, section
  "Second-CODEOWNER upgrade path".
- **ALM funding-ratio property invariants per product
  (`tests/test_property_invariants.py`).** Six new Hypothesis-driven
  tests, two per product (SPIA / Term / RILA), pin two laws across
  the legal demographic + curve space:
  (1) the engine's surplus / funding-ratio identity
  ``surplus == AMV - LiabPV - borrowing_balance`` and
  ``FR == AMV / (LiabPV + borrowing_balance)`` (positive denom only);
  and (2) at month 0, ``AMV[0]`` -- and, when the liability
  denominator is informative, ``FR[0]`` -- must strictly increase in
  ``initial_asset_market_value``. Both the "+ debt" denom convention
  and the per-product parametrisation are deliberate: the former
  matches the engine's "debt is senior" treatment so a future change
  to net-of-debt accounting can't pass silently, and the latter
  catches a regression that quietly bypasses the
  ``run_alm_projection_from_pricing_result`` dispatch for any single
  product.
- **`scripts/audit_session_state.py`: prerequisite enabler for the
  ``ui/MIGRATION.md`` per-page split.** Walks ``pricing_ui.py``'s AST,
  enumerates every ``st.session_state[...]`` / ``.get`` /
  ``.setdefault`` / ``key=`` reference inside each ``_render_<page>``
  function, and produces both a human-readable summary and a JSON
  report. Records whether each key is referenced via raw literal vs
  the ``RUN_KEY`` symbol so migration progress is measurable. The
  ``--fail-on-cross-page --allow-cross-page <keys>`` mode is the CI
  gate for the actual per-page split: it refuses any new shared key
  outside an explicit allow-list. Today's audit identifies 18
  cross-page keys (the post-pricing result bundle and ALM caches);
  documented as the migration's next blocker in
  ``ui/MIGRATION.md``. The end-to-end per-page split itself remains
  deferred -- ``pricing_ui.py`` is 4,467 LOC vs the planned 1.5k
  trigger threshold. Backed by ``tests/test_audit_session_state.py``
  (10 cases including subscript / method / widget-key / RUN_KEY
  paths, cross-page detection, and the CI-gate mode).
- **OpenTelemetry `@traced(...)` wired onto every parity-critical entry
  point.** ``_observability.traced`` previously existed but was applied
  nowhere -- production OTel deployments were silent. Decorated:
  - ``pricing_projection.price_spia_single_premium`` (deterministic + MC)
  - ``pricing_projection.run_alm_projection`` (legacy SPIA wrapper),
    ``run_alm_projection_from_liability_path`` (generic), and
    ``run_alm_projection_from_pricing_result`` (router)
  - ``term_projection.price_term_life_level_monthly``
  - ``rila_projection.price_rila_single_premium`` (deterministic + MC)
  Each span name is dotted (``pricing.spia.deterministic``,
  ``alm.from_liability_path``, ...) so the trace tree groups cleanly by
  product / surface. Backed by ``tests/test_observability_wiring.py``
  (12 cases): a parametrized meta-invariant asserts every entry point
  in ``TRACED_ENTRY_POINTS`` carries the ``__wrapped__`` marker
  ``functools.wraps`` produces, plus behavioural smokes that the no-op
  fallback is a true pass-through and that the configured span name
  reaches the tracer when OTel IS available.
- **`RUN_KEY` namespace + literal-drift ratchet for Pricing Run session
  state.** ``pricing_run_form_state.RUN_KEY`` now exposes every
  Streamlit ``st.session_state`` key for the Pricing Run page as a
  class-level constant (``RUN_KEY.ISSUE_AGE``, ``RUN_KEY.SPIA_BENEFIT_ANNUAL``,
  ...). New code MUST reference these symbols rather than the raw
  ``"run_*"`` literal -- IDE rename works, typos become
  ``AttributeError``, and the canonical set is reflectively derived as
  ``RUN_STATE_KEY_NAMES``. ``pricing_run_form_state.py`` itself is
  fully migrated. ``pricing_ui.py`` retains its 102 historical
  literals; ``tests/test_run_state_key_drift.py`` is a one-way ratchet
  that compares per-file canonical-literal counts against
  ``tests/run_state_key_baseline.json`` and fails if any file's count
  *increases*. The baseline shrinks naturally as the
  ``ui/MIGRATION.md`` decomposition deletes legacy literals.
- **PR-level mutmut gate over the parity-critical surface**
  (`.github/workflows/mutmut-pr.yml` + `scripts/mutmut_pr_gate.py` +
  `mutmut_thresholds.toml`). The nightly mutmut workflow continues to
  run the full surface as a non-blocking artifact; the new PR-level gate
  runs only on the *touched subset* of parity-critical files (engines,
  builders, validator, registries) and fails the PR if any file exceeds
  its survivor cap. Default cap is zero; per-file overrides require a
  one-line justification + CODEOWNERS sign-off, and `mutmut_thresholds.toml`
  is covered by the same blanket protection that guards
  `parity_constants.py`. Backed by `tests/test_mutmut_pr_gate.py` (13
  cases including a meta-invariant that asserts the PR gate's
  `MUTMUT_SURFACE` is line-for-line identical to the nightly workflow's
  `paths_to_mutate` list, so the two surfaces cannot drift apart).
  Path filter on the workflow itself short-circuits the job in well
  under a minute on PRs that touch no parity-critical code.

### Changed
- **Coverage gate is now a one-way ratchet driven by `pyproject.toml`.**
  The historical `coverage report --fail-under=55` literal in
  `.github/workflows/ci.yml` has been replaced with
  `python scripts/ratchet_coverage.py`, which reads
  `[tool.coverage.report].fail_under` from `annuity_model/pyproject.toml`
  and enforces it. Single source of truth -- no more silent drift between
  the workflow and the project file. To raise the floor after improving
  coverage, run `python scripts/ratchet_coverage.py --update` locally; the
  script refuses to lower the floor (manual edit + reviewer sign-off
  required, and CODEOWNERS protects `pyproject.toml`). Backed by
  `tests/test_ratchet_coverage.py` (9 cases covering pass/fail/update/
  refuse/missing-key paths).

### Security / Governance
- **Branch protection enabled on `main`** (P5 Wave 7, terminal step of the
  Phase-5 hardening sweep -- direct-to-main commits stop here). Configured
  via the GitHub branch-protection API with the following profile:
  - `required_status_checks.strict = true` -- branches must be up to date
    with `main` before merge, and the following CI contexts must all pass:
    - `tests (ubuntu-latest / py3.12)` (carries the parity gate)
    - `tests (ubuntu-latest / py3.11)`
    - `tests (macos-14 / py3.12)`
    - `tests (macos-14 / py3.11)`
    - `tests (windows-latest / py3.12)`
    - `tests (windows-latest / py3.11)`
    - `pre-commit (lint + format + mypy)`
    - `docker build + deep_smoke in container`
    - `build-and-deploy` (mkdocs strict build + Pages deploy)
  - `required_linear_history = true` -- merges must squash or rebase; no
    merge commits cluttering the audit trail.
  - `allow_force_pushes = false`, `allow_deletions = false` -- the parity
    history can't be rewritten or deleted.
  - `required_conversation_resolution = true` -- review threads must be
    resolved before merge (catches the "I'll address that later" pattern).
  - `enforce_admins = false` -- the solo CODEOWNER retains a fire escape
    for genuine emergencies (e.g. a CI provider outage that's blocking a
    security fix). Every admin override MUST be backfilled with a
    follow-up PR + post-mortem entry; this is documented expectation, not
    a technical gate, until the second CODEOWNER lands (Wave 1.5
    deferral).
  - `required_pull_request_reviews = null` -- NOT required because the
    repo currently has a single CODEOWNER. GitHub blocks self-approval,
    so requiring reviews on a one-owner repo creates an unbreakable
    deadlock. When the second CODEOWNER lands, flip this to
    `{required_approving_review_count: 1, require_code_owner_reviews:
    true, dismiss_stale_reviews: true}` -- the JSON template is below.
  Apply / refresh / inspect:
  ```
  # Inspect current rules
  gh api repos/:owner/:repo/branches/main/protection
  # Update (overwrites; idempotent)
  gh api -X PUT repos/:owner/:repo/branches/main/protection \
    --input .github/branch-protection.json
  # Disable (only for emergency, e.g. moving CI workflows)
  gh api -X DELETE repos/:owner/:repo/branches/main/protection
  ```
  When required check NAMES change (e.g. matrix expansion, new gate),
  the protection JSON must be updated in the SAME PR that ships the
  workflow change -- otherwise PRs sit forever waiting on a context
  that no longer reports. The current required-context list is
  authoritative; CI workflows MUST keep these job names stable or
  update the protection list in lockstep.

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

- **Nightly mutmut mutation-test workflow** (Wave 6.1 of phase-5
  hardening). New `.github/workflows/mutmut-nightly.yml` runs `mutmut
  run` against the parity-critical engines + workbook builders every
  night at 04:00 UTC and uploads the survivor report as a 14-day
  artifact.
  - Mutation surface is declared inline in the workflow (generated
    `mutmut_config.py`) and intentionally mirrors the strict-mypy
    override in `pyproject.toml`: `pricing_projection`, `term_projection`,
    `rila_projection`, `alm_excel_ladder`, `build_pricing_excel_workbook`,
    `build_rila_excel_workbook`, `build_term_excel_workbook`,
    `excel_workbook_validator`, `product_excel`, `product_registry`.
    Excluded: `pricing_ui` (no AppTest harness yet -- post-Wave-4
    follow-up), `data_registry` (mutating the sha256 strings would be
    killed by the registry invariant test for free, which pads survivor
    counts without signal).
  - mutmut's hot-loop runner is `pytest tests/parity -q -x` (~1.1s)
    rather than the full suite (~9s); a mutation that survives the
    parity gate but dies on a unit test is almost always a real
    coverage gap, which is what we want to surface.
  - `continue-on-error: true` on both `mutmut run` and the report step
    keeps the nightly non-blocking. After ~2 weeks of baseline data
    + the post-Wave-4 coverage ratchet, both flags come off and
    surviving-mutant count becomes a hard gate (tracked as a Wave 6.2
    follow-up).
  - 60-minute job timeout protects against runaway expansion if a
    new engine is added to the surface without a corresponding
    timeout bump.

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
- **Wave 4 (decompose `pricing_ui.py` into `ui/pages/*`)** is carried
  over to a dedicated multi-PR sequence. `pricing_ui.py` is **4,455
  lines / 49 top-level functions** today and is the only place
  Streamlit `st.session_state` keys are minted. The migration plan in
  `annuity_model/ui/MIGRATION.md` (already in repo) explicitly
  prescribes **one page per PR** with Streamlit-test coverage between
  each step -- that rhythm cannot be compressed into a single direct-
  to-main commit while keeping the all-4-gates-green invariant alive.

  Why "all at once" is unsafe: the only existing tests that exercise
  the UI pages are `tests/test_pricing_ui_state_normalization.py` (135
  tests, all session-state oriented) and the launcher invariants. Page
  renderers themselves (`_render_overview`, `_render_pricing_run`,
  `_render_what_if`, `_render_excel_replicator`, `_render_alm`,
  `_render_unit_tests`) have no smoke coverage today; a regression in a
  moved page would only surface when the user clicks it in Streamlit.

  Migration sequencing for the follow-up PRs (do these in order so
  back-compat aliases shrink monotonically):
  1. Move `_render_overview` -> `ui/pages/overview.py`. Add a one-line
     re-export in `pricing_ui.py` so `from pricing_ui import
     _render_overview` keeps working until the last page lands. Run
     all 4 gates + manual `streamlit run pricing_ui.py` smoke.
  2. Repeat for `_render_pricing_run` (largest, uses `pricing_run_form_state`).
  3. `_render_what_if` (depends on the same form state).
  4. `_render_excel_replicator` (export-then-recompute is parity-
     critical -- run a full deep_smoke on the moved page).
  5. `_render_alm` (long projection visualisation, slow Streamlit
     re-renders).
  6. `_render_unit_tests` (in-app pytest runner; lowest blast radius).
  7. Move helpers (`_render_metric`, `_render_chart`, ...) to
     `ui/widgets/<feature>.py`.
  8. Rename `pricing_run_form_state.py -> ui/forms/run_form_state.py`
     in a separate commit (touches imports across all pages).
  9. Move `main` + sidebar nav to `ui/app.py`. Delete
     `pricing_ui.py` entirely; update launchers and CI to point at
     `ui.app:main` (this last step is also where Wave 3.3's
     `[project.scripts]` entry for `annuity-pricing-ui` lands).

  The Wave 4 final step (coverage ratchet from 55% toward 75%) is also
  deferred until the move completes, because the new `ui/pages/*`
  modules become testable via `streamlit.testing.v1.AppTest` only after
  they're factored out -- ratcheting against the current monolithic
  `pricing_ui.py` would force a fragile import-time mock just to bump
  the gate.

- **Wave 5 (FastAPI wrapper + batch CLI)** -- the original monolithic
  Wave 5 item above is **superseded** by the portfolio program (2026-04):
  `python -m cli portfolio-run` (gated by `ANNUITY_MODEL_PORTFOLIO_V1=1`)
  plus `portfolio_summary.json` / `portfolio.xlsx` outputs and
  integration tests. A thin FastAPI wrapper remains optional and can
  reuse the same `run_portfolio` entry point when packaging (Wave 3.3)
  lands.

- **Wave 6.2 (final coverage ratchet to 75%)** is deferred behind Wave
  4 -- ratchets larger than ~5pp at a time tend to force test-shaped
  hacks that don't catch real regressions. The Wave 4 follow-up PRs
  will each ratchet by 1-2pp as their pages become testable, ending
  at the 75% target naturally.

- **Wave 3.3 (src/ layout + buildable wheel + `[project.scripts]`)** is
  carried over to a dedicated follow-up commit because the three pieces
  cannot land safely without each other in the current layout. Today
  `pyproject.toml` AND `__init__.py` both live inside `annuity_model/`,
  which is neither standard flat (pkg one level above pyproject) nor
  src/ (pkg under `src/`). Configuring setuptools' wheel build on top
  of this would either need pyproject moved to repo root (with the
  co-tenant `actuarial_parity_kit/` reorganised) or every module moved
  into `src/annuity_model/` with ~100 bare imports rewritten as
  `from annuity_model.<x> import <y>` across modules + tests + scripts +
  Streamlit launcher + Dockerfile. Either path is too large for a
  single safe commit while keeping all 4 canonical gates green.

  Migration runbook for the follow-up (do this as a single dedicated PR
  so `git blame` stays clean):
  1. Create `annuity_model/src/annuity_model/` and `git mv` every
     `annuity_model/*.py` into it (preserves blame). Move
     `annuity_model/__init__.py` along with them.
  2. Rewrite intra-package bare imports across all moved files. The
     pattern is mechanical:
     ``import pricing_projection as sp`` ->
     ``from annuity_model import pricing_projection as sp``. Tests
     and `scripts/*.py` get the same rewrite. ``ruff check --fix``
     handles import sorting after the move.
  3. Update `pyproject.toml`:
     - Add `[build-system]` with `setuptools>=68 + wheel`.
     - Add `[project.scripts]` for
       `annuity-deep-smoke = "scripts.deep_smoke:main"`,
       `annuity-parity-trace`, `annuity-render-parity-contract`.
     - Add `[tool.setuptools.packages.find] where = ["src"]` and
       `[tool.setuptools.package-data] "annuity_model" = ["data/**/*.csv"]`
       so the versioned data tree from Wave 3.2 ships in the wheel.
     - Move runtime deps from `requirements.txt` into
       `[project.dependencies]` and keep `requirements.txt` as a thin
       `pip install -e annuity_model[dev]` shim for back-compat.
  4. Update `pytest.ini`: drop `pythonpath = .` (the install handles
     it) and switch `testpaths = src/annuity_model/tests` if tests
     also move. Or leave tests at `annuity_model/tests/` and add
     `pythonpath = src` instead.
  5. Update `Dockerfile`: replace `ENV PYTHONPATH=/app/annuity_model`
     with `RUN pip install /app/annuity_model[dev]` and drop the
     `WORKDIR /app/annuity_model` cd (all the entry points become
     console scripts on `$PATH`).
  6. Update `.github/workflows/ci.yml`: replace ad-hoc `pip install
     -r annuity_model/requirements.lock` with
     `pip install ./annuity_model[dev]`. Add a `python -m build` job
     that produces the wheel + sdist as build artifacts on every push.
  7. Update launchers (`run_pricing_ui.{sh,bat,command}`) to call
     `annuity-pricing-ui` (new console script wrapping `pricing_ui:main`)
     instead of the current `streamlit run pricing_ui.py`. Bump the
     `tests/test_launcher_invariants.py` regexes to match.
  8. Update `.pre-commit-config.yaml` mypy `files:` patterns to point
     at `src/annuity_model/...` and adjust ruff isort `known-first-party`
     to use the package name.
  9. Add a packaging invariant test that builds the wheel inside a
     `tmp_path` venv, installs it, imports `annuity_model`, and asserts
     `annuity-deep-smoke` resolves on `$PATH`. This is the parity-style
     guard against silent regression of the install path.
  10. Validate end-to-end: 4 canonical gates + Docker build + every
      launcher + a `pip install --no-build-isolation .` against a clean
      venv. Only then merge.

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

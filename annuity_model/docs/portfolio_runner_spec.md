# Portfolio runner (v1)

## Purpose

Run **multiple policies** of **mixed `ProductType`** under one shared
`RunScenario`, aggregate each engine’s `LiabilityPath`, and expose:

- Seriatim `PolicyResult` rows.
- `rollups_by_product_type: dict[ProductType, LiabilityPath]` (present types only).
- `liability_path_total` on the **union** monthly grid (`k/12` years, `k = 1 … N`).
- Optional ALM on the total path (`run_portfolio(..., alm_assumptions=…)`).

## Feature flag

Enablement is centralized in **`portfolio_config.portfolio_v1_enabled()`** (same
rules as `run_pricing_ui.sh` / `run_pricing_ui.bat`):

- **`annuity_model/.disable-portfolio-v1`** (gitignored) → **off** (local opt-out).
- Else **`ANNUITY_MODEL_PORTFOLIO_V1`** truthy (`1`, `true`, …) → **on**; falsy
  (`0`, `false`, …) → **off**.
- Else **unset / empty** → **on** by default so `streamlit run pricing_ui.py` still
  shows the Portfolio sidebar without copying shell exports.

The Streamlit sidebar also offers a **session-only** “show Portfolio in Section”
checkbox when the flag is off (`PORTFOLIO_KEY.UI_FORCE_SIDEBAR` in
`pricing_run_form_state.py`). Core library functions (`run_portfolio`, aggregation,
workbook builder) remain importable regardless.

**Streamlit Cloud:** set `ANNUITY_MODEL_PORTFOLIO_V1=0` in secrets only if you need
to hide portfolio there; otherwise the default-on empty env keeps it visible.

## Inforce layout

Canonical example:
`tests/data/inforce/example_v1/inforce.csv`.

Column dispatch is implemented in `inforce_parsers.py` / `inforce_io.py`; each
row must include a `product_type` cell matching `ProductType.value` and
product-specific columns consistent with the underlying `*Contract` dataclass
for that type.

## Default economic package (Pricing Run parity)

Shared economics are materialized from the same ``run_*`` seed keys as the
**Pricing Run** tab via ``pricing_scenario_materialize``:

- **CLI** — ``run_scenario_for_portfolio_policies`` with an empty session dict
  applies ``build_run_form_seed_defaults`` (first-paint defaults) plus
  mortality keys pinned by ``reference_product_type_for_portfolio_scenario`` so
  young issue ages (e.g. Term at 35) stay inside the shared table when mixed
  with annuitant-style products.
- **Streamlit** — the Portfolio section seeds from the same session keys as
  Pricing Run, renders the shared yield / mortality / expense / horizon /
  spread / index controls (Monte Carlo omitted), and calls the same materializer
  at run time.

``portfolio_scenario.default_run_scenario()`` remains a thin wrapper around
``cli_default_run_scenario`` for tests and backward-compatible imports.

## Scalar rollups (`ProductTypeRollupScalars`)

Per type, the runner records:

- `policy_count` (always).
- `sum_single_premium` when **every** policy of that type exposes
  `single_premium` on its pricing result; otherwise `None`.
- `sum_undiscounted_cashflows` when every result exposes
  `expected_total_cashflows`; otherwise `None` (UI/JSON may fall back to the
  rolled-up path sum).

## JSON summary (`portfolio_summary.json`)

Produced by `portfolio_result_to_summary_dict`:

- `n_policies`
- `by_product_type`: keyed by `ProductType.value`, each with `policy_count`,
  `sum_single_premium`, `sum_undiscounted_cashflows`, `rollup_cf_sum`.
- `total_cf_sum` (sum of `liability_path_total.expected_total_cashflows`).

## Excel workbook (`build_portfolio_excel_workbook.py`)

Sheets: `Inputs`, `PolicyRegister`, `PolicyCashflows`, `ProductTypeRollups`,
`LiabilityAggregate`, `ModelCheck`, `README`. Per-policy monthly liability CF
are Python literals on **PolicyCashflows**; **LiabilityAggregate** column
``total_cf`` is a row-wise ``SUM`` over those policy columns. **ModelCheck**
includes the rollup identity ``SUM(by-type cols) - total_cf`` plus a second
reconciliation: Python snapshot of ``total_cf`` vs ``=LiabilityAggregate!C…``
and ``diff_excel_minus_python``. v1 does **not** embed the full SPIA
`Liabilities` + ALM ladder; portfolio ALM is run in Python when requested, not
re-derived in Excel.

## Invariants

On the portfolio union grid, elementwise:

\[
\sum_{\text{type}} \text{rollup\_path[type]} = \text{portfolio\_total}
\]

checked in-code via `assert_rollups_sum_to_total` with `PORTFOLIO_ROLLUP_TOL`
from `parity_constants.py`.

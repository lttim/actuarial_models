# Portfolio runner (v1)

## Purpose

Run **multiple policies** of **mixed `ProductType`** under one shared
`RunScenario`, aggregate each engine’s `LiabilityPath`, and expose:

- Seriatim `PolicyResult` rows.
- `rollups_by_product_type: dict[ProductType, LiabilityPath]` (present types only).
- `liability_path_total` on the **union** monthly grid (`k/12` years, `k = 1 … N`).
- Optional ALM on the total path (`run_portfolio(..., alm_assumptions=…)`).

## Feature flag

- **`ANNUITY_MODEL_PORTFOLIO_V1=1`** enables the Streamlit **Portfolio** sidebar
  section and the `python -m cli portfolio-run` subcommand. Core library functions
  (`run_portfolio`, aggregation, workbook builder) are importable regardless.
- **Local launchers** (`run_pricing_ui.sh`, `.command`, `.bat`) **default this to
  `1`** so double-click runs show the Portfolio section. Opt out with an empty
  **`annuity_model/.disable-portfolio-v1`** file or `ANNUITY_MODEL_PORTFOLIO_V1=0`.
- **Streamlit Cloud** (`streamlit_app.py`) does not set the variable unless you
  configure it in secrets.

## Inforce layout

Canonical example:
`tests/data/inforce/example_v1/inforce.csv`.

Column dispatch is implemented in `inforce_parsers.py` / `inforce_io.py`; each
row must include a `product_type` cell matching `ProductType.value` and
product-specific columns consistent with the underlying `*Contract` dataclass
for that type.

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

Sheets: `Inputs`, `PolicyRegister`, `ProductTypeRollups`, `LiabilityAggregate`,
`ModelCheck`, `README`. Cashflows are **Python literals**; `ModelCheck` column B
is `=SUM(by-type cols) - total` per month. v1 does **not** embed the full SPIA
`Liabilities` + ALM ladder; portfolio ALM is run in Python when requested, not
re-derived in Excel.

## Invariants

On the portfolio union grid, elementwise:

\[
\sum_{\text{type}} \text{rollup\_path[type]} = \text{portfolio\_total}
\]

checked in-code via `assert_rollups_sum_to_total` with `PORTFOLIO_ROLLUP_TOL`
from `parity_constants.py`.

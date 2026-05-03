# Runbook: portfolio (multi-policy) pricing

## When to use

You need **one shared economic scenario** (Pricing Run–equivalent yield, mortality,
horizon, expenses, index path) applied to **many policies** and a **single
aggregated liability path** plus optional ALM on that path.

## CLI

From `annuity_model/` with the project venv active:

```bash
export ANNUITY_MODEL_PORTFOLIO_V1=1
python -m cli portfolio-run \
  --inforce tests/data/inforce/example_v1/inforce.csv \
  --out .smoke/portfolio_run/
```

Outputs:

- `portfolio_summary.json`
- `portfolio.xlsx`

Optional parallelism:

```bash
python -m cli portfolio-run --inforce …/inforce.csv --out …/ --workers 4
```

## Streamlit

The **Portfolio** UI is a **separate sidebar section**, not inside **Pricing Run**.
In the left sidebar, find the **Section** radio, then choose **Portfolio
(multi-policy)** — that page has the inforce CSV upload, manual policy builder,
and **Run portfolio** button.

Enablement matches **`portfolio_config.portfolio_v1_enabled()`** (same rules as
`run_pricing_ui.sh` / `run_pricing_ui.bat`): **on by default** when the env var is
unset, unless **`annuity_model/.disable-portfolio-v1`** exists or
`ANNUITY_MODEL_PORTFOLIO_V1` is set to a falsy value (`0`, `false`, …). That way
`streamlit run pricing_ui.py` keeps the Portfolio section without copying shell
exports. If the section is off, the sidebar expander explains why and offers a
session-only checkbox to show it anyway.

**Streamlit Cloud:** set `ANNUITY_MODEL_PORTFOLIO_V1=0` in secrets only if you
must hide portfolio; otherwise the default-on behavior applies when unset.

## Debugging a parity break

1. Confirm each product row still prices in isolation (single-policy UI / tests).
2. Re-run `tests/parity/portfolio/test_portfolio_aggregation_parity.py` -- it
   asserts Python rollup sum == total.
3. Open `portfolio.xlsx` → `LiabilityAggregate` + `ModelCheck`; if formulas
   error, fix `build_portfolio_excel_workbook.py` and re-run strict validator
   tests.
4. If goldens drift intentionally, refresh **only** with
   `UPDATE_GOLDEN_PORTFOLIO=1` / `UPDATE_GOLDEN_SME=1` and document the change in
   `docs/model_change_log.md` when tolerances move.

## Acceptance recipe

From the repository root (with [`just`](https://github.com/casey/just) installed):

```bash
just portfolio-acceptance
```

This runs **Ring 7** in order: `just preflight`, `pytest tests/parity/portfolio`,
`pytest tests/integration`, `ANNUITY_MODEL_PORTFOLIO_V1=1` deep smoke,
`render_parity_contract.py --check`, CLI `portfolio-run` vs
`tests/data/inforce/example_v1/expected_summary.json`, then **`just
actuary-review-full`** (Gate 5 deterministic evidence). The recipe ends with a
reminder to inspect the emitted `portfolio.xlsx` **ModelCheck** links and
validator output (AGENTS.md).

CI: workflow **`portfolio-acceptance`** (job display name **portfolio
acceptance (ring 7)**) is listed in `.github/branch-protection.json` as a
required status check; apply updated protection with `gh api` per the JSON
file header comment when the contexts list changes.

## Gate 5 verdict (Ring 7 close-out, full scope)

Deterministic evidence: `python annuity_model/scripts/run_actuary_review.py
--scope full --iteration 1` (overwrites
`.cursor/actuary-reviews/_evidence-current.md`). Narrative verdict for this
close-out:

- **Verdict:** APPROVE-WITH-NOTES
- **Rationale:** Automated gates tied to the portfolio surface are green
  (`pytest`, `tests/parity/portfolio`, `tests/integration`, strict workbook
  validation on builder bytes, CLI JSON golden). Portfolio aggregation
  invariants (`PORTFOLIO_ROLLUP_TOL`) and `ModelCheck` **formula** wiring are
  covered in tests; generated workbooks should also be inspected with the
  static validator and ModelCheck formula-link checks per AGENTS.md.

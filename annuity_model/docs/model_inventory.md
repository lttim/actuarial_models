# Model inventory and tiering register

This file defines the model inventory control for `annuity_model`.
It provides ownership, intended use, key limitations, and validation cadence.

## Tiering rubric

- **Tier 1 (high materiality):** valuation/pricing logic that can materially affect business decisions or financial outcomes.
- **Tier 2 (medium materiality):** aggregation/reporting logic with indirect financial impact.
- **Tier 3 (support):** diagnostics and helper tooling with no direct model output authority.

## Inventory

| Model surface | Tier | Owner role | Validator role | Intended use | Validation cadence | Key limitations |
|---|---|---|---|---|---|---|
| SPIA pricing + ALM (`pricing_projection.py`, `build_pricing_excel_workbook.py`) | 1 | Product actuary | Independent actuarial reviewer | SPIA pricing and ALM parity outputs | Per PR + release | Assumption quality depends on selected mortality/yield datasets. |
| Term pricing (`term_projection.py`, `build_term_excel_workbook.py`) | 1 | Product actuary | Independent actuarial reviewer | Term premium and expected claims projection | Per PR + release | Simplifications inherited from selected assumptions. |
| RILA mechanics-production rebuild (`rila_projection.py`, `build_rila_excel_workbook.py`) | 1 | Product actuary | Independent actuarial reviewer | RILA pricing/projection mechanics with segment allocations, buffers, withdrawals, surrender charges, and GLWB state | Per PR + release | MVA, dynamic lapse calibration, illustration compliance, and statutory valuation remain out of scope. |
| Life accumulation products (`wl_projection.py`, `ul_projection.py`, `iul_projection.py`, `vul_projection.py`) | 1 | Product actuary | Independent actuarial reviewer | Life-product premium/account-value projections; IUL includes scheduled premiums, policy access, loan, surrender, and net death-benefit state | Per PR + release | Synthetic CSO placeholder risk if not overlaid with licensed data; no-lapse / overloan riders deferred. |
| Portfolio aggregation (`portfolio_runner.py`, `liability_aggregation.py`, `build_portfolio_excel_workbook.py`) | 2 | Platform actuary | Independent actuarial reviewer | Multi-policy rollups and portfolio ALM | Per PR + release | Shared scenario assumptions may be coarse for heterogeneous cohorts. |
| Scenario materialization (`pricing_scenario_materialize.py`) | 1 | Platform actuary | Independent actuarial reviewer | Shared assumption package for pricing and portfolio runs | Per PR + release | Deterministic defaults unless stochastic scenarios are explicitly configured. |
| Data artifact registry (`data_registry.py`) | 1 | Assumption steward | Independent actuarial reviewer | Controlled lookup for assumption data artifacts | Per PR + release | Governance metadata currently technical-first unless augmented by governance schema. |
| Streamlit UI orchestration (`pricing_ui.py`) | 2 | Platform engineering | Platform reviewer | Interactive model execution and inspection | Per PR | Monolithic composition increases regression coupling risk. |
| CLI orchestration (`cli.py`) | 2 | Platform engineering | Platform reviewer | Batch portfolio pricing and artifact output | Per PR | No centralized run ledger by default. |

## Control requirements

- Every new model surface must be added to this file in the same PR as implementation.
- Tier 1 surfaces require explicit actuarial reviewer signoff.
- Limitation updates must be mirrored in product spec docs and release notes.
- Planned second-reviewer activation should switch branch protection to the profile with required reviews.

## Change protocol

When updating this register:

1. Include a short rationale in [`docs/model_change_log.md`](model_change_log.md) if model-impacting.
2. Keep the table aligned with actual ownership in [`.github/CODEOWNERS`](../../.github/CODEOWNERS).
3. Re-run governance/readiness checks before release.

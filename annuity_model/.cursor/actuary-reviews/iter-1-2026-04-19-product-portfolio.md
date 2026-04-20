---
verdict: APPROVE
scope: product:portfolio
iteration: 1
max_iterations: 5
prior_verdict_path: null
subagent_id: 330a15ac-5f0e-46b7-b3ab-d42004daa58b
evidence_pack: .cursor/actuary-reviews/_evidence-current.md
findings: []
---

## Actuary SME verdict (iteration 1)

### Sign findings
No sign issues identified for this scope. Cached test status reports zero live failures; new ALM equivalence tests assert portfolio-run ALM matches a direct run on the aggregated `LiabilityPath` with nonnegative initial AUM on the canonical fixtures, and the UI guards negative aggregate premium messaging on the waterfall. Liability cashflows plotted are nominal expected outflows as stated in the chart caption.

### Band findings
Evidence-pack benchmark pins (`PORTFOLIO_*`) are present for SME filtering; no band breach is reported in the pack and no objective gate failure is listed. No further band review was required beyond the cached green status for this iteration.

### Sensitivity findings
Not in scope for this portfolio-aggregation iteration; none.

### Closed-form findings
Not applicable to portfolio aggregation / UI wiring; none.

### Cross-product findings
The implementation cleanly separates a **homogeneous** book (single `ProductType`): per-policy product waterfalls are merged with `_merge_profit_waterfall_row_sets`, matching the portfolio-level rows within tolerance in tests. For **mixed** books it falls back to a **generic PV bridge** that sums per-policy `pv_benefit`, `pv_monthly_expenses`, and `single_premium` plus per-policy issue expense, with an explicit caption that intermediate rows are not a single-product story—this is the right actuarial communication for a scalar-sum bridge. No cross-product ordering claims are implied beyond that.

### Assumption / methodology findings
**Portfolio aggregation:** Existing rollup invariants remain supported; `padded_cashflows_on_portfolio_grid` left-aligns type rollups on the portfolio monthly grid, zero-pads trailing months, and refuses truncation—consistent with “no orphan mass” and a shared union grid. Unit tests cover padding and rejection of over-long paths.

**Liability projection chart:** Nominal expected cashflows are shown for aggregate and each `ProductType`, with per-type series padded via the helper so the chart aligns with the aggregate horizon; captions state that the same series feed ALM.

**Baseline ALM on the aggregated path:** `alm_engine_baseline_assumptions()` mirrors Pricing Run ALM defaults, giving a deterministic baseline for batch/CLI. `run_portfolio` ALM is equated to `run_alm_projection_from_liability_path` on `res.liability_path_total` with matching yield curve, spread, assumptions, and initial assets set to summed single premium—`test_run_portfolio_baseline_alm_matches_direct_liability_path` (surplus/funding) and `test_all_products_default_portfolio_runs_with_baseline_alm` (duration gap) lock this down. CLI `--alm` with fallback when aggregate premium is non-positive avoids hard failure while still delivering pricing-only output—reasonable operational behavior.

### Documentation alignment findings
`docs/model_change_log.md` tail in the evidence pack ends at 2026-04-19 portfolio benchmark entries; nothing in this pack proves a gap relative to the project’s “model-impacting” definition for these changes (aggregation presentation, baseline wiring, tests). Evidence also notes missing `spia_product_spec.md` / `test_spia_actuarial.py` paths as workspace pointers—pre-existing coverage gaps for SPIA as a product line, not introduced by this portfolio diff.

### Per-product defaults and outputs (full scope or touched products)

#### portfolio
**Pricing Run defaults:** N/A as a standalone pricing product; portfolio uses uploaded/manual inforce and scenario materialization. **Portfolio defaults / fixtures:** New `all_products_default_v1` holds one row per `ProductType` with plausible ballpark inputs; `test_all_products_default_inforce_loads_ten_distinct_types` confirms coverage. **Outputs:** Baseline ALM on the aggregate path, extended `portfolio_summary` ALM metrics, liability projection multiselect (aggregate + types, cumulative toggle), and waterfall with homogeneous vs mixed logic are actuarially interpretable; UI apptest asserts ALM population for canonical inforce and presence of the projection multiselect.

#### spia
none (SPIA participates only as one row type in mixed/homogeneous portfolio tests; no new SPIA engine assertions in this diff).

### Prior-finding regression check
Omitted—iteration 1.

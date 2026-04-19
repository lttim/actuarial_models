# Model change log

A change is "model-impacting" iff it modifies any of:

* `parity_constants.py` (any tolerance or epsilon value).
* Mortality table loader, q_x source, or generational projection.
* Yield-curve construction, interpolation, or extrapolation policy.
* ALM disinvestment / reinvestment rule, borrowing policy, or rebalance
  policy.
* RILA crediting (cap, floor, participation, fee, segment window).
* SPIA payment timing / cessation policy.
* Term Life premium / claim cashflow definition.

Every entry MUST include:

| Field             | Required content                                                |
|-------------------|------------------------------------------------------------------|
| Date              | ISO `YYYY-MM-DD`                                                 |
| Version / PR      | Semver tag if released, plus PR link                             |
| Author / reviewer | The CODEOWNER who approved                                       |
| Summary           | One sentence: what changed                                       |
| Justification     | Why -- regulatory, bug fix, methodology improvement              |
| Parity trace      | Path to before/after CSVs from `scripts/parity_trace.py`         |
| Backward compat   | Whether existing scenarios reproduce within the prior tolerance  |

Entries are appended at the **bottom**; never edit historical entries.

---

## 2026-04-18 -- Tolerance constants centralised (no numeric change)

- **PR:** hardening-roadmap (this branch)
- **Author / reviewer:** lttim
- **Summary:** Created `parity_constants.py` as the single source of truth for
  tolerance / epsilon values. Refactored `tests/parity/test_alm_parity.py`,
  `test_rila_parity.py`, `test_term_parity.py`, and `excel_formula_sim.py`
  to import from the new module. No numeric values changed.
- **Justification:** Eliminates the doc-vs-test drift class of bug. Makes
  CI's "tolerance was secretly weakened" check possible
  (`scripts/render_parity_contract.py --check`).
- **Parity trace:** N/A (refactor only). Verified by full
  `pytest tests/parity` green and `scripts/render_parity_contract.py --check`
  green.
- **Backward compat:** Yes -- bit-for-bit identical numerics to prior tag.

## 2026-04-18 -- RILA PV tolerance tightened from 1e-3 to 1e-4

- **PR:** hardening-roadmap (P0 doc-drift fix)
- **Author / reviewer:** lttim
- **Summary:** Brought `tests/parity/test_rila_parity.py` PV assertions in
  line with `docs/rila_parity_contract.md` (`1e-4`).
- **Justification:** The contract has stated `1e-4` since v1.0; the test was
  silently looser. No production scenario observed at the looser tolerance,
  so the tightening is a documentation-conformance fix, not a methodology
  change.
- **Parity trace:** Existing parity tests pass at the tighter tolerance.
- **Backward compat:** Yes.

## 2026-04-19 -- Phase 0: Foundation for seven-product rollout

- **PR:** seven-product-rollout (Phase 0)
- **Author / reviewer:** lttim
- **Summary:** Landed the Phase 0 deliverables for the seven-product
  rollout per `docs/seven_product_rollout_plan.md`:
  * **`lapse.py`** — static lapse / persistency framework
    (`LapseAssumption`, `combined_monthly_survival`, monthly hazard
    helpers). Opt-in only; existing engines unchanged.
  * **`crediting.py`** — strategy hierarchy
    (`CreditingStrategy`, `FixedDeclaredRate`,
    `AnnualPointToPointCapped`). RILA's `segment_credited_return`
    refactored to delegate to `AnnualPointToPointCapped` — public name
    preserved, RILA golden JSON byte-identical.
  * **`account_value.py`** — single-source UL/IUL/VUL monthly AV
    cycle (`AVConfig`, `evolve_account_value`). Used by Phases 5/6/7.
  * **`mortality_2017_cso.py` + four CSV artifacts** under
    `data/mortality/cso_2017_ult/` — synthetic placeholder Gompertz-
    Makeham approximation of CSO 2017 Ultimate (sex × smoker). NOT
    licensed CSO; production users overlay their own file at the same
    path. README in the data dir documents the placeholder status.
  * **`actuarial_benchmarks.py` + `docs/actuarial_benchmarks.md`** —
    per-product band constants + rationale narrative; cross-checked
    by `scripts/render_actuarial_benchmarks.py --check` (now part of
    `just preflight`).
  * **`parity_constants.py`** extended with
    `LIFE_MODELCHECK_TOL`, `ANNUITY_ACCUM_MODELCHECK_TOL`, `AV_TOL`,
    `LAPSE_DECREMENT_TOL`, and per-product `*_PV_TOL` constants for
    MYGA / FIA / VA / WL / UL / IUL / VUL.
  * **`ProductType` enum** extended with `MYGA`, `FIA`,
    `UNIVERSAL_LIFE`, `INDEXED_UL`, `VARIABLE_UL` (added; existing
    `WHOLE_LIFE` and `VARIABLE_ANNUITY` repurposed from "scaffold-only"
    to fully implemented in later phases).
  * **`liability_layouts.py`** — added 7 new layout entries up front
    (Phase 0). Accumulation products (RILA / MYGA / FIA / VA) use
    `total_cf_col=M, discount_col=O`; life products (SPIA / Term / WL /
    UL / IUL / VUL) use `total_cf_col=S, discount_col=O`.
  * **`product_registry`** mortality-mode wiring — life products
    default to `cso_2017_ult`; annuity products keep `rp2014_mp2016`.
- **Justification:** Foundation for seven new products (Phases 1–7).
  Implementing the abstractions up front avoids re-deriving the same
  math seven times and keeps each per-product engine and Excel builder
  small.
- **Parity trace:** RILA golden JSON byte-identical after the
  back-compat refactor (`tests/parity/test_golden_modelcheck.py`
  green). All existing tests green: `pytest tests/parity -q` (32
  passed) and `pytest -q --ignore=tests/parity` (534 passed).
- **Backward compat:** Yes — existing SPIA / Term / RILA outputs are
  byte-identical. New constants are additive; new modules are
  standalone.

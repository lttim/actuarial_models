# Seven-Product Rollout Plan — VA, MYGA, FIA, WL, UL, IUL, VUL

**Status:** DRAFT (planning).
**Scope:** Add seven insurance products to the platform under a single mega-PR
that stays in a green state at every internal phase boundary.
**Decisions locked in by sponsor (2026-04-18 Q&A):**

| Decision | Choice |
|----------|--------|
| Per-product depth | **Minimal v1** — single-premium, basic features, no riders. Mirrors how RILA shipped. |
| Premium structure (life products) | **Single premium only** in v1 across WL / UL / IUL / VUL. |
| Lapse modeling | **Static lapse table** (annual rates by policy year). Add a small generic lapse framework all products can opt into. |
| Mortality tables (life) | **Add 2017 CSO Ultimate** (sex × smoker-distinct) as a new mortality source under `data_registry/`. |
| Phasing / merge cadence | **Single mega-PR** with internal phasing (this document). Each internal phase leaves the four canonical gates green. |

---

## 0. Guiding principles (non-negotiable)

These flow from `AGENTS.md`, `docs/AI_AGENT_PREFLIGHT.md`,
`.cursor/rules/actuarial-parity.mdc`, and `.cursor/rules/excel-formula-safety.mdc`.

1. **Python is the source of truth, Excel is the auditor.** `MODELCHECK_TOL = 0.0`
   stays exact for every implemented product, including the seven new ones.
2. **Every internal phase ends green on the four canonical gates** (parity,
   full pytest, `deep_smoke.py`, `render_parity_contract.py --check`).
3. **Existing SPIA / Term / RILA parity is byte-perfect at every commit.** Any
   diff in their golden JSON or recalc cells is a stop-the-line failure.
4. **`validate_workbook_or_raise(wb)` is called immediately before every
   `wb.save(...)`** for every new builder.
5. **Per-product Excel column letters live in `LIABILITY_LAYOUTS`** —
   never hard-coded inline.
6. **No new ad-hoc tolerances in code.** New constants land in
   `parity_constants.py` AND `docs/model_change_log.md` in the same commit.
7. **Meta-invariant tests are the contract.** Any failure is fixed by wiring
   the missing entry, never by silencing the test.
8. **No copyrighted table data committed.** 2017 CSO base rates are loaded
   from a placeholder synthetic table (mirroring how RP-2014 is handled today)
   plus a documented lookup contract; production users provide their own
   licensed CSO file.
9. **Single-premium contract API everywhere in v1.** Premium-stream abstraction
   and rider menu are deferred to v2 (out of scope).
10. **UI per-product branches stay registry-driven** wherever feasible.
    Hardcoded `if product == X` blocks in `pricing_ui.py` are limited to
    contract widgets and contract construction (matches today's pattern).

---

## 1. Architectural extensions (Phase 0 — Foundation)

The current platform has three abstractions worth extending before adding any
product, so we don't pay the same boilerplate seven times:

### 1.1 Lapse / persistency framework — `lapse.py` (new)

A small standalone module:

```python
@dataclass(frozen=True)
class LapseAssumption:
    annual_lapse_rates_by_year: tuple[float, ...]  # q_w by policy year (year 1, 2, ...)
    ultimate_rate: float = 0.0                     # used after the table runs out

    def monthly_decrements(self, n_months: int) -> np.ndarray:
        """Return monthly lapse rates of length n_months (1 - (1-q_w_y)**(1/12) per month)."""

def combined_monthly_survival(
    *,
    mortality_monthly_q: np.ndarray,
    lapse_monthly_q: np.ndarray,
) -> np.ndarray:
    """S(t) = ∏ (1 - q_x_m(s)) * (1 - q_w_m(s)) for s = 0..t-1."""

def lapse_decrement_from_csv(path: str) -> LapseAssumption: ...
def default_lapse_assumption() -> LapseAssumption:
    """Industry-pattern declining-then-ultimate (e.g. 8/7/6/5/4/3/2 ultimate 2%)."""
```

* **Per-product opt-in:** every new product engine accepts an optional
  `lapse: LapseAssumption | None = None`. If `None`, no lapse decrement is
  applied (back-compat with existing engines).
* **No retrofit to SPIA / Term / RILA in this PR** (would invalidate golden
  values). Their behavior stays mortality-only; the optional `lapse=` slot is
  added for future use but defaults to `None` — verified by golden JSON.

### 1.2 Crediting-strategy framework — `crediting.py` (new)

Extracts the inline RILA logic into a strategy hierarchy:

```python
class CreditingStrategy(Protocol):
    def credit_segment(self, *, raw_index_return: float) -> float: ...

@dataclass(frozen=True)
class FixedDeclaredRate(CreditingStrategy):
    annual_rate: float
    def credit_segment(self, *, raw_index_return: float) -> float:  # ignores index
        return self.annual_rate

@dataclass(frozen=True)
class AnnualPointToPointCapped(CreditingStrategy):
    participation: float
    cap: float
    floor: float
    def credit_segment(self, *, raw_index_return: float) -> float:
        x = self.participation * raw_index_return
        return max(self.floor, min(self.cap, x))
```

* **RILA is migrated** to use `AnnualPointToPointCapped` internally; the inline
  `segment_credited_return` keeps its public name as a thin wrapper for
  back-compat (no behavior change → byte-identical golden JSON).
* **MYGA / FIA / IUL** use these strategies directly.

### 1.3 Account-value engine — `account_value.py` (new)

Single source for the UL/IUL/VUL monthly AV equation. Generic enough that
"variable" (sub-account return) and "indexed" (segment crediting) are just
different `monthly_credit` callables:

```python
@dataclass(frozen=True)
class AVConfig:
    initial_premium: float
    premium_load_pct: float                # % deducted at issue (one-time)
    monthly_expense_charge: float          # flat $ per month
    db_type: Literal["return_of_av", "level_face"]  # GMDB / UL DB
    face_amount: float

def evolve_account_value(
    *,
    config: AVConfig,
    n_months: int,
    monthly_credit_rate: np.ndarray,       # length n_months, supplied by caller
    monthly_coi_q: np.ndarray,             # qx for the COI calc, length n_months
) -> np.ndarray:
    """AV[t+1] = max(0, (AV[t] + premium_load_credit) * (1 + cred_rate) - COI - exp_charge)
       where COI = qx * NAR (NAR = max(0, face - AV))."""
```

* **UL** uses a flat declared rate as `monthly_credit_rate`.
* **IUL** uses a strategy-driven rate (zero except on segment anniversaries).
* **VUL** uses sub-account returns.
* **WL / MYGA / FIA / VA** do not use this module (no explicit COI machinery).

### 1.4 Life mortality — extend `pricing_projection` with `MortalityTable2017CSO`

* New class `MortalityTable2017CSO` modeled on `MortalityTableQx`.
* **Sex** (`male` / `female`) and **smoker class** (`nonsmoker` / `smoker`) are
  constructor parameters; the class loads the matching `q_x` row from a CSV.
* **Data:** four new artifacts in `data_registry.REGISTRY`:
  * `cso_2017_ult_male_nonsmoker_qx`
  * `cso_2017_ult_female_nonsmoker_qx`
  * `cso_2017_ult_male_smoker_qx`
  * `cso_2017_ult_female_smoker_qx`
  * Each is a placeholder (synthetic blended rates with documented warning,
    same convention as RP-2014). Production users overlay their own files.
* **No changes to existing mortality tables.** Annuity products keep using
  RP-2014/MP-2016 or SSA tables.
* **Mortality-mode wiring in `product_registry.py`** in the same Phase 0 commit:
  * Add `"cso_2017_ult"` to `_MORTALITY_MODE_LABELS` →
    `"2017 CSO Ultimate (sex × smoker)"`.
  * Set `_PRODUCT_DEFAULT_MORTALITY_MODE` for the four life products
    (WHOLE_LIFE, UNIVERSAL_LIFE, INDEXED_UL, VARIABLE_UL) to
    `"cso_2017_ult"`.
  * Set `_PRODUCT_MORTALITY_MODE_OPTIONS` for those same four to
    `("cso_2017_ult", "qx_csv", "synthetic")`.
  * VA stays on annuitant table (`"rp2014_mp2016"`); MYGA / FIA same.

### 1.5 ProductType enum extension

Add five new members to `ProductType` (WHOLE_LIFE and VARIABLE_ANNUITY already
exist as scaffolded-but-unimplemented):

```python
class ProductType(str, Enum):
    SPIA = "spia"
    TERM_LIFE = "term_life"
    RILA = "rila"
    WHOLE_LIFE = "whole_life"          # currently scaffold; implemented in this PR
    VARIABLE_ANNUITY = "variable_annuity"  # currently scaffold; implemented in this PR
    MYGA = "myga"                       # NEW
    FIA = "fia"                         # NEW
    UNIVERSAL_LIFE = "universal_life"   # NEW
    INDEXED_UL = "indexed_ul"           # NEW (IUL)
    VARIABLE_UL = "variable_ul"         # NEW (VUL)
```

### 1.6 Optional shared Excel builder — `excel_builder_helpers.LifeProductBuilderTemplate`

A lightweight declarative template that emits the shared sheets common to all
new products (`Inputs`, `YieldCurve`, `MonthlyCurve`, `MortalMonthly`,
`Liabilities`, `ModelCheck`) and lets each builder add product-specific extras
(`SegmentCredits` for IUL, `SubAccountPath` for VUL, etc.). This keeps each
new `build_<product>_excel_workbook.py` to ~150 LOC instead of ~500.

### 1.7 New tolerance constants in `parity_constants.py`

| Constant | Initial value | Purpose |
|----------|---------------|---------|
| `LIFE_MODELCHECK_TOL` | `0.0` | ModelCheck cells for WL / UL / IUL / VUL |
| `ANNUITY_ACCUM_MODELCHECK_TOL` | `0.0` | ModelCheck cells for MYGA / FIA / VA |
| `AV_TOL` | `1e-6` | Account-value reconciliation per-month |
| `LAPSE_DECREMENT_TOL` | `1e-12` | Combined survival sanity (multiplicative) |
| `IUL_PV_TOL` | `1e-4` | IUL PV / single-premium implicit equation tolerance |
| `VUL_PV_TOL` | `1e-4` | VUL PV / single-premium implicit equation tolerance |
| `VA_PV_TOL` | `1e-4` | VA GMDB PV |
| `MYGA_PV_TOL` | `1e-4` | MYGA PV (deterministic, but kept for symmetry) |
| `FIA_PV_TOL` | `1e-4` | FIA PV |
| `WL_PV_TOL` | `1e-4` | Whole life PV |
| `UL_PV_TOL` | `1e-4` | UL PV |

Every change goes through `model_change_log.md` in the same commit, then
`scripts/render_parity_contract.py --check` is rerun. The renderer reads
`parity_constants.__all__`, so each new constant must be exported in the
`__all__` list.

### 1.8 Mortality unions, contract unions, builder spec unions

* `product_registry.ProductContract` is widened to include the seven new
  contract dataclasses.
* The mypy-strict glob `products.*.engine` already covers any new
  `products/<name>/engine.py` shim. No `pyproject.toml` edits needed for the
  per-product engines beyond verifying via `tests/test_mypy_strict_glob.py`.
* Each new builder spec class is added to the dispatcher's
  `_BUILDER_SPEC_TYPES` map via `@register_builder`.

### 1.9 Phase 0 deliverables — exit checklist

* [ ] `lapse.py`, `crediting.py`, `account_value.py` land with full unit tests
      under `tests/test_lapse.py`, `tests/test_crediting.py`,
      `tests/test_account_value.py`.
* [ ] `MortalityTable2017CSO` lands with synthetic CSV data + four
      `data_registry.REGISTRY` entries; `tests/test_data_registry_invariants.py`
      green; new `tests/test_mortality_2017_cso.py` covers loader + lookup.
* [ ] `parity_constants.py` extended; `__all__` updated;
      `docs/model_change_log.md` updated; `scripts/render_parity_contract.py`
      rerun; `--check` green.
* [ ] `ProductType` extended (5 new members); existing tests still green
      because the meta-invariants only require *implemented* products to be
      wired (the new members will be unimplemented at end of Phase 0).
* [ ] `_MORTALITY_MODE_LABELS` / `_PRODUCT_DEFAULT_MORTALITY_MODE` /
      `_PRODUCT_MORTALITY_MODE_OPTIONS` extended for the four life-product
      enum members (with `cso_2017_ult` as the default for life,
      `rp2014_mp2016` for VA, default annuity mode for MYGA / FIA).
* [ ] RILA migrated to use `AnnualPointToPointCapped` internally with
      back-compat wrapper. Verified by running RILA golden JSON unchanged.
* [ ] **`tests/test_observability_wiring.py::TRACED_ENTRY_POINTS`** kept
      in sync — Phase 0 itself adds no entries; the per-product phases
      will (see Section 2 Step H of the per-product template).
* [ ] **Actuarial benchmarks framework** lands as a pair:
      `actuarial_benchmarks.py` (Python constants — empty/skeleton at
      Phase 0; populated row-by-row in Phases 1–7) plus the narrative
      `docs/actuarial_benchmarks.md` and the cross-check script
      `scripts/render_actuarial_benchmarks.py` (with `--check` mode
      mirroring `render_parity_contract.py`). The script is wired into
      `just preflight`. Each per-product Step P imports band constants
      from `actuarial_benchmarks.py`; no inline literals in tests.
* [ ] **Initial handoff** `.cursor/handoffs/<timestamp>-phase-0-foundation.md`
      created at the end of Phase 0 (per Section 12.5). This is the
      starting point for Phase 1.
* [ ] Four canonical gates green at end of Phase 0.

---

## 2. Per-product execution template

For each product **P** in (MYGA, FIA, VA, WL, UL, IUL, VUL), in this order:

> Note: Order is **simplest → most complex**, so each product reuses
> abstractions hardened in earlier phases.

| Step | Artifact | Owning file(s) |
|------|----------|----------------|
| **A** | Engine module | `<P>_projection.py` (Contract dataclass, Result dataclass, `price_<P>_*`, `liability_path_from_<P>_projection`, `register_liability_path_converter` at module bottom) |
| **B** | Excel builder | `build_<P>_excel_workbook.py` (`<P>ExcelBuildSpec`, `<P>_excel_spec_from_launcher`, `build_<P>_workbook_from_spec`, `@register_builder`, calls `validate_workbook_or_raise`) |
| **C** | Liability layout | one entry in `LIABILITY_LAYOUTS` (`liability_layouts.py`) |
| **D** | Adapter & registry | `<P>ProductAdapter` in `product_registry.py`; `_PRODUCT_ADAPTERS`, `_PRODUCT_CAPABILITIES`, `_PRODUCT_MORTALITY_MODE_OPTIONS`, `_PRODUCT_DEFAULT_MORTALITY_MODE`, `_PRODUCT_DISPLAY_NAME`, `_PRICING_METRIC_FORMATTERS`, `_PRODUCT_UI_CONFIG` extended (the placeholder "(coming soon)" / "scaffolded but not implemented yet" entries are replaced with real labels) |
| **E** | Subpackage shim | `products/<P>/{__init__,schema,engine,excel,ui}.py` — generated by `scripts/scaffold_product.py --code <P> --display-name "..." --contract-class <P>Contract --result-class <P>ProjectionResult`, then implementation backfilled |
| **F** | Streamlit UI | New `elif selected_product == ProductType.<P>:` branches in `_render_run_and_results` (contract widgets ~lines 2218–2309 and contract construction ~lines 2573–2606); `pricing_run_form_state.RUN_KEY` extended; new keys appended to `PRICING_RUN_NUMBER_INPUT_KEYS` (numeric inputs only); `build_run_form_seed_defaults` extended; `_normalize_run_state_for_selected_product` extended (e.g. force `run_use_index = True` for FIA / IUL / VUL; force `run_use_index = False` for MYGA / WL / UL); `_clear_dependent_state_on_pricing_change` extended |
| **G** | Per-product validator | `_PRODUCT_VALIDATORS[<P>]` entry — required when product has cross-input rules (e.g. FIA / IUL: `cap >= floor`; UL / IUL / VUL: `face_amount > 0`, `single_premium > 0`, `premium_load_pct in [0, 1)`; VA: `single_premium > 0`, `me_charge_annual in [0, 0.05]`; MYGA: `declared_rate_annual in [-0.5, 1.0]`, `guarantee_years in [1, 30]`) |
| **H** | Observability | `@traced("pricing.<P>.deterministic")` (and `.monte_carlo` if applicable) on the new entry-point functions; **`tests/test_observability_wiring.py::TRACED_ENTRY_POINTS`** extended with the new tuple(s); test must stay green |
| **I** | Parity test | `tests/parity/test_<P>_parity.py` — copy `test_term_parity.py` template; assert ModelCheck reconciles within `LIFE_MODELCHECK_TOL` / `ANNUITY_ACCUM_MODELCHECK_TOL`; assert engine-shape invariants (e.g. UL: AV ≥ 0; FIA: AV monotone non-decreasing when floor ≥ 0); assert workbook structure formula needles |
| **J** | Golden JSON | `tests/parity/golden/<P>.json` — generate once with `UPDATE_GOLDEN_MODELCHECK=1`, then byte-exact henceforth |
| **K** | Regression matrix | `_build_<P>_fixture` in `tests/test_regression_matrix.py`; entry in `_FIXTURE_BUILDERS`; surface coverage (deterministic + (optional) Monte Carlo + liability path + ALM + Excel build/validate + metric formatter + validate_run_inputs); add `EXPECTED_SKIPS` only with documented reason |
| **L** | Excel recalc case | `_CASE_BUILDERS[<P>]` in `tests/parity/test_excel_recalc_per_product.py` (the gate `test_every_implemented_product_has_a_recalc_case` enforces presence) |
| **M** | UI AppTest | `tests/ui/test_apptest_<P>.py` smoke (renders Pricing Run page, can run pricing, no exceptions); `tests/ui/test_apptest_full_workflow.py` automatically covers the new product because it walks `implemented_product_types()` |
| **N** | Deep smoke | `build_<P>()` function and tuple entry in `scripts/deep_smoke.py` |
| **O** | Property tests | New laws in `tests/test_property_invariants.py` (e.g. UL: AV ≥ 0; IUL: monthly credit ∈ [floor, cap] per segment; VA: GMDB PV ≥ 0; lapse: combined survival monotone non-increasing) |
| **P** | **Actuarial assessment** | `tests/parity/test_<P>_actuarial.py` — sanity assertions (Section 13.2), benchmark band assertions (Section 13.3 row), sensitivity matrix (Section 13.4 rows applicable), closed-form cross-validation (Section 13.5 row). **Failure means investigate-then-fix the engine, NEVER widen the band.** This is the gate that catches "engine is internally consistent but actuarially nonsense" — the failure mode pure parity tests cannot see. |
| **Q** | Documentation | `docs/<P>_product_spec.md` (one-pager: assumptions, default inputs, illustrative output); `docs/<P>_parity_contract.md` ONLY if new tolerances introduced; `docs/model_change_log.md` entry; `docs/glossary.md` updated; `docs/CHANGELOG.md` entry |
| **R** | Gate run | All **four canonical gates** + `tests/test_excel_export_validation.py` + `tests/parity/` (which now includes the actuarial test from Step P) + `tests/test_meta_invariants.py` + `tests/test_observability_wiring.py` + `tests/test_run_state_key_drift.py` (key ratchet) green; mutmut PR gate green on touched files |
| **S** | **Phase handoff** | User runs `!handoff phase-N-<product>` per Section 12.5 to persist phase-exit state into `.cursor/handoffs/`. The next phase starts fresh with `!recall <prior-phase-slug>` (Section 12.2). |

The scaffold script auto-generates step E. Steps A and B are the only ones
that need real actuarial implementation; the rest is mechanical wiring guided
by the meta-invariant tests' failure messages. **Step P (actuarial assessment)
is the human-judgment gate**: it cannot be auto-generated and must be reviewed
by someone who knows the product. **Step S (phase handoff)** is the
context-hygiene checkpoint that lets the next phase start in a clean
environment.

---

## 3. Phases 1–7 — per-product implementation specifics

### Phase 1 — MYGA (Multi-Year Guaranteed Annuity)

**Concept.** Single-premium fixed deferred annuity. Issuer guarantees a
declared annual rate for `guarantee_years` (typically 3, 5, 7). At end of
guarantee period, contract is rolled or surrendered. Liability is the
accumulated value at maturity, weighted by survival.

**Default illustrative inputs:**

| Input | Default | Rationale |
|-------|---------|-----------|
| `issue_age` | 60 | Typical MYGA buyer is pre-retiree |
| `sex` | `"male"` | Matches existing default; both supported |
| `single_premium` | `100_000.0` | Standard illustration size |
| `declared_rate_annual` | `0.045` | Plausible 5y MYGA rate in current rate environment |
| `guarantee_years` | `5` | Most common MYGA term |
| `surrender_charge_schedule_pct` | `(7, 6, 5, 4, 3, 0, 0, 0, 0, 0)` | Standard 5y MYGA schedule (year 1 → 7%, year 6+ → 0%); used for v1 to scope a future rider, NOT priced into single premium |
| `mortality_mode` (UI default) | `"rp2014_mp2016"` | Same as SPIA (annuitant table) |
| `lapse` | `None` for v1 default; configurable | Static lapse table optional |

**Contract dataclass:**

```python
@dataclass(frozen=True)
class MYGAContract:
    issue_age: int
    sex: Literal["male", "female"]
    single_premium: float
    declared_rate_annual: float
    guarantee_years: int = 5
    payment_freq_per_year: int = 12
```

**Engine.**

* Monthly accumulation: `AV[t] = single_premium * (1 + declared_rate_annual)^(t/12)`.
* `n_months = guarantee_years * 12`.
* **Cashflow shape (matters for ALM and ModelCheck):**
  * **Maturity payout** at `t = T`: `AV[T] * survival[T-1]` (alive-weighted).
  * **In-period death payout** at every month `t in [1, T)`:
    `AV[t] * P(death in month t)` where
    `P(death in month t) = survival[t-1] - survival[t]`.
  * **Lapse payout** (when `lapse=` is supplied):
    `AV[t] * P(lapse in month t)` (no surrender charge in v1's pricing
    even if a schedule is recorded for display).
  * The sum of the three weighted cashflows is the engine's
    `expected_total_cashflows[t]`; the liability path uses that vector.
* Single premium per the contract IS the input — engine returns the
  pre-computed accumulation path and the cashflow vector, and computes
  PV(total CF), `economic_reserve`, and `single_premium` echo for ModelCheck.
* Capabilities: `supports_economic_scenario=False`,
  `supports_monte_carlo=False`.

**Why first.** Validates the new patterns with the simplest possible math:
no indexing, no COI, no AV equation, no sub-account.

### Phase 2 — FIA (Fixed Indexed Annuity)

**Concept.** Single-premium deferred annuity with crediting linked to an
index. Floor=0 (typical FIA, no negative crediting), cap+participation
applied to annual P2P. Maturity payout = AV at horizon, weighted by survival.

**Default illustrative inputs:**

| Input | Default | Rationale |
|-------|---------|-----------|
| `issue_age` | 60 | Typical FIA buyer |
| `sex` | `"male"` | Consistent default |
| `single_premium` | `100_000.0` | Standard illustration |
| `participation` | `0.80` | Plausible mid-tier FIA |
| `cap` | `0.07` | 7% cap, plausible |
| `floor` | `0.0` | FIA hallmark (no negative crediting) |
| `rider_fee_annual` | `0.0` | No rider in v1 |
| `segment_months` | `12` | Annual P2P |
| `horizon_years` | `10` | Typical FIA accumulation horizon |
| `index_scenario_csv_path` | default S&P 500 baseline scenario | Deterministic; same as RILA |

**Engine.**

* Reuse `crediting.AnnualPointToPointCapped(participation, cap, floor)`.
* Walk monthly: `AV[t]` accumulates by segment-credited return on
  anniversaries; otherwise unchanged.
* **Cashflow shape:** same three-bucket pattern as MYGA (maturity at `T`,
  in-period death at each month, optional lapse) — all weighted by the
  appropriate decrement and paid as `AV[t]`.
* Capabilities: `supports_economic_scenario=True`, `supports_monte_carlo=True`
  (reuse RILA's GBM index simulator).

**Reuses:** RILA's index scenario CSV machinery, GBM simulator, crediting
strategy module.

### Phase 3 — VA (Variable Annuity)

**Concept.** Single-premium deferred VA with a sub-account modeled as GBM.
M&E charge taken monthly from AV. Basic GMDB = max(AV, premium) at death.
Maturity payout = AV at horizon weighted by survival; in-period death
benefits = GMDB.

**Default illustrative inputs:**

| Input | Default | Rationale |
|-------|---------|-----------|
| `issue_age` | 55 | Typical VA buyer |
| `sex` | `"male"` | |
| `single_premium` | `100_000.0` | |
| `me_charge_annual` | `0.014` | Industry-typical 140 bps M&E |
| `gmdb_basis` | `"return_of_premium"` | Simplest GMDB |
| `subaccount_drift_annual` | `0.06` | 6% expected return |
| `subaccount_vol_annual` | `0.15` | 15% vol |
| `horizon_years` | `20` | |
| `index_scenario_csv_path` | reuse S&P 500 baseline scenario | Default deterministic path |

**Engine.**

* AV walks monthly: `AV[t+1] = AV[t] * (S_idx[t+1]/S_idx[t]) * (1 - me_monthly)`,
  where `S_idx` is the sub-account level path (deterministic CSV by
  default; GBM-simulated under Monte Carlo).
* Death-benefit cashflow per month: `max(AV[t], premium) * P(death in month t)`
  (the GMDB return-of-premium guarantee).
* Maturity cashflow at horizon: `AV[n_months] * survival[n_months-1]`.
* Optional lapse decrement applied via the lapse framework; lapse payout
  per month: `AV[t] * P(lapse in month t)` (no surrender charge in v1).
* Capabilities: economic scenario + Monte Carlo both `True` (reuse GBM
  simulator).
* **Sub-account source:** by default, `index_scenario_csv_path` re-uses
  the existing S&P 500 baseline scenario; the user can override with a
  custom CSV; Monte Carlo uses `simulate_index_levels_gbm` with the
  user-supplied drift / vol / seed.

### Phase 4 — Whole Life (WL)

**Concept.** Single-premium paid-up whole life. Level death benefit
`face_amount` payable at death-month-end for life. Mortality is 2017 CSO
Ultimate (sex × smoker).

**Default illustrative inputs:**

| Input | Default | Rationale |
|-------|---------|-----------|
| `issue_age` | 45 | Typical permanent-life buyer |
| `sex` | `"male"` | |
| `smoker_class` | `"nonsmoker"` | Default |
| `face_amount` | `250_000.0` | Standard illustration |
| `single_premium` | (computed) PV of benefits + expenses | Solved given face |
| `mortality_mode` | `"cso_2017_ult"` | New default for life products |
| `horizon_age` | `120` | End of mortality table |

**Engine.**

* Liability cashflow: `expected_benefit[t] = face_amount * P(death in month t)`,
  where `P(death in month t) = survival[t-1] - survival[t]`.
* Single premium = PV of benefits + PV of monthly expenses (industry-typical
  policy fee of $0 default).
* Capabilities: `supports_economic_scenario=False`, `supports_monte_carlo=False`.

**Why this slot.** Introduces 2017 CSO mortality and the
"face × death-prob" cashflow shape used by the next three products too.

### Phase 5 — Universal Life (UL)

**Concept.** Single-premium UL with explicit COI. Monthly: load → declared
rate credit → COI → flat expense charge. DB = max(face, AV) (Type A).
Liability = DB × P(death in month).

**Default illustrative inputs:**

| Input | Default | Rationale |
|-------|---------|-----------|
| `issue_age` | 45 | |
| `sex` | `"male"` | |
| `smoker_class` | `"nonsmoker"` | |
| `face_amount` | `250_000.0` | |
| `single_premium` | `25_000.0` | Lump-sum into UL; AV grows from there |
| `premium_load_pct` | `0.06` | Industry-typical front-end load |
| `monthly_expense_charge` | `7.50` | Industry-typical UL flat $ |
| `declared_rate_annual` | `0.04` | Plausible UL crediting rate |
| `db_type` | `"level_face"` | Type A |
| `horizon_age` | `120` | |

**Engine.**

* `account_value.evolve_account_value(...)` with a flat declared rate.
* `monthly_coi_q[t]` = `q_x` from 2017 CSO at attained-age-month-t,
  applied to NAR = `max(0, face - AV[t])`.
* Liability cashflow: `DB[t] * P(death month t)` where
  `DB[t] = max(face, AV[t])` (Type A reduces to face when AV < face).
* **AV-runs-out (lapse) handling:** if `AV[t]` reaches 0 before horizon,
  the contract terminates: `expected_total_cashflows[s] = 0` for every
  `s >= t` and `survival[s]` is held flat at `survival[t-1]` for stats.
  This is a **policy lapse from depletion**, distinct from the optional
  static-lapse decrement layered on top.
* `db_type` is hard-coded to `"level_face"` (Type A) in v1; UI does not
  expose a selector. The dataclass field exists so v2 can add Type B
  (`face + AV`) without breaking the schema.
* Capabilities: `supports_economic_scenario=False`, `supports_monte_carlo=False`.

### Phase 6 — Indexed Universal Life (IUL)

**Concept.** UL with annual P2P crediting (using `crediting` module from
Phase 0). All other UL mechanics unchanged.

**Default illustrative inputs:** like UL, plus:

| Input | Default | Rationale |
|-------|---------|-----------|
| `participation` | `1.00` | Industry-typical for capped IUL |
| `cap` | `0.10` | 10% cap, plausible |
| `floor` | `0.0` | Standard IUL floor |
| `index_scenario_csv_path` | reuse S&P 500 baseline | Deterministic |

**Engine.**

* `account_value.evolve_account_value(...)` with `monthly_credit_rate`
  computed from `crediting.AnnualPointToPointCapped(...).credit_segment(...)`
  on segment anniversaries (zero in non-anniversary months).
* All other UL mechanics identical, including AV-depletion lapse handling
  and `db_type = "level_face"` v1 default.
* Capabilities: `supports_economic_scenario=True`,
  `supports_monte_carlo=True` (reuse GBM index simulator).

### Phase 7 — Variable Universal Life (VUL)

**Concept.** UL with sub-account return as the credit. All other UL
mechanics unchanged. Reuses VA's sub-account model.

**Default illustrative inputs:** like UL, plus:

| Input | Default | Rationale |
|-------|---------|-----------|
| `subaccount_drift_annual` | `0.06` | |
| `subaccount_vol_annual` | `0.15` | |

**Engine.**

* `account_value.evolve_account_value(...)` with monthly sub-account return
  from a deterministic CSV or GBM path (re-uses VA's sub-account source).
* All other UL mechanics identical, including AV-depletion lapse handling
  and `db_type = "level_face"` v1 default.
* Capabilities: `supports_economic_scenario=True`,
  `supports_monte_carlo=True`.

---

## 4. Phase 8 — Streamlit UI completion + comprehensive regression

By the end of Phases 1–7, every product has its sidebar/contract widget
block in place because step F of the per-product template required it.
This phase is the **end-to-end UI smoke and full regression**:

1. **`tests/ui/test_apptest_full_workflow.py`** is verified to walk all 10
   implemented products (it auto-derives from `implemented_product_types()`).
   Asserts each product's Pricing Run page renders, that "Run pricing"
   produces a `pricing_res`, and that the Excel download bytes pass
   `excel_workbook_validator`.
2. **All 10 products' AppTest smokes pass** (`tests/ui/test_apptest_<P>.py`).
3. **`scripts/deep_smoke.py`** exercises all 10 products + the 4 ALM-enabled
   variants we want to keep ahead of regression (SPIA + RILA + UL + IUL).
4. **`tests/parity/test_excel_recalc_per_product.py`** runs both layers
   (always-on Python literal check + ModelCheck formula-link contract) for
   all 10 products.
5. **Mypy strict glob test** (`tests/test_mypy_strict_glob.py`) confirms
   `products.<new>.engine`, `.excel`, `.schema`, `.ui` are all picked up.
6. **Mutmut PR gate** runs against all touched parity-critical files; zero
   surviving mutants per `mutmut_thresholds.toml` `default = 0`.
7. **Hypothesis property tests** run with `_FAST_SETTINGS` for the new
   products' invariants.

**Failure mode handling.** If any AppTest or recalc gate fails, the failing
product's `tests/ui/test_apptest_<P>.py` and
`tests/parity/test_<P>_parity.py` are the local repro path; the runbooks
in `docs/runbooks/` apply unchanged.

---

## 5. Phase 9 — Documentation, release prep, sign-off

1. **`docs/CHANGELOG.md`**: one entry per product under `[Unreleased]`.
2. **`docs/model_change_log.md`**: one consolidated entry for each
   parity-impacting addition (all products, plus lapse / CSO / crediting /
   AV framework introductions).
3. **`docs/model_parity_contract.md`** auto-renders from
   `parity_constants.py` via `scripts/render_parity_contract.py`. **Never
   edited by hand.** Run `--check` to confirm sync.
4. **`docs/parity_test_checklist.md`** updated if new always-on items
   appear (e.g. "10 products' golden JSON green").
5. **`README.md`** ("annuity_model" and root) product list updated;
   architecture mermaid extended.
6. **`docs/glossary.md`** updated with VA, MYGA, FIA, WL, UL, IUL, VUL,
   GMDB, NAR, COI, AV, M&E, segment crediting.
7. **Per-product spec docs** (`docs/<P>_product_spec.md`) — one-pager each.
8. **Per-product parity contract addenda** (`docs/<P>_parity_contract.md`)
   only if a product introduces a new tolerance.
9. **CODEOWNERS** updated to route new files to the same reviewer set as
   the existing parity-critical surface.
10. **`.github/workflows/`** — no changes required if `parity-gate.yml` and
    `ci.yml` already enumerate `pytest tests/parity/` and `deep_smoke.py`
    (they do).
11. **Final four-gate run** locally + CI; `just preflight` prints
    `READY TO COMMIT`.

---

## 6. Testing strategy — comprehensive

The platform already has 9 distinct test families (parity, unit, ui, meta,
regression matrix, property, mutmut, kit-template, perf). The plan adds
**zero new test families**; instead each existing family is extended once
per product.

### 6.1 Per-product test additions (10 products total)

| Test family | New file(s) per new product | Purpose |
|-------------|---------------------------|---------|
| Parity (always-on) | `tests/parity/test_<P>_parity.py` | Per-month engine invariants + ModelCheck reconciliation within `LIFE_MODELCHECK_TOL` / `ANNUITY_ACCUM_MODELCHECK_TOL` |
| Parity (golden) | `tests/parity/golden/<P>.json` | Byte-exact ModelCheck snapshot; updates only via `UPDATE_GOLDEN_MODELCHECK=1` |
| Parity (workbook) | one entry in `_CASE_BUILDERS` | Always-on cached check + ModelCheck formula-link contract |
| Regression matrix | one fixture in `_FIXTURE_BUILDERS` | Surface coverage cube |
| Unit (engine) | `tests/test_<P>_projection.py` | Engine-level branch coverage |
| Unit (Excel) | none new — covered by `tests/test_excel_export_validation.py` parametrized over `implemented_product_types()` | |
| UI AppTest | `tests/ui/test_apptest_<P>.py` | Streamlit page smoke |
| Property (hypothesis) | new functions in `tests/test_property_invariants.py` | Mathematical laws |
| Meta-invariants | none — they auto-cover new products via `implemented_product_types()` | |

### 6.2 Foundation test additions

* `tests/test_lapse.py` — combined survival, decrement composition,
  monthly rate conversion, edge cases (zero rates, ultimate kicks in).
* `tests/test_crediting.py` — `FixedDeclaredRate` and
  `AnnualPointToPointCapped` strategies, including bounds, monotonicity,
  RILA back-compat (RILA's existing `segment_credited_return` returns
  identical values).
* `tests/test_account_value.py` — AV non-negativity, COI math, NAR
  computation, premium load handling, monthly cycle reproduction.
* `tests/test_mortality_2017_cso.py` — loader, sex × smoker dispatch,
  monotonicity in age, sanity bands.

### 6.3 Cross-cutting regression checkpoints

After **each** internal phase boundary:

```bash
# Local equivalent of CI parity-gate:
just preflight           # = parity + pytest + deep_smoke + render_parity_contract --check
pytest tests/ui          # AppTest UI smoke
pytest tests/test_meta_invariants.py tests/test_builder_registry_invariants.py \
       tests/test_data_registry_invariants.py tests/test_products_registry.py \
       tests/test_products_subpackage_shims.py tests/test_mypy_strict_glob.py
```

### 6.4 Existing-product regression guardrails

The plan **changes zero numerics** for SPIA / Term / RILA. Their golden
JSON files are byte-exact, no new tolerances apply to them, and no engine
modifications are made (the only RILA change is a back-compat refactor to
use `AnnualPointToPointCapped` internally — covered by the unchanged
golden JSON).

### 6.5 Mutmut PR gate

Every new parity-critical module is auto-included in the
`MUTMUT_SURFACE` glob via the existing pattern. The PR gate runs only on
touched files, so it remains fast. New modules must have zero surviving
mutants per `mutmut_thresholds.toml`'s `default = 0`.

### 6.6 Property-based laws (hypothesis)

| Product | Law |
|---------|-----|
| MYGA | `AV(T) == single_premium * (1 + declared_rate)^guarantee_years` (deterministic) |
| FIA | per-segment credit ∈ `[floor, cap]`; AV non-decreasing if floor ≥ 0 |
| VA | GMDB cashflow ≥ 0; AV ≥ 0 always |
| WL | `pv_benefit > 0` for any positive face; monotone in face |
| UL | AV ≥ 0; COI ≤ qx*face; if AV ≥ face, NAR == 0 |
| IUL | per-segment credit ∈ `[floor, cap]`; falls back to UL when cap = 0 = floor |
| VUL | AV ≥ 0; with zero subaccount return AV trajectory matches no-credit UL within `AV_TOL` |
| Lapse framework | combined_survival monotone non-increasing in `t` |

### 6.7 Always-on UI workflow gate

`tests/ui/test_apptest_full_workflow.py` already iterates
`implemented_product_types()`, so it auto-grows from 3 to 10 products.
The plan does **not** add a parallel test; it relies on the existing
contract.

---

## 7. Streamlit UI specifics — non-repetition strategy

The current `pricing_ui.py` is 4,467 LOC; the per-page split (`ui/pages/`)
is gated on shrinking it below 1,500 LOC. Adding seven products without
exploding the file is a real risk.

**Tactics already in place we must use:**

* `get_product_capabilities(product_type)` flags drive whether the index
  scenario / Monte Carlo expanders render — extending this to 10 products
  is mechanical.
* `get_product_ui_config(product_type)` provides the per-product info
  message and download filenames — already supports unimplemented products
  with placeholder messages, so we just fill in the real labels.
* `get_pricing_metrics(product_type, result)` is dispatcher-driven and
  doesn't grow `pricing_ui.py` at all.
* `_PRODUCT_VALIDATORS` registry is empty today; each new product's
  pre-flight constraints land here, not as inline `if` blocks.

**New convention to enforce:**

* The per-product contract widget block (~lines 2218–2309) and the
  contract construction block (~lines 2573–2606) are the *only*
  acceptable places for `if/elif product_type == X` in `pricing_ui.py`.
* Anything else needing per-product behavior MUST go through a registry
  (`_PRICING_METRIC_FORMATTERS`, `get_product_capabilities`,
  `get_product_ui_config`, or a new registry added in this PR).

**Estimated `pricing_ui.py` LOC delta after this PR:** roughly
`+70 LOC × 7 products = +500 LOC` (contract widget + construction blocks
only — life products with smoker class + face + premium-load + monthly
expense + crediting params trend toward the higher end). Total ≈ 5,000 LOC.
Still gated above the 1,500 threshold for the per-page split, but the
threshold is a target for a *separate* future refactor (`ui/MIGRATION.md`).

**New `RUN_KEY` entries** in `pricing_run_form_state.py` (each goes into
`RUN_KEY` and, where numeric, into `PRICING_RUN_NUMBER_INPUT_KEYS`):

| Product | New keys |
|---------|----------|
| MYGA | `MYGA_SINGLE_PREMIUM`, `MYGA_DECLARED_RATE`, `MYGA_GUARANTEE_YEARS` |
| FIA  | `FIA_PARTICIPATION`, `FIA_CAP`, `FIA_FLOOR`, `FIA_RIDER_FEE`, `FIA_HORIZON_YEARS` |
| VA   | `VA_SINGLE_PREMIUM`, `VA_ME_CHARGE_ANNUAL`, `VA_GMDB_BASIS`, `VA_SUBACCT_DRIFT`, `VA_SUBACCT_VOL` |
| WL   | `WL_FACE_AMOUNT`, `WL_SMOKER_CLASS` |
| UL   | `UL_FACE_AMOUNT`, `UL_SMOKER_CLASS`, `UL_SINGLE_PREMIUM`, `UL_PREMIUM_LOAD_PCT`, `UL_MONTHLY_EXPENSE`, `UL_DECLARED_RATE` |
| IUL  | `IUL_FACE_AMOUNT`, `IUL_SMOKER_CLASS`, `IUL_SINGLE_PREMIUM`, `IUL_PREMIUM_LOAD_PCT`, `IUL_MONTHLY_EXPENSE`, `IUL_PARTICIPATION`, `IUL_CAP`, `IUL_FLOOR` |
| VUL  | `VUL_FACE_AMOUNT`, `VUL_SMOKER_CLASS`, `VUL_SINGLE_PREMIUM`, `VUL_PREMIUM_LOAD_PCT`, `VUL_MONTHLY_EXPENSE`, `VUL_SUBACCT_DRIFT`, `VUL_SUBACCT_VOL` |

Per-product keys (rather than reused generic `RUN_PARTICIPATION`) avoid
the cross-product state leak that the existing Term-vs-RILA split also
solves. The `tests/test_run_state_key_drift.py` ratchet auto-picks them
up via `RUN_KEY` reflection.

---

## 8. Excel-recalculation correctness — non-repetition strategy

Each new builder calls into the shared helpers in `excel_builder_helpers.py`
plus the proposed `LifeProductBuilderTemplate` (Phase 0, §1.6). The
constraints on every builder:

1. **`validate_workbook_or_raise(wb)`** before every `wb.save(...)` —
   enforced by code review and `tests/test_excel_export_validation.py`.
2. **`liability_layout_for(product_type)`** for all column-letter access —
   enforced by `excel_workbook_validator`'s cross-sheet column resolution
   check.
3. **No hand-rolled `IF(cond, value)` formulas** — every conditional has
   an explicit false branch; enforced by validator's arity check.
4. **Any new function used in a formula** registers its arity in
   `excel_workbook_validator.FUNCTION_ARITIES` in the same commit.
5. **ModelCheck pattern reuses `write_model_check_sheet`** — no
   per-product ModelCheck implementation.
6. **Per-product workbook contract case** in
   `tests/parity/test_excel_recalc_per_product.py::_CASE_BUILDERS` — the
   gate `test_every_implemented_product_has_a_recalc_case` enforces
   presence.

### 8.1 Proposed `LIABILITY_LAYOUTS` entries

Defining these up front keeps the cross-sheet validator happy and avoids
the SPIA-S vs RILA-M class of bug the existing layout registry exists to
prevent. All seven products land with their own layout entry in Phase 0
(but the entry can stay unused until the product's builder lands).

| Product | `total_cf_col` | `discount_col` | Notes |
|---------|----------------|----------------|-------|
| `myga` | `M` | `O` | Same layout as RILA (accumulation product) |
| `fia` | `M` | `O` | Same as RILA |
| `variable_annuity` | `M` | `O` | Same as RILA |
| `whole_life` | `S` | `O` | Same as Term/SPIA (life cashflow shape) |
| `universal_life` | `S` | `O` | Same as Term/SPIA |
| `indexed_ul` | `S` | `O` | Same as Term/SPIA |
| `variable_ul` | `S` | `O` | Same as Term/SPIA |

The choice is driven by which sheet columns the existing shared ALM
helper expects (`liability_total_col` defaults to `S` for SPIA/Term-style
liability layouts and `M` for RILA-style accumulation layouts). New
products opt into one or the other via the layout entry, and the shared
helper is unchanged.

---

## 9. Risk register and mitigations

| Risk | Likelihood | Impact | Mitigation |
|------|-----------|--------|------------|
| RILA back-compat refactor (Phase 0 §1.2) silently changes golden values | Low | High | Keep `segment_credited_return` as the public name (unchanged behavior). RILA golden JSON is byte-exact gate; any drift fails parity. |
| 2017 CSO synthetic data is mistaken for real licensed table | Medium | Medium | Per-file docstring banner; data_registry `source` field labels it placeholder; loader logs a warning when synthetic file is used. |
| Mega-PR is too large to review meaningfully | High | High | Internal phasing produces stand-alone commits; each commit leaves green gates. PR description maps each commit to the corresponding phase. Reviewer can sign off per-phase. |
| `pricing_ui.py` becomes unreadable after +500 LOC | Medium | Low | Hard-cap per-product blocks (~70 LOC each); enforce registry-only path for any non-widget per-product logic. |
| Lapse framework misapplied to existing products silently | Low | High | Default `lapse=None` everywhere; existing engines never call combined survival path; explicit unit test `test_lapse_default_is_no_op_for_<P>`. |
| UL/IUL/VUL implicit single-premium equation infeasibility (à la RILA's `RILAPricingInfeasibleError`) | Medium | Medium | Reuse the RILA pattern: explicit `<P>PricingInfeasibleError` for the AV-can't-cover-COI case. Tests cover the boundary. |
| Desktop spreadsheet automation differences between platforms | Medium | Medium | Automated gates avoid desktop spreadsheet subprocesses; new products use static validation, Python snapshots, and ModelCheck formula-link checks. |
| Property tests find a real bug late | Medium | Low | Hypothesis is fast (~25 examples per case); run on every internal phase. |
| Mutmut surviving mutants in new modules | Medium | Medium | Run `mutmut_pr_gate.py` before each phase commit; address survivors by adding tests. `default = 0` enforces. |
| Static lapse table doesn't cover dynamic-lapse use cases | Low | Low | Documented as v1 limitation in `docs/lapse_framework.md`; `LapseAssumption` extensible to dynamic in v2 without breaking v1 callers. |
| **Agent context window fills mid-phase, work degrades silently** | High (without discipline) | High | Mandatory `!handoff` at every phase boundary (Step S). Per-phase context budget (~25k tokens of grounding, see Section 12.1). Subagent delegation for exploration. Mid-phase recovery protocol (Section 12.8). |
| **Engine produces internally consistent but actuarially nonsense output** (passes parity, fails reality check) | Medium | High | Mandatory actuarial assessment (Step P) per product. Benchmark bands centralized in `docs/actuarial_benchmarks.md`. Closed-form cross-validation where available (Section 13.5). Resolution playbook in Section 13.7. |
| **Reviewer disagrees with a benchmark band** | Medium | Low | Bands are explicit and documented; band changes go through the same review path as tolerance changes (changelog entry + reviewer sign-off). Wide bands by design — they catch order-of-magnitude bugs, not assumption disagreements. |
| **Handoff doc drifts from actual repo state** | Low | Medium | Each handoff includes `git status` + `git rev-parse HEAD` per the template; first action of next phase is to verify those still match. Mismatch = stale handoff, recompute. |

---

## 9.1 How to review this PR

The single mega-PR is large by design (sponsor decision). The internal
phase boundaries are the seams a reviewer should use:

1. **Review Phase 0 in isolation first.** It's the foundation: lapse,
   crediting, AV, CSO mortality, parity_constants additions. Verify the
   four canonical gates green at the Phase 0 commit; this is the
   "everything else builds on this" boundary. RILA's golden JSON
   un-changed despite the back-compat refactor is the single most
   important Phase 0 invariant.

2. **Review each product phase as a stand-alone unit.** Each phase is a
   commit (or contiguous commit cluster) that ends with all four gates
   green and one new product implemented. The PR description maps
   commit hash → phase number. A reviewer can sign off per-phase.

3. **Read the per-product spec docs first.** `docs/<P>_product_spec.md`
   is the actuarial intent; reading it before the engine code makes the
   engine code obvious. Six 1-pagers + one (UL) longer doc = readable
   in ~30 minutes.

4. **Spot-check three representative products in depth.** A reviewer
   who has time for three of the seven should pick **MYGA** (simplest,
   sanity check the patterns), **UL** (introduces the AV equation that
   IUL/VUL extend), and **VA** (the only one with sub-account + GMDB
   semantics).

5. **Trust the meta-invariants for the rest.** If
   `tests/test_meta_invariants.py`,
   `tests/test_builder_registry_invariants.py`, and
   `tests/test_products_registry.py` are green, the wiring is correct
   for every product (these tests exist precisely to make per-product
   wiring review unnecessary).

6. **Excel side: focus on `LifeProductBuilderTemplate`** (Phase 0) and
   one product's `build_<P>_excel_workbook.py`; the others are
   structural copies. Trust
   `tests/parity/test_excel_recalc_per_product.py` to catch behavioral
   drift.

7. **Streamlit side: run the AppTest harness locally**, not just CI.
   `pytest tests/ui -v` shows per-product render success in ~30 seconds.

8. **Last: rebuild the workbook ZIP for one product end-to-end** with
   `streamlit run pricing_ui.py`, click through the run, download the
   `.xlsx`, inspect the `ModelCheck` formulas and confirm they link to the
   expected validated liability summary rows. The runbook
   `docs/runbooks/regenerate_excel_cache.md` walks through this.

---

## 10. Sequencing recap (single mega-PR with internal phase boundaries)

Each phase ends with three required checkpoints in this order: ① actuarial
assessment passes, ② four canonical gates green, ③ phase handoff written.
A new chat session is recommended at the start of every phase to keep the
agent context window healthy (Section 12).

```
Phase 0: Foundation
  ├─ lapse.py              + tests
  ├─ crediting.py          + tests + RILA back-compat refactor
  ├─ account_value.py      + tests
  ├─ MortalityTable2017CSO + 4 data_registry artifacts + tests
  ├─ ProductType extension (5 new members)
  ├─ parity_constants additions + model_change_log entry
  ├─ LifeProductBuilderTemplate
  ├─ docs/actuarial_benchmarks.md (single source for all per-product bands)
  ├─ GATE: 4 canonical gates green; existing 3 products byte-exact
  └─ HANDOFF: !handoff phase-0-foundation
Phase 1: MYGA       — Steps A–S of per-product template (Section 2)
                      ├─ Step P: actuarial assessment vs §13.3 MYGA row
                      ├─ Step R: 4 canonical gates
                      └─ Step S: !handoff phase-1-myga
Phase 2: FIA        — same template; Step P uses §13.3 FIA row
                      └─ HANDOFF: !handoff phase-2-fia
Phase 3: VA         — same template (uses GBM, GMDB); Step P uses §13.3 VA row
                      └─ HANDOFF: !handoff phase-3-va
Phase 4: WL         — same template (uses 2017 CSO); Step P uses §13.3 WL row
                      └─ HANDOFF: !handoff phase-4-wl
Phase 5: UL         — same template (uses account_value); Step P uses §13.3 UL row
                      └─ HANDOFF: !handoff phase-5-ul
Phase 6: IUL        — same template (UL + crediting); Step P uses §13.3 IUL row
                      └─ HANDOFF: !handoff phase-6-iul
Phase 7: VUL        — same template (UL + sub-account); Step P uses §13.3 VUL row
                      └─ HANDOFF: !handoff phase-7-vul
Phase 8: Comprehensive regression
  ├─ AppTest full workflow walks all 10 products
  ├─ deep_smoke walks all 10 products
  ├─ ModelCheck formula-link contract walks all 10 products
  ├─ Mypy strict glob covers all new shims
  ├─ Mutmut PR gate green
  ├─ Hypothesis property gates green
  ├─ Actuarial assessments green for all 7 new products
  └─ HANDOFF: !handoff phase-8-regression
Phase 9: Documentation + release prep
  ├─ CHANGELOG, model_change_log, parity contract auto-render
  ├─ README, glossary, per-product spec docs
  ├─ CODEOWNERS, four-gate final run
  ├─ Verify every phase handoff doc still aligns with current HEAD
  └─ READY TO COMMIT (final !handoff phase-9-release archives the trail)
```

---

## 11. Definition of done (whole PR)

* [ ] All 10 products in `implemented_product_types()`.
* [ ] All four canonical gates exit 0 locally and on CI.
* [ ] `tests/parity/test_excel_recalc_per_product.py` `_CASE_BUILDERS` has
      10 entries; `test_every_implemented_product_has_a_recalc_case` green.
* [ ] All 10 golden JSON files byte-exact; `MODELCHECK_TOL` still `0.0`.
* [ ] `tests/test_meta_invariants.py`, `tests/test_builder_registry_invariants.py`,
      `tests/test_products_registry.py`, `tests/test_products_subpackage_shims.py`,
      `tests/test_mypy_strict_glob.py`, `tests/test_data_registry_invariants.py`,
      `tests/test_observability_wiring.py`, `tests/test_run_state_key_drift.py`
      all green.
* [ ] `tests/ui/test_apptest_full_workflow.py` covers all 10 products.
* [ ] `scripts/deep_smoke.py` exits 0 with all 10 products.
* [ ] `scripts/render_parity_contract.py --check` green; the rendered
      contract reflects every new tolerance constant.
* [ ] `docs/model_change_log.md` has one consolidated entry covering this PR
      (foundation + each implemented product).
* [ ] `docs/CHANGELOG.md` entries under `[Unreleased]` — one per product,
      one for foundation, one for the lapse / crediting / account_value /
      CSO additions.
* [ ] No copyrighted CSO data is committed; placeholder synthetic CSV with
      documented warning + override path is the shipped artifact.
* [ ] Mutmut PR gate green on all touched parity-critical files.
* [ ] **Coverage ratchet** (`scripts/ratchet_coverage.py`) holds at the
      committed `fail_under` (currently `55.0`); never *raised* in this PR
      to avoid coupling product additions with a coverage bump (separate
      concern, separate review).
* [ ] `pricing_ui.py` LOC growth ≤ 600 LOC and confined to per-product
      contract widget + construction blocks (no new cross-cutting hardcodes).
* [ ] **CODEOWNERS** updated to route the new files (engines, builders,
      `products/<P>/`, parity tests, golden JSON) to the same reviewer set
      as the existing parity-critical surface.
* [ ] **Existing 3 products' golden JSON unchanged** — verified by the
      ratchet behavior of `test_golden_modelcheck.py` (no `UPDATE_GOLDEN_*`
      env var was set during the run).
* [ ] **Actuarial assessment green for every implemented product**
      (`tests/parity/test_<P>_actuarial.py` for each of the 7 new
      products). No band was widened to make a test pass; if a band
      shifted, the change is documented in `docs/actuarial_benchmarks.md`
      with sign-off rationale.
* [ ] **Closed-form cross-validation passes** (Section 13.5) for every
      product where one applies (MYGA, FIA-floor=cap=0, IUL-cap=floor=0,
      VUL-σ=0 reduces to UL).
* [ ] **Sensitivity signs correct for every product** (Section 13.4).
      No directional reversal anywhere.
* [ ] **Phase handoff trail complete**: `.cursor/handoffs/` contains one
      handoff per phase (`phase-0-foundation`, `phase-1-myga`, ...,
      `phase-9-release`), each capturing the exit state per the template
      in Section 12.5 and the rule in `.cursor/rules/handoff-recall.mdc`.
* [ ] **Each handoff still aligns with current HEAD** — verified during
      Phase 9. A drifted handoff means a phase was not properly closed;
      reconcile before sign-off.
* [ ] `just preflight` prints `READY TO COMMIT`.

---

## 12. Context window hygiene for the AI agent

The mega-PR runs across many implementation sessions. Each session has a
finite context window; without active discipline the agent will degrade by
mid-Phase-4 and produce subtly wrong code by Phase 6. This section codifies
the techniques that keep the agent operating in a clean context for every
phase.

### 12.1 Per-phase context budget

**Target: each phase begins with ≤ 25k tokens of grounding context loaded.**
That leaves ample headroom for the actual implementation work (reading other
files, writing code, running tests, debugging).

Concrete budget breakdown:

| Item | Tokens | Notes |
|------|--------|-------|
| `annuity_model/AGENTS.md` ("Before completing any task" section only) | ~2k | Use offset/limit; never read the whole file. |
| This plan's Section 3 entry for the current product (only) | ~5k | Use offset/limit. The full plan is ~30k; reading it all every phase is the primary waste. |
| `docs/<P>_product_spec.md` | ~1k | Per-product 1-pager (created in Phase 0 prep or first thing in each phase). |
| `docs/actuarial_benchmarks.md` (just the row for the current product) | ~1k | Use offset/limit. |
| Prior phase's handoff doc (`.cursor/handoffs/<...>-phase-N-1-*.md`) | ~1k | The single most important context-bridging artifact. |
| ONE template engine file (e.g. `rila_projection.py` or `term_projection.py`) | ~5k | Whichever is the closest analog for the current product. |
| ONE template Excel builder | ~10k | Same logic — closest analog. |
| `tests/parity/test_<template>_parity.py` | ~3k | The test scaffold to copy. |
| **Subtotal** | **~28k** | |

Anything beyond this is loaded on-demand via Grep + Read offset/limit.

### 12.2 Bootstrap routine — start of every phase

When a fresh chat opens for Phase N, the agent's first actions in order:

1. `!recall <prior-phase-slug>` (e.g. `!recall phase-0-foundation` for
   the start of Phase 1). If working in the same chat as the prior phase,
   prior context is already loaded; skip this.
2. Read this plan's Section 3 entry for the current product **only** —
   use `Read` with `offset` and `limit` to avoid pulling the entire 900+
   line plan.
3. Read `annuity_model/AGENTS.md` "Before completing any task" section
   (offset/limit; only the canonical-gates block).
4. Read `docs/<P>_product_spec.md` (the per-product 1-pager).
5. Read `docs/actuarial_benchmarks.md` row for the current product.
6. Read **one** template set:
   * For accumulation products (MYGA / FIA / VA): `rila_projection.py`,
     `build_rila_excel_workbook.py`, `tests/parity/test_rila_parity.py`.
   * For life products (WL / UL / IUL / VUL): `term_projection.py`,
     `build_term_excel_workbook.py`, `tests/parity/test_term_parity.py`.
7. **Confirm starting state is green**: run `just preflight` BEFORE
   writing any code. If anything is red, the prior phase didn't close
   cleanly — fix that first or recall a different handoff.

### 12.3 During-phase discipline

| Tactic | Rule |
|--------|------|
| **Grep before Read** | To find a specific symbol or pattern, use `Grep` (returns matched lines). Do not Read whole files to scan for a symbol. |
| **Offset/limit reads** | When you need a specific section of a long file (e.g. `pricing_ui.py` lines 2218–2309), pass `offset` and `limit` to `Read`. Never read whole files like `pricing_ui.py` (4,467 LOC) or `pricing_projection.py` (2,279 LOC). |
| **Subagents for exploration** | For multi-file questions (e.g. "how is X used across the codebase?"), launch a `Task` with `subagent_type="explore"`. The subagent has its own context window; only the summary returns to the parent. |
| **Subagents for parallel work** | For independent sub-tasks (e.g. drafting an engine while writing its Excel builder), launch parallel `Task`s with `subagent_type="generalPurpose"`. |
| **Document-driven, not code-driven** | The plan + spec docs are the design source of truth. Do not re-derive design from existing engine code each phase. The existing code is a structural template only. |
| **Don't reload contracts** | Tests like `test_meta_invariants.py`, `test_run_state_key_drift.py` are CONTRACTS — read them once at PR start, not at the start of every phase. |
| **One feature per turn when possible** | Avoid "do all of Step F then all of Step I in one turn" when each could be a checkpoint. Smaller turns = easier recovery on context warnings. |

### 12.4 Mandatory handoff cadence

The `.cursor/rules/handoff-recall.mdc` protocol is the persistence
mechanism between sessions.

| Trigger | Action |
|---------|--------|
| End of every phase (Step S) | User types `!handoff phase-N-<product>` (e.g. `!handoff phase-1-myga`). Required. |
| Mid-phase context warning | If the agent reports "context filling" or starts hallucinating file paths, type `!handoff phase-N-<product>-mid-<topic>` and open a fresh chat. |
| Long debugging session | After ~30 min of debugging without progress, save a `phase-N-<product>-debug-<topic>` handoff and start fresh. |
| Before starting a new phase | First action in the new chat is `!recall <prior-phase-slug>`. |

**The handoff doc is the contract; the chat transcript is ephemeral.**

### 12.5 Per-phase exit handoff template

Each handoff file follows the template in `.cursor/rules/handoff-recall.mdc`.
Specifically for THIS work, fill these sections precisely:

* **§1 Original goal:** "Implement Phase N (<product>) per
  `annuity_model/docs/seven_product_rollout_plan.md`"
* **§2 Current status:** "All per-product steps A–S complete. Four canonical
  gates green. Actuarial assessment passed. ModelCheck reconciles to 0.00
  on regenerated workbook."
* **§3 Decisions made:** Any non-obvious actuarial or technical choice
  (e.g. "FIA segment_months = 12 with no monthly cap option in v1").
* **§4 Files touched:** Explicit list of `<P>_projection.py`,
  `build_<P>_excel_workbook.py`, `products/<P>/*`,
  `tests/parity/test_<P>_parity.py`, `tests/parity/test_<P>_actuarial.py`,
  `tests/parity/golden/<P>.json`, etc.
* **§5 Key commands:** `just preflight` exit code, last
  `pytest tests/parity` summary, `deep_smoke.py` output.
* **§6 Open questions:** Any deferred decisions for the next phase.
* **§7 Next concrete step:** "Begin Phase N+1 (<next-product>) per Section 3
  of the plan; bootstrap via Section 12.2."
* **§8 Gotchas:** e.g. "FIA participation > 1.0 makes pricing infeasible";
  "UL with single_premium < $X depletes AV before age 90";
  "actuarial benchmark band for <product> is currently at the lower end —
  next reviewer should consider tightening".
* **§9 Relevant rules / docs consulted:** Always cite the plan, the
  product spec, and `docs/actuarial_benchmarks.md`.

### 12.6 Subagent delegation patterns

| Need | Subagent type | Example prompt |
|------|---------------|----------------|
| Find usages of a symbol across the codebase | `explore` | "Find every call site of `liability_path_for` and report the file/line and surrounding context." |
| Verify a design decision against existing patterns | `explore` | "Compare how SPIA, Term, and RILA handle X. Report the consistent pattern or document the divergence." |
| Run an exploratory math sanity check | `generalPurpose` | "Compute the closed-form SP-WL premium for issue_age=45, face=$250k, rate=4%, q_x from CSO-2017-male-nonsmoker placeholder. Compare to the engine output in `tests/parity/test_wl_actuarial.py::test_wl_sp_within_band`." |
| Set up a debugging environment | `shell` | Branch creation, env setup, test runs in a worktree. |

**Critical: NEVER delegate the actual product implementation to a subagent.**
Implementation requires sustained context that subagent isolation breaks.
Only exploration, sanity checks, and infrastructure tasks delegate.

### 12.7 Document-driven, not code-driven

The phase pattern is:

1. READ the spec doc → understand intent.
2. READ the plan section → understand the wiring.
3. WRITE the engine → implementation.
4. WRITE tests (including actuarial in Step P) → validation.
5. VERIFY gates → confirm.
6. WRITE handoff (Step S) → persist.

The agent should never "look at how RILA does X to figure out what to do
for FIA". The plan + spec doc tell the agent what to do; the existing
code is a structural template only, not a design source.

### 12.8 Recovery if context fills mid-phase

**Symptoms:** the agent starts referencing wrong file paths, hallucinating
function names, making the same correction twice, or producing
inconsistent type signatures. **Action:**

1. STOP coding immediately.
2. Run all gates to capture current state: `just preflight`.
3. Save partial handoff: `!handoff phase-N-<product>-mid-recovery`.
4. Start a fresh chat. First action: `!recall phase-N-<product>-mid-recovery`.
5. Re-read just the spec doc + plan section + the partial code that was
   written so far (use Grep to locate the partial files; Read with
   offset/limit to inspect them).
6. Continue from where the partial work left off.

**Goal: NO phase requires more than ~30 minutes of sustained agent
attention before a checkpoint.** Most phases will fit comfortably; the
larger ones (Phase 5 UL, Phase 8 regression) may need 2–3 sub-handoffs
within the phase.

### 12.9 What NOT to do

| Anti-pattern | Why it's bad | Better |
|--------------|-------------|--------|
| "Just read the whole `pricing_ui.py` to find the contract widget block" | 4,467 LOC = ~50k tokens wasted | `Grep` for `selected_product == ProductType.RILA`; `Read` with offset/limit around the match |
| "Re-read the plan from the top each phase" | 900+ lines each time = 30k tokens × N phases | Read only the Section 3 subsection for the current product |
| "Iterate by reading test failures, no separate handoff" | Transcript fills with stack traces; degrades by Phase 4 | After each phase: handoff + new chat |
| "Trust me, I remember the design" | The agent doesn't actually remember across sessions; the handoff is the memory | Always recall before starting; cite the handoff in the work |
| "Run full pytest after every code edit" | 30s × many edits = wasted context with output text | Run focused subset first (`pytest tests/parity/test_<P>_parity.py -x`); full suite at gate time |

---

## 13. Actuarial assessment per product

Parity tests prove **Excel == Python**. Property tests prove **mathematical
laws hold**. Neither catches **"the engine produces a numerically
self-consistent answer that is actuarially wrong"** (e.g. WL single
premium for a 45-year-old male = $1; passes parity, totally unrealistic).

This section adds an **actuarial reasonableness gate per product**. It is a
SEPARATE TEST from parity, with its own tolerance philosophy: bands are
**wide** (real-world product pricing depends on assumption choices), but
a result outside the band is a **red flag** that requires investigation
and engine fix — never band widening.

### 13.1 Definition of "actuarially appropriate"

A pricing engine output is actuarially appropriate iff:

a. **Sign correctness** — every quantity has the expected sign
   (PV ≥ 0, survival ∈ [0, 1], AV ≥ 0).
b. **Order-of-magnitude correctness** — every quantity falls within a
   documented benchmark band (Section 13.3).
c. **Sensitivity correctness** — directional response to assumption
   shocks matches actuarial intuition (Section 13.4).
d. **Closed-form correctness** — where a closed-form benchmark exists,
   the engine matches it within tight tolerance (Section 13.5).
e. **Cross-product consistency** — where two products should produce
   comparable outputs (e.g. SP-WL vs UL with declared rate ≈ COI), they do
   (Section 13.5).

### 13.2 Universal sanity assertions (every product)

Add to every per-product engine test (`tests/parity/test_<P>_actuarial.py`):

```python
def test_<P>_actuarial_sanity_signs():
    res = price_<P>(...)  # baseline scenario
    # a) Signs
    assert res.single_premium >= 0
    assert res.pv_benefit >= 0
    assert (res.survival_to_payment >= 0).all()
    assert (res.survival_to_payment <= 1).all()
    diffs = np.diff(res.survival_to_payment, prepend=1.0)
    assert diffs.max() <= 1e-10  # survival monotone non-increasing
    assert (res.discount_factors > 0).all()
    assert (res.discount_factors <= 1).all()  # for non-negative rates
    if hasattr(res, "account_value_end_month"):
        assert (res.account_value_end_month >= 0).all()  # AV never negative
```

### 13.3 Per-product benchmark bands

These are the order-of-magnitude smell tests. Bands are wide enough that a
typical illustrative scenario falls inside, but tight enough that a
calculation bug pushes outside.

**Single source of truth:** band constants live in a new module
`actuarial_benchmarks.py` (sibling of `parity_constants.py`); the rationale
narrative lives in `docs/actuarial_benchmarks.md`. Tests import constants
by name from the Python module; reviewers read the docs to understand WHY
a band is what it is. The two files are kept in sync via a
`scripts/render_actuarial_benchmarks.py --check` gate (mirrors the
existing `render_parity_contract.py` pattern). Tightening / loosening any
band is a one-file Python edit + a one-paragraph doc change in the same
commit.

| Product | Scenario | Quantity | Benchmark band | Source / rationale |
|---------|----------|----------|----------------|---------------------|
| **MYGA** | $100k SP, 4.5% rate, 5y, age 60 | `account_value_end_month[-1]` | $124,500 – $124,800 | Closed-form: $100k × 1.045^5 = $124,618 |
| **MYGA** | Same | PV(maturity payout) at 4.5% flat | $99k – $101k | If discount = declared rate, PV ≈ premium × survival(T) |
| **FIA** | $100k SP, 80% pop, 7% cap, 0% floor, 10y, age 60, S&P baseline | `account_value_end_month[-1]` | $115k – $200k | Floor 0 / cap 0.07 × 0.8 part = up to 5.6%/yr; 10y bound ∈ [0%, 73%] |
| **VA** | $100k SP, 6% drift, 1.4% M&E, 20y, age 55 | `E[account_value_end_month[-1]]` (deterministic flat-S&P scenario) | $100k – $110k (flat path); $170k – $260k (Monte Carlo mean over GBM paths) | Lognormal moment: exp((0.06-0.014)×20) ≈ 2.51 → ~$251k MC mean |
| **WL** | $250k face, age 45 male NS, 4% flat, CSO-2017-NS placeholder | `single_premium` | $40k – $90k | Industry SP-WL pricing range; depends heavily on mortality table |
| **UL** | $250k face, $25k SP, age 45 male NS, 4% credit, 4% flat | `account_value_end_month[20*12-1]` | $15k – $55k | After 20y of COI + expense, declared rate barely covers |
| **UL** | Same | "AV depletes by attained age" | 70 – 95 | Roughly when COI rises faster than credit |
| **IUL** | Same as UL but 80% pop, 10% cap, 0% floor | `account_value_end_month[20*12-1]` | ≥ UL value at 20y | IUL with floor 0 dominates UL with same declared rate when index ≥ 0 cumulative |
| **VUL** | Same as UL but 6% drift, 15% vol | `E[account_value_end_month[20*12-1]]` (MC mean) | ≥ UL value at 20y at 4% credit | Higher expected return |

### 13.4 Sensitivity matrix

For each product (where applicable), assert:

| Shock | Expected response | Test pattern |
|-------|-------------------|--------------|
| Yield curve **+100bps** | PV of future cashflows ↓ for life products (more discounting) | `assert PV(yc+100bps) < PV(base) - epsilon` |
| Mortality **×1.10** | SP ↑ for life products (more death claims) | `assert SP(qx×1.10) > SP(base) + epsilon` |
| Mortality **×1.10** | PV ↓ for annuities paying alive (fewer survivors) | `assert PV_annuity(qx×1.10) < PV_annuity(base) - epsilon` |
| Index **+10% scenario** | AV ↑ for FIA / VA / IUL / VUL | `assert AV(idx×1.10) > AV(base) + epsilon` |
| Cap **↑** | Expected AV ↑ for FIA / IUL | `assert AV(cap=0.12) > AV(cap=0.08) + epsilon` |
| Floor **↑** | Expected AV ↑ for FIA / IUL (less downside) | `assert AV(floor=-0.02) < AV(floor=0.0) + epsilon` |
| Sub-account drift **↑** | Expected AV ↑ for VA / VUL | `assert E_AV(drift=0.08) > E_AV(drift=0.06) + epsilon` |
| Sub-account vol **↑** | GMDB option value ↑ for VA (floor more in-the-money) | `assert PV_gmdb(vol=0.20) > PV_gmdb(vol=0.10) + epsilon` |

A failed sensitivity sign is **never** a tolerance issue; it is always a
real engine sign bug. Resolution: see Section 13.7.

### 13.5 Closed-form cross-validation

| Product | Closed form | Engine output | Tolerance | Constant |
|---------|-------------|---------------|-----------|----------|
| MYGA | `AV(T) = SP × (1 + i)^T` | `res.account_value_end_month[-1]` | `1e-2` ($) | `MYGA_CLOSED_FORM_AV_TOL` |
| WL | `SP = face × Σ v^t × _{t-1\|}q_x` (NSP_x with monthly granularity) | `res.single_premium` | `1.00` ($) | `WL_NSP_TOL` |
| FIA at floor=cap=0 | `AV(T) = SP` (no growth) | Numerical | `1e-6` ($) | `AV_TOL` |
| IUL with cap=floor=0 | reduces to no-credit UL with the same declared rate | Numerical | `AV_TOL` | `AV_TOL` |
| VUL with σ=0, drift=r | reduces to UL with declared rate `r` | Numerical | `AV_TOL` | `AV_TOL` |
| VA with M&E=0, deterministic flat S&P | `AV(T) = SP` (no return, no charge) | Numerical | `AV_TOL` | `AV_TOL` |

These are **strong** assertions — they pin the engine to math that has
exactly one right answer. Failure means an engine bug, not a band
disagreement.

### 13.6 Implementation pattern (one file per product)

`tests/parity/test_<P>_actuarial.py` template:

```python
"""Actuarial reasonableness tests for <P>.

Bands live in actuarial_benchmarks.py (Python constants) with their
rationale in docs/actuarial_benchmarks.md. Tolerances for closed-form
matches live in parity_constants.py.
"""

import numpy as np
import pytest

import <P>_projection as eng
import pricing_projection as sp
from actuarial_benchmarks import (
    P_BENCHMARK_SP_LO,
    P_BENCHMARK_SP_HI,
    P_SENSITIVITY_EPS,
)
from parity_constants import AV_TOL

pytestmark = [pytest.mark.parity, pytest.mark.product_<P>]


def _baseline_contract():
    # ... canonical scenario defined in this file (and ONLY this file)
    ...


def test_<P>_actuarial_sanity_signs():
    """Section 13.2 universal sanity checks."""
    ...


def test_<P>_single_premium_within_benchmark_band():
    """Section 13.3 order-of-magnitude check.

    Failure means the engine is producing actuarially nonsense numbers.
    Investigate the engine; do NOT widen the band.
    """
    res = eng.price_<P>(contract=_baseline_contract(), ...)
    assert <P>_BENCHMARK_LO <= res.single_premium <= <P>_BENCHMARK_HI, (
        f"<P> SP={res.single_premium:,.0f} fell outside the documented band "
        f"[{<P>_BENCHMARK_LO:,.0f}, {<P>_BENCHMARK_HI:,.0f}]. "
        f"See docs/actuarial_benchmarks.md row '<P>' and Section 13.7 of "
        f"the rollout plan for the resolution playbook."
    )


def test_<P>_yield_sensitivity_negative_pv():
    """Section 13.4 sensitivity sign."""
    base = eng.price_<P>(contract=_baseline_contract(), ...)
    shocked = eng.price_<P>(contract=_baseline_contract(), ...)  # +100bps
    assert shocked.pv_benefit < base.pv_benefit - 1.0, (
        f"+100bps yield shock did not reduce PV(benefit). This is a sign bug."
    )


def test_<P>_closed_form_match():
    """Section 13.5 closed-form cross-validation."""
    ...
```

### 13.7 Resolution playbook when an actuarial check fails

**If a benchmark band fails:**

1. Capture the full pricing result + assumptions in the test failure
   message.
2. Compute the closed-form benchmark separately (in a notebook or by
   hand).
3. Diff: where does the engine diverge from the closed-form?
4. Common root causes:
   * Wrong COI calculation (monthly q_x vs annual; NAR vs face).
   * Wrong discount factor (continuous vs discrete compounding).
   * Wrong index return convention (cumulative vs incremental;
     log vs simple).
   * Wrong cashflow timing (BOM vs EOM; premium-vs-claim alignment).
   * Wrong survival convention (start-of-month vs end-of-month).
   * Wrong sex/smoker dispatch in the mortality lookup.
5. Fix at the engine level. **NEVER widen the benchmark band to make the
   test pass** — that defeats the purpose of the gate.
6. Add a `@pytest.mark.regression` test capturing the exact scenario as a
   permanent guard.

**If a sensitivity sign is wrong:**

1. The engine has a sign bug. This is **never** a tolerance issue.
2. Locate the responsible coefficient (often a `+/-` swap inside an
   accumulator or a wrong default for a `direction` argument).
3. Add a `@pytest.mark.regression` test capturing the sign.

**If a closed-form check fails:**

1. The engine and the closed-form should match within tight tolerance.
2. Failure indicates either (a) the closed-form is misapplied (re-check
   the formula) or (b) an engine bug.
3. Hand-compute one period of the engine cycle to localize.

### 13.8 Where the benchmark numbers come from

The bands in Section 13.3 are derived from:

* **Closed-form formulas** where they exist (MYGA accumulation, WL net
  single premium, lognormal moments for VA / VUL).
* **Industry reference materials**: SOA Educational Material;
  Dickson / Hardy / Waters, *Actuarial Mathematics for Life Contingent
  Risks*; Bowers et al., *Actuarial Mathematics*.
* **Plausibility**: "what would a competent actuary expect to see."

The bands are intentionally **wide** because:

* Mortality tables differ (CSO 2017 vs CSO 2001 vs RP-2014 vs SSA — easily
  ±20% on SP-WL).
* Yield curves differ across snapshots.
* Expense assumptions vary by product line.

But they are tight enough to catch the **"missed a factor of 12"** or
**"wrong sign on COI"** class of bugs that pure parity tests cannot.

### 13.9 Governance

* **One source of truth (code):** band constants live in
  `actuarial_benchmarks.py` (Python). Tests import by name; never inline.
* **One source of truth (rationale):** the why-this-band narrative lives
  in `docs/actuarial_benchmarks.md`. The two are kept in sync by
  `scripts/render_actuarial_benchmarks.py --check` (mirrors the existing
  `render_parity_contract.py` pattern, added in Phase 0).
* **Band changes go through review:** similar discipline to tolerance
  changes in `parity_constants.py`. Each change requires:
  1. Constant edit in `actuarial_benchmarks.py`.
  2. Narrative paragraph appended to a "Band change log" section at the
     bottom of `docs/actuarial_benchmarks.md` explaining the reason.
  3. Reviewer sign-off (CODEOWNERS routes the file to the same reviewer
     set as `parity_constants.py`).
* **Per-product reviewer sign-off:** ideally an actuary reviews the
  baseline scenario + bands for each product before its phase commits.
  When that's not feasible, the implementing agent documents the
  reasoning chain and flags it in the §6 handoff "open questions" for a
  later actuarial review pass.
* **Phase 0 deliverables include the empty `actuarial_benchmarks.py`
  module + the narrative skeleton `docs/actuarial_benchmarks.md`** so
  every per-product phase can populate its row without inventing the
  framework on the fly.

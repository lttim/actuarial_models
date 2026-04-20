# Per-product actuarial benchmark bands

This document is the **rationale** companion to
[`actuarial_benchmarks.py`](../actuarial_benchmarks.py). The Python
module owns the *executable* constants; this doc owns the *why*. The
two are kept in sync by
[`scripts/render_actuarial_benchmarks.py`](../scripts/render_actuarial_benchmarks.py)
(`--check` mode is part of `just preflight`).

> **Discipline.** Bands are intentionally wide enough that a
> reasonable illustrative scenario falls inside, but tight enough that
> a calculation bug pushes outside. **A failing band is investigated
> and fixed at the engine level — never widened to make the test pass**
> (Section 13.7 of the rollout plan).

## Source rationale

Bands derive from one or more of:

* **Closed-form formulas** where they exist (MYGA accumulation, WL net
  single premium, lognormal moments for VA / VUL).
* **Industry reference materials**: SOA Educational Material; Dickson /
  Hardy / Waters, *Actuarial Mathematics for Life Contingent Risks*;
  Bowers et al., *Actuarial Mathematics*.
* **Plausibility**: "what would a competent actuary expect to see?"

Bands are intentionally **wide** because mortality tables differ
(synthetic CSO 2017 placeholder vs licensed CSO 2017 vs CSO 2001 vs
RP-2014 vs SSA — easily ±20% on SP-WL), yield curves differ across
snapshots, and expense assumptions vary by product line.

But they are tight enough to catch the **"missed a factor of 12"** or
**"wrong sign on COI"** class of bugs that pure parity tests cannot.

---

## Generated bands (do not edit by hand)

<!-- BEGIN GENERATED bands -->
| Quantity | Value | Units | Constant |
|----------|-------|-------|----------|
| MYGA AV(T) lower | `1.245e+05` | USD | `actuarial_benchmarks.MYGA_BENCHMARK_AV_T_LO` |
| MYGA AV(T) upper | `1.248e+05` | USD | `actuarial_benchmarks.MYGA_BENCHMARK_AV_T_HI` |
| MYGA PV(maturity) lower | `9.9e+04` | USD | `actuarial_benchmarks.MYGA_BENCHMARK_PV_LO` |
| MYGA PV(maturity) upper | `1.01e+05` | USD | `actuarial_benchmarks.MYGA_BENCHMARK_PV_HI` |
| MYGA closed-form AV(T) tolerance | `1e-02` | USD | `actuarial_benchmarks.MYGA_CLOSED_FORM_AV_TOL` |
| MYGA sensitivity epsilon | `1` | USD | `actuarial_benchmarks.MYGA_SENSITIVITY_EPS` |
| FIA AV(T) lower | `1e+05` | USD | `actuarial_benchmarks.FIA_BENCHMARK_AV_T_LO` |
| FIA AV(T) upper | `2e+05` | USD | `actuarial_benchmarks.FIA_BENCHMARK_AV_T_HI` |
| FIA sensitivity epsilon | `1` | USD | `actuarial_benchmarks.FIA_SENSITIVITY_EPS` |
| VA AV(T) flat-S&P lower | `6e+04` | USD | `actuarial_benchmarks.VA_BENCHMARK_AV_T_FLAT_LO` |
| VA AV(T) flat-S&P upper | `1.1e+05` | USD | `actuarial_benchmarks.VA_BENCHMARK_AV_T_FLAT_HI` |
| VA E[AV(T)] MC lower | `1.7e+05` | USD | `actuarial_benchmarks.VA_BENCHMARK_AV_T_MC_LO` |
| VA E[AV(T)] MC upper | `3.2e+05` | USD | `actuarial_benchmarks.VA_BENCHMARK_AV_T_MC_HI` |
| VA sensitivity epsilon | `1` | USD | `actuarial_benchmarks.VA_SENSITIVITY_EPS` |
| WL single premium lower | `3e+04` | USD | `actuarial_benchmarks.WL_BENCHMARK_SP_LO` |
| WL single premium upper | `1e+05` | USD | `actuarial_benchmarks.WL_BENCHMARK_SP_HI` |
| WL NSP closed-form tolerance | `1` | USD | `actuarial_benchmarks.WL_NSP_TOL` |
| WL sensitivity epsilon | `10` | USD | `actuarial_benchmarks.WL_SENSITIVITY_EPS` |
| UL AV(20y) lower | `5,000` | USD | `actuarial_benchmarks.UL_BENCHMARK_AV_20Y_LO` |
| UL AV(20y) upper | `6e+04` | USD | `actuarial_benchmarks.UL_BENCHMARK_AV_20Y_HI` |
| UL depletion age lower | `70` | Years | `actuarial_benchmarks.UL_BENCHMARK_DEPLETION_AGE_LO` |
| UL depletion age upper | `120` | Years | `actuarial_benchmarks.UL_BENCHMARK_DEPLETION_AGE_HI` |
| UL sensitivity epsilon | `1` | USD | `actuarial_benchmarks.UL_SENSITIVITY_EPS` |
| IUL AV(20y) lower | `5,000` | USD | `actuarial_benchmarks.IUL_BENCHMARK_AV_20Y_LO` |
| IUL AV(20y) upper | `2e+05` | USD | `actuarial_benchmarks.IUL_BENCHMARK_AV_20Y_HI` |
| IUL sensitivity epsilon | `1` | USD | `actuarial_benchmarks.IUL_SENSITIVITY_EPS` |
| VUL E[AV(20y)] MC lower | `5,000` | USD | `actuarial_benchmarks.VUL_BENCHMARK_AV_20Y_MC_LO` |
| VUL E[AV(20y)] MC upper | `2.5e+05` | USD | `actuarial_benchmarks.VUL_BENCHMARK_AV_20Y_MC_HI` |
| VUL sensitivity epsilon | `1` | USD | `actuarial_benchmarks.VUL_SENSITIVITY_EPS` |
| Portfolio total CF sum lower | `2.5e+06` | USD | `actuarial_benchmarks.PORTFOLIO_TOTAL_CF_SUM_LO` |
| Portfolio total CF sum upper | `3.2e+06` | USD | `actuarial_benchmarks.PORTFOLIO_TOTAL_CF_SUM_HI` |
| Portfolio duration gap lower | `-50` | Years | `actuarial_benchmarks.PORTFOLIO_DURATION_GAP_LO` |
| Portfolio duration gap upper | `50` | Years | `actuarial_benchmarks.PORTFOLIO_DURATION_GAP_HI` |
| Portfolio rollup sum consistency tol | `1e-09` | abs | `actuarial_benchmarks.PORTFOLIO_SUM_CONSISTENCY_TOL` |
<!-- END GENERATED bands -->

---

## Per-product narrative

### MYGA — Multi-Year Guaranteed Annuity

* **Scenario.** Single premium $100k, declared rate 4.5%/yr, 5y guarantee, age 60 male.
* **AV(T) band $124.5k–$124.8k.** Closed form is exact:
  `AV(T) = 100,000 × 1.045^5 = 124,618.20`. The band ±$120 absorbs the
  small drift from monthly compounding under the engine's actual
  monthly schedule, plus a tiny survival haircut (5y death prob from
  the synthetic mortality is on the order of 1e-3).
* **PV band $99k–$101k.** When the discount curve equals the declared
  rate, the maturity-payout PV equals the survival-weighted premium —
  ≈ $100k.

### FIA — Fixed Indexed Annuity

* **Scenario.** $100k SP, 80% participation, 7% cap, 0% floor, 10y, age 60, S&P baseline scenario.
* **AV(T) band $100k–$200k.** Floor 0 means AV cannot decrease. Cap × participation
  means upper-bound annual credit is `0.07 × 0.8 = 5.6%`; 10 years compounded gives
  `1.056^10 ≈ 1.73`. The S&P baseline scenario realizes ~mid-band; the wide range
  accommodates segment-by-segment volatility.

### VA — Variable Annuity

* **Scenario.** $100k SP, 6% drift, 1.4% M&E, 20y, age 55.
* **Flat-S&P band $60k–$110k.** With S&P held flat (0% return), 20y of 1.4% M&E
  shrinks AV multiplicatively by `(1 - 0.014/12)^240 ≈ 0.755`, i.e. ~$75k. Band
  widens for slight S&P drift and survival weighting.
* **MC mean band $170k–$320k.** Lognormal moment:
  `E[AV(T)] = 100,000 × exp((μ - M&E) × T) = 100,000 × exp(0.046 × 20) ≈ 251,000`.
  Wide band absorbs survival weighting and stochastic noise at modest n_sims.

### WL — Whole Life (single premium)

* **Scenario.** $250k face, age 45 male NS, 4% flat, CSO 2017 placeholder.
* **SP band $30k–$100k.** Synthetic CSO 2017 placeholder rates are slightly lower
  than published CSO Ultimate at age 45 NS (ratio ~0.7×). Combined with industry
  SP-WL pricing range ($40k–$90k for licensed tables), the lower bound is widened
  to $30k. Production deployments using licensed CSO files should re-tighten
  toward the upper half of this band.

### UL — Universal Life

* **Scenario.** $250k face, $25k SP, age 45 male NS, 4% credit, 4% flat curve.
* **AV(20y) band $5k–$60k.** After 20y of COI + monthly $7.50 expense, the 4%
  declared rate barely covers cost-of-insurance + load. With the synthetic CSO
  placeholder (lower than licensed), AV survives longer than typical industry
  illustrations. Wide band reflects both AV-survival and AV-depletion plausible
  cases.
* **Depletion age band 70–120.** AV depletion typically occurs when COI rates
  rise faster than the credit covers. With placeholder CSO, depletion may not
  occur within the model horizon; the upper bound 120 captures that case.

### IUL — Indexed UL

* **Scenario.** Same as UL but 80% participation, 10% cap, 0% floor.
* **AV(20y) band $5k–$200k.** IUL with floor 0 dominates UL with same declared
  rate when cumulative index is non-negative; it can lag if cap is rarely hit and
  index is flat. Wide band reflects both possibilities.

### VUL — Variable UL

* **Scenario.** Same as UL but 6% drift, 15% vol on the sub-account.
* **MC mean AV(20y) band $5k–$250k.** Higher expected sub-account return
  drives higher expected AV; stochastic distribution is wide so the band
  reflects MC mean ± typical noise across paths.

---

## Sensitivity epsilons

The `*_SENSITIVITY_EPS` constants are the dollar tolerance used in
sign-direction tests (Section 13.4 of the rollout plan). They are
deliberately small ($1–$10) — sensitivity assertions are **directional**:
"yield + 100bps must reduce PV by *at least* `eps`". A failed sensitivity
sign is **always** a real engine sign bug, never a tolerance issue.

---

## Closed-form tolerances

* `MYGA_CLOSED_FORM_AV_TOL = 1e-2` — MYGA AV(T) vs `SP × (1+i)^T`. Tight because
  the engine path equals the closed form to machine precision modulo monthly
  compounding rounding.
* `WL_NSP_TOL = 1.0` — WL net single premium vs the textbook
  `face × Σ v^t × _{t-1|}q_x` formula (computed independently in the
  actuarial test). Allows for $1 of monthly aggregation rounding.
* `AV_TOL = 1e-6` (in `parity_constants`) — used as the closed-form match
  tolerance for FIA-floor=cap=0 (collapses to "no growth"), IUL-cap=floor=0
  (collapses to no-credit UL), and VUL-σ=0 (collapses to UL).

---

## Band change log

When a band tightens or loosens, append a paragraph here AND update the
constant in `actuarial_benchmarks.py` in the same commit.

* **2026-04-19 (initial Phase 0).** All bands above land with the
  initial seven-product rollout. Synthetic CSO placeholder + the
  S&P baseline scenario are the calibration source. Production users
  with licensed CSO data should re-validate WL / UL / IUL / VUL bands
  against their licensed tables and tighten where appropriate.

# RILA product specification (v1 — accumulation)

**Status:** Signed specification for model implementation (Python + Excel parity).
**Scope:** Registered index-linked annuity — **accumulation only** (no GLWB, no ratchets, no multi-segment ladders with overlapping renewals).

## Contract

| Input | Description |
|--------|-------------|
| Issue age / sex | Mortality for monthly death probabilities (same engines as SPIA). |
| Horizon | Projection runs to `horizon_age` on a monthly grid (`dt = 1/12`). |
| Index scenario | Monthly levels `month = 0..N` and `sp500_level` (reuse SPIA CSV loader). Month `k` level is end of policy month `k` (see parity doc). |
| Participation `p` | Applied to segment raw return before cap/floor (e.g. `0.8`). |
| Cap `c` | Annual crediting cap per segment as a **decimal** (e.g. `0.10` = +10%). |
| Floor `f` | Annual crediting floor per segment as a **decimal** (e.g. `0.00`). |
| Rider fee | Annual M&E-style charge on account value, applied as **simple** `AV *= (1 - fee_annual/12)` each month after any segment credit. |
| Segment length | **12 months** — point-to-point on index from end of month `k-12` to end of month `k` for each `k` multiple of 12 (`k >= 12`). |
| Death benefit | **Account value at end of month** after that month’s segment credit (if any) and fee. (ROP / GMDB deferred to a later version.) |

## Crediting mathematics

For a segment ending at month `k` (12, 24, …):

- `R = L_k / L_{k-12} - 1` where `L_j` is the index level at end of policy month `j` (`L_0` = month 0 from scenario).
- `R* = max(f, min(c, p * R))`.
- Account (per dollar of premium basis for pricing): `AV <- AV * (1 + R*)`.

Months without a segment anniversary only apply the monthly fee to `AV`.

## Pricing

- **Relative account:** simulate with initial `AV = 1` through the scenario to obtain monthly **expected death benefits** per $1 premium: `claim_k = death_prob_k * AV_k`.
- Let `K = Σ claim_k * DF(t_k)`. With SPIA-style expenses:
  `premium = (policy_expense + pv_maintenance_expenses + K) / (1 - premium_expense_rate)`
  when `K < 1 - rate` and denominator positive.
- Scale cashflows and reserves by `premium` for reporting.

## Out of scope (v1)

- Buffer / asymmetric designs, volatility control indices, interim high-water marks.
- Surrender charges and partial withdrawals.
- GLWB / income riders.

## Excel replication

- Liability sheet reproduces the same month indexing, survival, discount factors, crediting, fee, and expected claim columns as Python.
- `ModelCheck` compares Python snapshot pricing metrics to Excel summary cells (see `docs/rila_parity_contract.md`).

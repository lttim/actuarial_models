# Lapse / persistency framework (v1)

This document describes the **static lapse table** framework introduced
for the seven-product expansion. All new products
(MYGA, FIA, VA, WL, UL, IUL, VUL) accept an optional
`lapse: LapseAssumption | None = None` parameter; existing products
(SPIA, Term, RILA) remain mortality-only.

## Module surface

```python
from lapse import (
    LapseAssumption,            # frozen dataclass
    combined_monthly_survival,  # mortality × lapse composition
    default_lapse_assumption,   # 8/7/6/5/4/3/2 ultimate 2% template
    lapse_decrement_from_csv,   # CSV loader (policy_year, q_w)
    monthly_mortality_q_from_annual,  # annual qx -> monthly qx_m
)
```

## Decrement model

* **Annual rate by policy year.** Inputs are stored as
  `tuple[float, ...]` indexed 0-based for policy year 1, 2, ...
  with an `ultimate_rate` for years past the table.
* **Within-year hazard.** Constant-force-of-decrement convention
  (matches the existing mortality module): `q_w_m = 1 - (1 - q_w)^(1/12)`.
* **Composition with mortality.** Independent decrements multiply:
  `S(t) = ∏_{s=0}^{t-1} (1 - q_x_m(s)) × (1 - q_w_m(s))`.
* **No interactions.** No dynamic lapse (rates do not depend on
  in-the-moneyness or interest rate environment), no surrender-charge
  recapture (deferred to v2), no per-cohort overrides.

## v1 limitations

* **Static only.** A future v2 may add dynamic-lapse callables; the v1
  surface is forward-compatible (callers continue to pass a
  `LapseAssumption | None`).
* **No surrender charge in pricing.** Surrender-charge schedules can be
  recorded on the contract dataclass for display purposes (e.g. MYGA
  schedule `(7, 6, 5, 4, 3, 0, 0, 0, 0, 0)`), but they are not priced
  into the single premium in v1.
* **Default is `None`.** Engines never apply lapse unless the caller
  explicitly passes a `LapseAssumption`. The
  `tests/test_lapse_default_is_no_op_for_<P>` regression tests assert
  this for every implemented product.

## Default assumption

`default_lapse_assumption()` returns:

| Policy year | Annual lapse rate |
|-------------|-------------------|
| 1           | 8% |
| 2           | 7% |
| 3           | 6% |
| 4           | 5% |
| 5           | 4% |
| 6           | 3% |
| 7           | 2% |
| 8+ (ultimate) | 2% |

This is a **generic placeholder template**. Production users override
with their own table via `lapse_decrement_from_csv(path)` or by
constructing `LapseAssumption` directly.

## CSV format

```csv
policy_year,q_w
1,0.10
2,0.08
3,0.06
,0.02
```

The blank `policy_year` row is the optional ultimate row. Years must
be 1-indexed and contiguous; missing intermediate years raise
`ValueError`.

## Existing-product invariant

SPIA / Term / RILA do NOT accept a `lapse=` parameter today and never
apply lapse. The existing 3 products' golden JSON files are byte-
identical after Phase 0; this is verified by the
`tests/parity/test_golden_modelcheck.py` ratchet.

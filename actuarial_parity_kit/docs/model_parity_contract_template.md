# Actuarial Model Parity Contract

**Version:** 1.0  
**Last updated:** [DATE]  
**Applies to:** [PRODUCT] [ENGINE] (Python ↔ Excel)

This document is the authoritative source of truth for numerical parity between the Python
calculation engine and the Excel workbook.

Any change to calculation logic must be reflected in **both** engines and must be tested
using the parity suite in `tests/parity/`.

---

## 1. Parity Definition

For any given set of inputs, both engines must agree to within the tolerances below at
**every step** of the projection horizon, for every tracked state variable.

| Variable | Absolute tolerance | Notes |
|----------|--------------------|-------|
| [Primary value] | 1e-4 | [Units] |
| [Secondary value] | 1e-4 | [Units] |
| [Index/factor] | 1e-10 | Dimensionless |
| [Add more rows as needed] | | |

> **Display rounding** is the only place where values may be rounded. All internal
> calculations use full float64 precision.

---

## 2. Ordering / Tie-Break Policy

When two or more items have equal or near-equal sort keys, the **lower-indexed item is
always processed first**. This is the universal rule for all selection/ordering logic.

### Python implementation
```python
order = np.argsort(
    sort_key
    + (depleted_condition) * SENTINEL    # push depleted items to end
    + np.arange(n) * EPSILON             # epsilon tie-break by index
)
```

### Excel implementation
```
adjusted_key[k] = sort_key[k] + (k+1)*EPSILON
min_key = MIN(IF(active[k], adjusted_key[k], SENTINEL))
select[k] = IF(ABS(adjusted_key[k] - min_key) < THRESHOLD, ...)
```

**Critical**: `THRESHOLD < EPSILON / 2` to prevent double-selection.

---

## 3. Epsilon Policy

| Parameter | Value | Rationale |
|-----------|-------|-----------|
| Per-item epsilon (Python) | [value] | Unique sort key per item |
| Per-item epsilon (Excel) | [value] | Item k=0 gets non-zero epsilon |
| Comparison threshold (Excel) | [value] | Must be < epsilon/2 |
| Depleted-item sentinel | [value] | Pushes zero items to end |

---

## 4. Floating-Point Accumulation Warning

Values computed by repeated addition/subtraction from a starting value WILL differ from
values that were reset to an exact reference, even when both should represent the same
quantity. The error is typically ~1e-16 but CAN flip comparison order.

**Rule**: Never use raw floating-point comparison for ordering. Always use epsilon-adjusted keys.

---

## 5. Yield Curve / Discount Factor Policy

<!-- Adapt for your product's discount mechanics -->
- Interpolation method: [log-linear / linear / other]
- Boundary handling: [flat extrapolation / other]
- Spread application: [additive / multiplicative / other]

---

## 6. Selection/Disinvestment Policy

<!-- Document your product's specific selection rules -->
- Rule: [shortest-first / pro-rata / other]
- Trigger condition: [when/how]
- Post-selection state reset: [description]

---

## 7. Known Risk Areas

| Area | Failure mode | Mitigation |
|------|-------------|------------|
| Selection tie-break | Double-selection when two items differ by ~1e-16 | THRESHOLD < EPSILON/2 |
| Accumulated sort keys | Wrong item selected | Epsilon tie-break by index |
| Stale Excel cache | data_only=True returns old values | Always fully recalculate |
| [Add product-specific risks] | | |

---

## 8. Change Control

Any change to this contract requires:
1. PR description explaining the change and why.
2. All `tests/parity/` tests passing.
3. Updated version number and date in this header.

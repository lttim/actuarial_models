# IUL Product Spec

**Status:** Mechanics-production prototype
**Engine:** `iul_projection.py`
**Workbook:** `build_iul_excel_workbook.py`

## Scope

Indexed Universal Life is modelled as a monthly permanent-life account-value
projection with annual point-to-point indexed segment crediting. Python owns the
deterministic engine and Monte Carlo aggregation. Excel workbooks independently
recalculate selected deterministic/sample paths from workbook inputs and editable
monthly schedule tables.

Included mechanics:

- Single premium plus scheduled flexible premiums through `LevelPremiumSchedule`.
- Premium load, monthly expense charge, COI, NAR, and account-value termination.
- Annual point-to-point indexed crediting with participation, cap, and floor.
- Level-face and return-of-account-value death benefit options.
- Scheduled withdrawals.
- Fixed policy loans: draws, repayments, monthly loan interest, loan balance,
  net death benefit, surrender charge, and surrender value.
- Static lapse placeholder support through the pricing engine interface.

Out of first scope:

- No-lapse guarantees.
- Overloan protection.
- Dynamic lapse calibration.
- Illustration compliance.
- Statutory valuation and reserves.
- Full Excel Monte Carlo.

## Monthly Mechanics

For each projection month:

1. Accrue loan interest on the beginning loan balance.
2. Add single premium in month 1 and scheduled flexible premiums, net of load.
3. Apply segment credit on segment anniversaries; otherwise credit is zero.
4. Compute gross death benefit and NAR.
5. Deduct COI and monthly policy expense.
6. Apply loan draws and repayments.
7. Pay scheduled withdrawals subject to available account value.
8. If account value reaches zero, mark the policy terminated.
9. Compute loan balance, net death benefit, surrender charge, and surrender value.

Cashflows in expected-benefit PV include death claims plus policy access
cashflows for withdrawals and loan draws. Expenses remain separated in the
monthly expense PV.

## Excel Contract

The workbook contains:

- `Inputs`: core product and economic assumptions.
- `PolicySchedules`: editable monthly planned premiums, withdrawals, loan
  draws, loan repayments, surrender charge rates, loan rate, and death benefit
  type.
- `IndexScenario`: deterministic/sample index path.
- `Liabilities`: Python audit columns plus formula-driven IUL mechanics columns
  for premium, withdrawal, loans, crediting, AV, death benefit, surrender value,
  and PV cashflows.
- `ModelCheck`: Python snapshot versus Excel formula summary.

The Excel formula path is for deterministic/sample audit only. Portfolio and
Monte Carlo aggregation remain Python-owned.

## Validation

Required development controls:

- Unit tests for crediting, scheduled premiums, withdrawals, loan balance
  roll-forward, surrender value, and net death benefit.
- Step-level invariants in `tests/test_iul_projection.py`.
- Static workbook formula validation through `tests/test_excel_export_validation.py`.
- Per-product ModelCheck parity through `tests/parity/test_excel_recalc_per_product.py`.
- Golden and SME-lite review evidence before re-baselining product goldens.

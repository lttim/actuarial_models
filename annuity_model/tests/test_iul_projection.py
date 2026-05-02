from __future__ import annotations

import numpy as np
import pytest

import iul_projection as iul
import pricing_projection as sp
from policy_features import (
    LevelPremiumSchedule,
    LoanTerms,
    MonthlySchedule,
    SurrenderChargeSchedule,
)


def _yc() -> sp.YieldCurve:
    return sp.YieldCurve.from_flat_rate(0.04)


def _mort_zero() -> sp.MortalityTableQx:
    ages = np.arange(0, 121, dtype=int)
    return sp.MortalityTableQx(ages, np.zeros_like(ages, dtype=float))


def test_iul_scheduled_premium_withdrawal_loan_and_surrender_state():
    contract = iul.IULContract(
        issue_age=45,
        sex="male",
        face_amount=250_000.0,
        single_premium=10_000.0,
        premium_load_pct=0.10,
        monthly_expense_charge=0.0,
        planned_premiums=LevelPremiumSchedule(
            modal_premium=1_000.0, mode_months=12, start_month=2, end_month=14
        ),
        withdrawals=MonthlySchedule((0.0,) * 11 + (500.0,)),
        loan_terms=LoanTerms(
            annual_rate=0.12,
            draws=MonthlySchedule((0.0,) * 5 + (2_000.0,)),
            repayments=MonthlySchedule((0.0,) * 11 + (100.0,)),
        ),
        surrender_charges=SurrenderChargeSchedule((0.07, 0.05)),
        cap=0.0,
        floor=0.0,
    )
    res = iul.price_iul_single_premium(
        contract=contract,
        yield_curve=_yc(),
        mortality=_mort_zero(),
        horizon_age=47,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        index_s0=100.0,
        index_levels_payment=np.full(24, 100.0),
    )
    assert res.premium_cashflows[0] == pytest.approx(10_000.0)
    assert res.premium_cashflows[1] == pytest.approx(1_000.0)
    assert res.loan_draw_cashflows[5] == pytest.approx(2_000.0)
    assert res.loan_balance_end_month[5] == pytest.approx(2_000.0)
    assert res.loan_interest_dollars[6] > 0.0
    assert res.loan_repayment_cashflows[11] == pytest.approx(100.0)
    assert res.withdrawal_cashflows[11] == pytest.approx(500.0)
    assert res.surrender_charge_dollars[11] > 0.0
    assert res.surrender_value_end_month[11] < res.account_value_end_month[11]
    assert res.net_death_benefit_end_month[11] == pytest.approx(
        max(0.0, res.db_end_month[11] - res.loan_balance_end_month[11])
    )


def test_iul_rejects_bad_index_levels():
    contract = iul.IULContract(issue_age=45, sex="male")
    with pytest.raises(ValueError, match="index levels"):
        iul.price_iul_single_premium(
            contract=contract,
            yield_curve=_yc(),
            mortality=_mort_zero(),
            horizon_age=46,
            expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
            index_s0=100.0,
            index_levels_payment=np.array([100.0, 0.0] + [100.0] * 10),
        )

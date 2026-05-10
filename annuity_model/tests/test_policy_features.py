from __future__ import annotations

import numpy as np
import pytest

from annuity_model.policy_features import (
    GLWBRider,
    LevelPremiumSchedule,
    LoanTerms,
    MonthlySchedule,
    SegmentAllocation,
    SurrenderChargeSchedule,
    buffer_credited_return,
    normalize_segment_allocations,
    segment_credited_return,
)


def test_monthly_schedule_pads_and_rejects_negative_values():
    sched = MonthlySchedule((10.0, 20.0))
    np.testing.assert_allclose(sched.values(4), np.array([10.0, 20.0, 0.0, 0.0]))
    with pytest.raises(ValueError, match="non-negative"):
        MonthlySchedule((1.0, -1.0)).values(2)


def test_level_premium_schedule_places_modal_premiums():
    sched = LevelPremiumSchedule(modal_premium=100.0, mode_months=3, start_month=2, end_month=8)
    np.testing.assert_allclose(
        sched.values(10),
        np.array([0.0, 100.0, 0.0, 0.0, 100.0, 0.0, 0.0, 100.0, 0.0, 0.0]),
    )


def test_surrender_charge_schedule_maps_policy_years_to_months():
    rates = SurrenderChargeSchedule((0.07, 0.05)).monthly_rates(30)
    assert np.all(rates[:12] == pytest.approx(0.07))
    assert np.all(rates[12:24] == pytest.approx(0.05))
    assert np.all(rates[24:] == pytest.approx(0.0))


def test_segment_allocations_normalize_weights_and_apply_buffer():
    allocs = normalize_segment_allocations(
        (
            SegmentAllocation(
                weight=25.0, design="cap_floor", participation=1.0, cap=0.10, floor=0.0
            ),
            SegmentAllocation(
                weight=75.0, design="buffer", participation=1.0, cap=0.12, buffer=0.10
            ),
        )
    )
    assert sum(a.weight for a in allocs) == pytest.approx(1.0)
    assert segment_credited_return(allocation=allocs[0], raw_index_return=0.20) == pytest.approx(
        0.10
    )
    assert segment_credited_return(allocation=allocs[1], raw_index_return=-0.15) == pytest.approx(
        -0.05
    )
    assert (
        buffer_credited_return(raw_index_return=-0.08, participation=1.0, cap=0.12, buffer=0.10)
        == 0.0
    )


def test_glwb_and_loan_terms_validate_bounds():
    GLWBRider(enabled=True, fee_annual=0.01, rollup_annual=0.05, withdrawal_rate=0.04)
    with pytest.raises(ValueError, match="withdrawal_rate"):
        GLWBRider(enabled=True, withdrawal_rate=1.5)
    loan = LoanTerms(annual_rate=0.06)
    assert loan.monthly_rate() > 0.0
    with pytest.raises(ValueError, match="annual_rate"):
        LoanTerms(annual_rate=-0.01).monthly_rate()

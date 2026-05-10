"""Tests for the static lapse / persistency framework."""

from __future__ import annotations

import csv
import tempfile

import numpy as np
import pytest

from annuity_model import lapse


def test_lapse_assumption_validates_rates():
    with pytest.raises(ValueError, match="must be in"):
        lapse.LapseAssumption(annual_lapse_rates_by_year=(0.05, 1.0))
    with pytest.raises(ValueError, match="ultimate_rate"):
        lapse.LapseAssumption(annual_lapse_rates_by_year=(0.05,), ultimate_rate=1.5)
    # Valid construction does not raise
    lapse.LapseAssumption(annual_lapse_rates_by_year=(0.0, 0.5, 0.99), ultimate_rate=0.0)


def test_annual_rate_lookup_handles_table_and_ultimate():
    la = lapse.LapseAssumption(annual_lapse_rates_by_year=(0.08, 0.07, 0.06), ultimate_rate=0.02)
    assert la.annual_rate_for_policy_year(1) == pytest.approx(0.08)
    assert la.annual_rate_for_policy_year(2) == pytest.approx(0.07)
    assert la.annual_rate_for_policy_year(3) == pytest.approx(0.06)
    # Ultimate kicks in past the table.
    assert la.annual_rate_for_policy_year(4) == pytest.approx(0.02)
    assert la.annual_rate_for_policy_year(99) == pytest.approx(0.02)
    with pytest.raises(ValueError):
        la.annual_rate_for_policy_year(0)


def test_monthly_decrements_constant_force_within_year():
    la = lapse.LapseAssumption(annual_lapse_rates_by_year=(0.12,), ultimate_rate=0.0)
    monthly = la.monthly_decrements(24)
    assert monthly.shape == (24,)
    expected_y1 = 1.0 - (1.0 - 0.12) ** (1.0 / 12.0)
    np.testing.assert_allclose(monthly[:12], np.full(12, expected_y1))
    # Year 2 falls back to ultimate (0.0).
    assert np.allclose(monthly[12:], 0.0)


def test_monthly_decrements_zero_when_no_table():
    la = lapse.LapseAssumption(annual_lapse_rates_by_year=(), ultimate_rate=0.0)
    assert np.allclose(la.monthly_decrements(36), 0.0)


def test_combined_monthly_survival_decreasing_and_bounded():
    qm = np.full(12, 0.001, dtype=float)
    qw = np.full(12, 0.005, dtype=float)
    surv = lapse.combined_monthly_survival(mortality_monthly_q=qm, lapse_monthly_q=qw)
    assert surv.shape == (12,)
    assert np.all(surv > 0.0) and np.all(surv <= 1.0)
    # Monotone non-increasing
    assert np.all(np.diff(surv) <= 0.0)
    # Composition matches manual cumulative product.
    expected = np.cumprod((1.0 - qm) * (1.0 - qw))
    np.testing.assert_allclose(surv, expected, rtol=0.0, atol=1e-15)


def test_combined_monthly_survival_validates_shape():
    with pytest.raises(ValueError):
        lapse.combined_monthly_survival(
            mortality_monthly_q=np.zeros(12), lapse_monthly_q=np.zeros(11)
        )


def test_default_lapse_assumption_matches_documented_table():
    la = lapse.default_lapse_assumption()
    assert la.annual_lapse_rates_by_year == (0.08, 0.07, 0.06, 0.05, 0.04, 0.03, 0.02)
    assert la.ultimate_rate == 0.02


def test_lapse_csv_round_trip():
    rows = [
        {"policy_year": "1", "q_w": "0.10"},
        {"policy_year": "2", "q_w": "0.08"},
        {"policy_year": "3", "q_w": "0.06"},
        {"policy_year": "", "q_w": "0.02"},  # ultimate
    ]
    with tempfile.NamedTemporaryFile(
        "w", suffix=".csv", delete=False, encoding="utf-8", newline=""
    ) as fh:
        writer = csv.DictWriter(fh, fieldnames=["policy_year", "q_w"])
        writer.writeheader()
        for r in rows:
            writer.writerow(r)
        path = fh.name
    la = lapse.lapse_decrement_from_csv(path)
    assert la.annual_lapse_rates_by_year == (0.10, 0.08, 0.06)
    assert la.ultimate_rate == pytest.approx(0.02)


def test_monthly_mortality_q_from_annual_round_trip():
    annual = np.array([0.0, 0.01, 0.05, 0.5], dtype=float)
    monthly = lapse.monthly_mortality_q_from_annual(annual)
    assert monthly.shape == annual.shape
    # round-trip back: annual = 1 - (1 - monthly)^12
    back = 1.0 - (1.0 - monthly) ** 12.0
    np.testing.assert_allclose(back, annual, rtol=0.0, atol=1e-12)

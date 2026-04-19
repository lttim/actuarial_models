"""Tests for the 2017 CSO Ultimate placeholder mortality table."""

from __future__ import annotations

import numpy as np
import pytest

import mortality_2017_cso as cso


@pytest.mark.parametrize(
    "sex,smoker_class",
    [
        ("male", "nonsmoker"),
        ("female", "nonsmoker"),
        ("male", "smoker"),
        ("female", "smoker"),
    ],
)
def test_loader_round_trips_each_cohort(sex, smoker_class):
    table = cso.MortalityTable2017CSO.load(sex=sex, smoker_class=smoker_class)
    assert table.sex == sex
    assert table.smoker_class == smoker_class
    assert table.ages.shape == (121,)
    assert table.qx.shape == (121,)
    assert int(table.ages[0]) == 0
    assert int(table.ages[-1]) == 120
    assert np.all((table.qx >= 0.0) & (table.qx <= 0.999))


def test_loader_validates_inputs():
    with pytest.raises(ValueError, match="sex"):
        cso.MortalityTable2017CSO.load(sex="other", smoker_class="nonsmoker")  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="smoker_class"):
        cso.MortalityTable2017CSO.load(sex="male", smoker_class="vape")  # type: ignore[arg-type]


def test_male_smoker_higher_than_male_nonsmoker():
    m_ns = cso.MortalityTable2017CSO.load(sex="male", smoker_class="nonsmoker")
    m_sk = cso.MortalityTable2017CSO.load(sex="male", smoker_class="smoker")
    # Smoker is higher across all working ages (older ages cap saturates).
    for age in (25, 35, 45, 55, 65, 75, 85):
        assert m_sk.qx_at_int_age(age) > m_ns.qx_at_int_age(age)


def test_female_lower_than_male_nonsmoker():
    m_ns = cso.MortalityTable2017CSO.load(sex="male", smoker_class="nonsmoker")
    f_ns = cso.MortalityTable2017CSO.load(sex="female", smoker_class="nonsmoker")
    for age in (25, 45, 65, 85):
        assert f_ns.qx_at_int_age(age) < m_ns.qx_at_int_age(age)


def test_qx_monotone_non_decreasing_after_age_30():
    """Past age 30 the synthetic Gompertz curve should rise monotonically."""
    table = cso.MortalityTable2017CSO.load(sex="male", smoker_class="nonsmoker")
    qx_above_30 = table.qx[30:]
    diffs = np.diff(qx_above_30)
    assert np.all(diffs >= -1e-9), "qx must be non-decreasing past age 30"


def test_monthly_survival_to_payment_returns_decreasing_array():
    table = cso.MortalityTable2017CSO.load(sex="male", smoker_class="nonsmoker")
    surv = table.monthly_survival_to_payment(issue_age=45, n_months=120)
    assert surv.shape == (120,)
    assert np.all((surv > 0.0) & (surv <= 1.0))
    assert np.all(np.diff(surv) <= 0.0), "survival must be non-increasing"
    # First-month survival is very close to 1 at age 45.
    assert surv[0] > 0.999


def test_monthly_survival_ignores_valuation_year_argument():
    table = cso.MortalityTable2017CSO.load(sex="male", smoker_class="nonsmoker")
    s1 = table.monthly_survival_to_payment(issue_age=45, n_months=24, valuation_year=2025)
    s2 = table.monthly_survival_to_payment(issue_age=45, n_months=24, valuation_year=2050)
    np.testing.assert_array_equal(s1, s2)


def test_artifact_path_resolves_to_existing_file():
    p = cso.cso_2017_artifact_path(sex="male", smoker_class="nonsmoker")
    assert p.is_file()


def test_alias_function_matches_static_method():
    a = cso.MortalityTable2017CSO.load(sex="female", smoker_class="smoker")
    b = cso.load_2017_cso_ultimate(sex="female", smoker_class="smoker")
    np.testing.assert_array_equal(a.ages, b.ages)
    np.testing.assert_array_equal(a.qx, b.qx)
    assert a.sex == b.sex == "female"
    assert a.smoker_class == b.smoker_class == "smoker"

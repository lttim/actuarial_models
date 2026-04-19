"""Tests for the shared UL/IUL/VUL account-value engine."""

from __future__ import annotations

import numpy as np
import pytest

import account_value as av


def test_avconfig_validates_inputs():
    with pytest.raises(ValueError, match="initial_premium"):
        av.AVConfig(
            initial_premium=-1.0, premium_load_pct=0.0, monthly_expense_charge=0.0,
            db_type="level_face", face_amount=100.0,
        )
    with pytest.raises(ValueError, match="premium_load_pct"):
        av.AVConfig(
            initial_premium=100.0, premium_load_pct=1.0, monthly_expense_charge=0.0,
            db_type="level_face", face_amount=100.0,
        )
    with pytest.raises(ValueError, match="monthly_expense_charge"):
        av.AVConfig(
            initial_premium=100.0, premium_load_pct=0.0, monthly_expense_charge=-1.0,
            db_type="level_face", face_amount=100.0,
        )
    with pytest.raises(ValueError, match="db_type"):
        av.AVConfig(
            initial_premium=100.0, premium_load_pct=0.0, monthly_expense_charge=0.0,
            db_type="invalid", face_amount=100.0,  # type: ignore[arg-type]
        )
    with pytest.raises(ValueError, match="face_amount"):
        av.AVConfig(
            initial_premium=100.0, premium_load_pct=0.0, monthly_expense_charge=0.0,
            db_type="level_face", face_amount=0.0,
        )


def test_evolve_av_zero_credit_zero_charge_preserves_premium():
    cfg = av.AVConfig(
        initial_premium=10_000.0,
        premium_load_pct=0.0,
        monthly_expense_charge=0.0,
        db_type="level_face",
        face_amount=10_000.0,
    )
    n = 12
    e = av.evolve_account_value(
        config=cfg,
        n_months=n,
        monthly_credit_rate=np.zeros(n),
        monthly_coi_q=np.zeros(n),
    )
    np.testing.assert_allclose(e.account_value_end_month, np.full(n, 10_000.0))
    np.testing.assert_allclose(e.coi_dollars, 0.0)
    # Type A: DB = max(face, AV) = face since AV == face.
    np.testing.assert_allclose(e.db_end_month, 10_000.0)
    # NAR = 0 because face == AV after credit.
    np.testing.assert_allclose(e.nar_end_month, 0.0)


def test_evolve_av_premium_load_applied_once_at_month_zero():
    cfg = av.AVConfig(
        initial_premium=100.0,
        premium_load_pct=0.20,
        monthly_expense_charge=0.0,
        db_type="level_face",
        face_amount=1_000.0,
    )
    n = 6
    e = av.evolve_account_value(
        config=cfg, n_months=n,
        monthly_credit_rate=np.zeros(n), monthly_coi_q=np.zeros(n),
    )
    # With NAR > 0 but qx=0, COI is 0 so AV stays constant after the load.
    expected_av = 100.0 * 0.80
    np.testing.assert_allclose(e.account_value_end_month, np.full(n, expected_av))


def test_evolve_av_coi_reduces_av_by_qx_times_nar():
    cfg = av.AVConfig(
        initial_premium=10_000.0,
        premium_load_pct=0.0,
        monthly_expense_charge=0.0,
        db_type="level_face",
        face_amount=100_000.0,
    )
    qx = np.full(1, 0.001)
    e = av.evolve_account_value(
        config=cfg, n_months=1,
        monthly_credit_rate=np.zeros(1), monthly_coi_q=qx,
    )
    expected_nar = 100_000.0 - 10_000.0
    expected_coi = 0.001 * expected_nar
    expected_av = 10_000.0 - expected_coi
    assert e.coi_dollars[0] == pytest.approx(expected_coi)
    assert e.nar_end_month[0] == pytest.approx(expected_nar)
    assert e.account_value_end_month[0] == pytest.approx(expected_av)


def test_evolve_av_terminates_when_av_runs_out():
    cfg = av.AVConfig(
        initial_premium=100.0,
        premium_load_pct=0.0,
        monthly_expense_charge=50.0,
        db_type="level_face",
        face_amount=1_000.0,
    )
    n = 6
    e = av.evolve_account_value(
        config=cfg, n_months=n,
        monthly_credit_rate=np.zeros(n), monthly_coi_q=np.zeros(n),
    )
    # Month 1: 100 - 50 = 50; Month 2: 50 - 50 = 0 -> terminate.
    assert e.account_value_end_month[0] == pytest.approx(50.0)
    assert e.account_value_end_month[1] == pytest.approx(0.0)
    assert bool(e.is_terminated_after_month[1])
    # All subsequent months stay terminated.
    assert np.all(e.is_terminated_after_month[1:])
    assert np.all(e.account_value_end_month[2:] == 0.0)
    assert np.all(e.coi_dollars[2:] == 0.0)


def test_evolve_av_av_never_negative():
    cfg = av.AVConfig(
        initial_premium=1_000.0,
        premium_load_pct=0.10,
        monthly_expense_charge=200.0,
        db_type="level_face",
        face_amount=10_000.0,
    )
    n = 24
    rng = np.random.default_rng(0)
    qx = rng.uniform(0.0, 0.05, size=n)
    cred = rng.uniform(-0.05, 0.05, size=n)
    e = av.evolve_account_value(
        config=cfg, n_months=n,
        monthly_credit_rate=cred, monthly_coi_q=qx,
    )
    assert np.all(e.account_value_end_month >= 0.0)


def test_evolve_av_validates_array_shapes():
    cfg = av.AVConfig(
        initial_premium=1_000.0, premium_load_pct=0.0, monthly_expense_charge=0.0,
        db_type="level_face", face_amount=10_000.0,
    )
    with pytest.raises(ValueError, match="monthly_credit_rate shape"):
        av.evolve_account_value(
            config=cfg, n_months=12,
            monthly_credit_rate=np.zeros(11), monthly_coi_q=np.zeros(12),
        )
    with pytest.raises(ValueError, match="monthly_coi_q shape"):
        av.evolve_account_value(
            config=cfg, n_months=12,
            monthly_credit_rate=np.zeros(12), monthly_coi_q=np.zeros(13),
        )


def test_evolve_av_zero_months_returns_empty_arrays():
    cfg = av.AVConfig(
        initial_premium=1_000.0, premium_load_pct=0.0, monthly_expense_charge=0.0,
        db_type="level_face", face_amount=10_000.0,
    )
    e = av.evolve_account_value(
        config=cfg, n_months=0,
        monthly_credit_rate=np.zeros(0), monthly_coi_q=np.zeros(0),
    )
    assert e.account_value_end_month.shape == (0,)
    assert e.coi_dollars.shape == (0,)

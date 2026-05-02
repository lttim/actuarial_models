"""Tests for the crediting-strategy framework + RILA back-compat."""

from __future__ import annotations

import math

import pytest

import crediting
import rila_projection as rp


def test_fixed_declared_rate_ignores_index():
    s = crediting.FixedDeclaredRate(annual_rate=0.04)
    assert s.credit_segment(raw_index_return=0.50) == pytest.approx(0.04)
    assert s.credit_segment(raw_index_return=-0.30) == pytest.approx(0.04)


def test_annual_p2p_capped_floor_and_cap_clamp():
    s = crediting.AnnualPointToPointCapped(participation=1.0, cap=0.10, floor=0.0)
    assert s.credit_segment(raw_index_return=0.20) == pytest.approx(0.10)
    assert s.credit_segment(raw_index_return=-0.20) == pytest.approx(0.0)
    assert s.credit_segment(raw_index_return=0.05) == pytest.approx(0.05)


def test_annual_p2p_participation_applied_first():
    s = crediting.AnnualPointToPointCapped(participation=0.5, cap=1.0, floor=-1.0)
    # Raw 0.20 * 0.5 = 0.10 within bounds.
    assert s.credit_segment(raw_index_return=0.20) == pytest.approx(0.10)


def test_annual_p2p_validates_bounds():
    with pytest.raises(ValueError, match="participation"):
        crediting.AnnualPointToPointCapped(participation=-0.1, cap=0.10, floor=0.0)
    with pytest.raises(ValueError, match="cap.*floor"):
        crediting.AnnualPointToPointCapped(participation=1.0, cap=0.05, floor=0.10)


def test_annual_p2p_buffer_absorbs_downside_before_loss():
    s = crediting.AnnualPointToPointBuffer(participation=1.0, cap=0.12, buffer=0.10)
    assert s.credit_segment(raw_index_return=0.20) == pytest.approx(0.12)
    assert s.credit_segment(raw_index_return=0.06) == pytest.approx(0.06)
    assert s.credit_segment(raw_index_return=-0.08) == pytest.approx(0.0)
    assert s.credit_segment(raw_index_return=-0.15) == pytest.approx(-0.05)


def test_annual_p2p_buffer_validates_bounds():
    with pytest.raises(ValueError, match="participation"):
        crediting.AnnualPointToPointBuffer(participation=-0.1, cap=0.10, buffer=0.10)
    with pytest.raises(ValueError, match="cap"):
        crediting.AnnualPointToPointBuffer(participation=1.0, cap=-0.01, buffer=0.10)
    with pytest.raises(ValueError, match="buffer"):
        crediting.AnnualPointToPointBuffer(participation=1.0, cap=0.10, buffer=1.10)


def test_segment_credited_return_from_strategy_round_trips():
    s = crediting.AnnualPointToPointCapped(participation=0.85, cap=0.09, floor=-0.02)
    # Identical to direct call.
    via_fn = crediting.segment_credited_return_from_strategy(
        strategy=s, raw_index_return=0.12
    )
    via_method = s.credit_segment(raw_index_return=0.12)
    assert via_fn == via_method


@pytest.mark.parametrize(
    "raw,participation,cap,floor",
    [
        (0.05, 1.0, 0.10, 0.0),
        (-0.10, 1.0, 0.10, 0.0),
        (0.30, 0.85, 0.09, -0.02),
        (-0.05, 0.85, 0.09, -0.02),
        (0.0, 1.0, 0.10, 0.0),
        (0.50, 0.5, 0.20, -0.10),
    ],
)
def test_rila_segment_back_compat_matches_strategy(raw, participation, cap, floor):
    """RILA's existing public ``segment_credited_return`` must be byte-
    identical to :class:`AnnualPointToPointCapped` (Section 1.2)."""
    direct = rp.segment_credited_return(
        raw=raw, participation=participation, cap=cap, floor=floor
    )
    strategy = crediting.AnnualPointToPointCapped(
        participation=participation, cap=cap, floor=floor
    )
    via_strategy = strategy.credit_segment(raw_index_return=raw)
    assert math.isclose(direct, via_strategy, rel_tol=0.0, abs_tol=0.0)


def test_fixed_declared_rate_collapses_when_strategy_used_in_segment_walk():
    """A degenerate AnnualPointToPointCapped with cap=floor=k acts like
    a FixedDeclaredRate(k)."""
    k = 0.04
    s_p2p = crediting.AnnualPointToPointCapped(participation=1.0, cap=k, floor=k)
    s_fixed = crediting.FixedDeclaredRate(annual_rate=k)
    for raw in (-0.50, -0.10, 0.0, 0.10, 0.50, 1.50):
        assert s_p2p.credit_segment(raw_index_return=raw) == pytest.approx(
            s_fixed.credit_segment(raw_index_return=raw)
        )

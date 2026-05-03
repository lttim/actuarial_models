"""Unit tests for :mod:`liability_aggregation`."""

from __future__ import annotations

import numpy as np
import pytest
from hypothesis import given
from hypothesis import strategies as st

import pricing_projection as sp
from liability_aggregation import (
    aggregate_by_product_type,
    aggregate_liability_paths,
    assert_rollups_sum_to_total,
    padded_cashflows_on_portfolio_grid,
)
from parity_constants import PORTFOLIO_ROLLUP_TOL
from product_registry import ProductType


def _path(n: int, scale: float) -> sp.LiabilityPath:
    dt = 1.0 / 12.0
    ty = np.arange(1, n + 1, dtype=float) * dt
    cf = np.full(n, scale, dtype=float)
    return sp.LiabilityPath(times_years=ty, expected_total_cashflows=cf)


def test_aggregate_single_path_round_trip() -> None:
    p = _path(5, 1.25)
    out = aggregate_liability_paths([p])
    np.testing.assert_allclose(out.times_years, p.times_years)
    np.testing.assert_allclose(out.expected_total_cashflows, p.expected_total_cashflows)


def test_aggregate_union_grid_sums() -> None:
    a = _path(3, 1.0)
    b = _path(5, 2.0)
    out = aggregate_liability_paths([a, b])
    assert len(out.expected_total_cashflows) == 5
    assert float(out.expected_total_cashflows[0]) == pytest.approx(3.0)
    assert float(out.expected_total_cashflows[2]) == pytest.approx(3.0)


def test_aggregate_by_product_type_and_invariant() -> None:
    typed = [
        (ProductType.SPIA, _path(4, 1.0)),
        (ProductType.SPIA, _path(4, 0.5)),
        (ProductType.TERM_LIFE, _path(4, 0.25)),
    ]
    by_t = aggregate_by_product_type(typed)
    total = aggregate_liability_paths([p for _, p in typed])
    assert set(by_t) == {ProductType.SPIA, ProductType.TERM_LIFE}
    assert_rollups_sum_to_total(
        rollups_by_product_type=by_t,
        portfolio=total,
        atol=PORTFOLIO_ROLLUP_TOL,
    )


@pytest.mark.property
@given(
    st.integers(min_value=3, max_value=36),
    st.floats(min_value=0.02, max_value=500.0, allow_nan=False),
    st.floats(min_value=0.02, max_value=500.0, allow_nan=False),
    st.floats(min_value=0.02, max_value=500.0, allow_nan=False),
)
def test_aggregate_partition_by_type_matches_total(
    n_months: int,
    scale_a: float,
    scale_b: float,
    scale_c: float,
) -> None:
    """Two SPIA-shaped paths + one other type still sum to the all-in aggregate."""
    types = (ProductType.SPIA, ProductType.SPIA, ProductType.TERM_LIFE)
    paths = [
        _path(n_months, float(scale_a)),
        _path(n_months, float(scale_b)),
        _path(n_months, float(scale_c)),
    ]
    typed = list(zip(types, paths, strict=True))
    total = aggregate_liability_paths(paths)
    by_t = aggregate_by_product_type(typed)
    assert_rollups_sum_to_total(
        rollups_by_product_type=by_t,
        portfolio=total,
        atol=PORTFOLIO_ROLLUP_TOL,
    )


def test_padded_cashflows_extends_with_zeros() -> None:
    short = _path(3, 2.0)
    out = padded_cashflows_on_portfolio_grid(short, 5)
    assert out.shape == (5,)
    np.testing.assert_allclose(out[:3], short.expected_total_cashflows)
    np.testing.assert_allclose(out[3:], 0.0)


def test_padded_cashflows_rejects_truncation() -> None:
    long = _path(6, 1.0)
    with pytest.raises(ValueError, match="exceeds target grid"):
        padded_cashflows_on_portfolio_grid(long, 4)


def test_aggregate_rejects_nonstandard_times_grid() -> None:
    bad = sp.LiabilityPath(
        times_years=np.array([0.1, 0.2]),
        expected_total_cashflows=np.array([1.0, 2.0]),
    )
    with pytest.raises(ValueError, match="standard monthly grid"):
        aggregate_liability_paths([bad])

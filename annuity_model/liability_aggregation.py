"""Aggregate :class:`pricing_projection.LiabilityPath` across policies and product types.

Used by the portfolio runner to build total-portfolio and per-:class:`product_registry.ProductType`
cashflow paths on a shared monthly grid (union of horizons, zero-padded).
"""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Iterable, Mapping, Sequence

import numpy as np

import pricing_projection as sp
from product_registry import ProductType


def _union_times_years(n_months: int) -> np.ndarray:
    dt = 1.0 / 12.0
    months = np.arange(1, n_months + 1, dtype=float)
    return months * dt


def _pad_cashflows(path: sp.LiabilityPath, n_months: int) -> np.ndarray:
    cf = np.asarray(path.expected_total_cashflows, dtype=float)
    if cf.ndim != 1:
        raise ValueError("expected_total_cashflows must be 1D.")
    if len(cf) > n_months:
        raise ValueError(
            f"path length {len(cf)} exceeds target grid {n_months}; "
            "cannot aggregate without truncation (refused)."
        )
    if len(cf) == 0:
        raise ValueError("empty liability path.")
    out = np.zeros(n_months, dtype=float)
    out[: len(cf)] = cf
    return out


def _validate_times_alignment(path: sp.LiabilityPath, n_months: int) -> None:
    ty = np.asarray(path.times_years, dtype=float)
    expected = _union_times_years(len(ty))
    if len(ty) != len(path.expected_total_cashflows):
        raise ValueError("times_years and expected_total_cashflows length mismatch.")
    if not np.allclose(ty, expected, rtol=0.0, atol=1e-12):
        raise ValueError(
            "LiabilityPath.times_years is not the standard monthly grid "
            "(k/12 for k=1..N); portfolio aggregation requires aligned grids."
        )


def aggregate_liability_paths(paths: Sequence[sp.LiabilityPath]) -> sp.LiabilityPath:
    """Sum cashflows onto the union monthly grid (max horizon among *paths*)."""
    if not paths:
        raise ValueError("aggregate_liability_paths requires at least one path.")
    n_months = max(len(np.asarray(p.expected_total_cashflows)) for p in paths)
    times = _union_times_years(n_months)
    total = np.zeros(n_months, dtype=float)
    for p in paths:
        _validate_times_alignment(p, len(np.asarray(p.expected_total_cashflows)))
        total += _pad_cashflows(p, n_months)
    return sp.LiabilityPath(times_years=times, expected_total_cashflows=total)


def aggregate_by_product_type(
    typed_paths: Iterable[tuple[ProductType, sp.LiabilityPath]],
) -> dict[ProductType, sp.LiabilityPath]:
    """Group paths by product type and aggregate within each group."""
    buckets: dict[ProductType, list[sp.LiabilityPath]] = defaultdict(list)
    for pt, path in typed_paths:
        buckets[pt].append(path)
    out: dict[ProductType, sp.LiabilityPath] = {}
    for pt in sorted(buckets, key=lambda x: x.value):
        out[pt] = aggregate_liability_paths(buckets[pt])
    return out


def padded_cashflows_on_portfolio_grid(
    path: sp.LiabilityPath, portfolio_n_months: int
) -> np.ndarray:
    """Return *path* expected cashflows left-aligned on a *portfolio_n_months* grid.

    Trailing months are zero. If *path* is longer than *portfolio_n_months*, raises
    (aggregation never truncates). Validates the standard monthly ``times_years`` grid.
    """
    _validate_times_alignment(path, len(np.asarray(path.expected_total_cashflows)))
    return _pad_cashflows(path, int(portfolio_n_months))


def assert_rollups_sum_to_total(
    *,
    rollups_by_product_type: Mapping[ProductType, sp.LiabilityPath],
    portfolio: sp.LiabilityPath,
    rtol: float = 0.0,
    atol: float = 1e-9,
) -> None:
    """Assert sum(by-type monthly CF) == portfolio monthly CF on the portfolio grid."""
    ty_p = np.asarray(portfolio.times_years, dtype=float)
    cf_p = np.asarray(portfolio.expected_total_cashflows, dtype=float)
    n = len(cf_p)
    summed = np.zeros(n, dtype=float)
    for _pt, path in sorted(rollups_by_product_type.items(), key=lambda x: x[0].value):
        _validate_times_alignment(path, len(np.asarray(path.expected_total_cashflows)))
        cf = _pad_cashflows(path, n)
        summed += cf
    if not np.allclose(summed, cf_p, rtol=rtol, atol=atol):
        diff = float(np.max(np.abs(summed - cf_p)))
        raise AssertionError(f"rollup sum != portfolio total: max_abs_diff={diff} (atol={atol}).")
    if not np.allclose(ty_p, _union_times_years(n), rtol=0.0, atol=1e-12):
        raise AssertionError("portfolio.times_years is not the standard monthly union grid.")


__all__ = [
    "aggregate_by_product_type",
    "aggregate_liability_paths",
    "assert_rollups_sum_to_total",
    "padded_cashflows_on_portfolio_grid",
]

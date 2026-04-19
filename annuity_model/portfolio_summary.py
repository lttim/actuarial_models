"""JSON-serializable summary dict for :class:`portfolio.PortfolioResult`."""

from __future__ import annotations

from typing import Any

from portfolio import PortfolioResult


def portfolio_result_to_summary_dict(res: PortfolioResult) -> dict[str, Any]:
    """Shape used by CLI output and integration goldens."""
    by_type: dict[str, Any] = {}
    for pt in sorted(res.rollups_by_product_type, key=lambda x: x.value):
        scal = res.product_type_scalar_rollups[pt]
        path = res.rollups_by_product_type[pt]
        by_type[pt.value] = {
            "policy_count": scal.policy_count,
            "sum_single_premium": scal.sum_single_premium,
            "sum_undiscounted_cashflows": scal.sum_undiscounted_cashflows,
            "rollup_cf_sum": float(path.expected_total_cashflows.sum()),
        }
    out: dict[str, Any] = {
        "n_policies": len(res.policy_results),
        "by_product_type": by_type,
        "total_cf_sum": float(res.liability_path_total.expected_total_cashflows.sum()),
    }
    if res.alm_result is not None:
        out["alm_duration_gap"] = float(res.alm_result.duration_gap)
    return out


__all__ = ["portfolio_result_to_summary_dict"]

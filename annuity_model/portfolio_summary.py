"""JSON-serializable summary dict for :class:`portfolio.PortfolioResult`."""

from __future__ import annotations

from typing import Any

import numpy as np

from portfolio import PortfolioResult


def _summary_float(value: float) -> float:
    """Stable JSON precision for cross-platform portfolio golden snapshots."""
    return round(float(value), 10)


def portfolio_result_to_summary_dict(res: PortfolioResult) -> dict[str, Any]:
    """Shape used by CLI output and integration goldens."""
    by_type: dict[str, Any] = {}
    for pt in sorted(res.rollups_by_product_type, key=lambda x: x.value):
        scal = res.product_type_scalar_rollups[pt]
        path = res.rollups_by_product_type[pt]
        by_type[pt.value] = {
            "policy_count": scal.policy_count,
            "sum_single_premium": _summary_float(scal.sum_single_premium),
            "sum_undiscounted_cashflows": _summary_float(scal.sum_undiscounted_cashflows),
            "rollup_cf_sum": _summary_float(path.expected_total_cashflows.sum()),
        }
    out: dict[str, Any] = {
        "n_policies": len(res.policy_results),
        "by_product_type": by_type,
        "total_cf_sum": _summary_float(res.liability_path_total.expected_total_cashflows.sum()),
    }
    if res.alm_result is not None:
        alm = res.alm_result
        fr = np.asarray(alm.funding_ratio, dtype=float)
        sur = np.asarray(alm.surplus, dtype=float)
        out["alm_duration_gap"] = _summary_float(alm.duration_gap)
        out["alm"] = {
            "duration_gap": _summary_float(alm.duration_gap),
            "duration_assets_mac": _summary_float(alm.duration_assets_mac),
            "duration_liabilities_mac": _summary_float(alm.duration_liabilities_mac),
            "pv01_net": _summary_float(alm.pv01_net),
            "pv01_assets": _summary_float(alm.pv01_assets),
            "pv01_liabilities": _summary_float(alm.pv01_liabilities),
            "funding_ratio_initial": _summary_float(fr[0]) if fr.size else None,
            "funding_ratio_min": _summary_float(np.nanmin(fr)) if fr.size else None,
            "surplus_min": _summary_float(np.min(sur)) if sur.size else None,
        }
    return out


__all__ = ["portfolio_result_to_summary_dict"]

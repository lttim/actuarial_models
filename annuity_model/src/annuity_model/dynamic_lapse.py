"""Dynamic lapse v2 demo helpers.

These helpers deliberately do not replace the existing static lapse framework.
They provide a transparent pricing-workbench overlay that can be promoted into
product engines once product-specific surrender/recapture rules are specified.
"""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np


@dataclass(frozen=True, slots=True)
class DynamicLapseConfig:
    base_annual_rate: float = 0.04
    rate_sensitivity: float = 2.0
    moneyness_sensitivity: float = 0.50
    floor: float = 0.005
    cap: float = 0.35


def dynamic_lapse_path(
    *,
    n_months: int,
    config: DynamicLapseConfig,
    rate_shock_bps: float = 0.0,
    moneyness: float = 0.0,
) -> np.ndarray:
    if n_months < 0:
        raise ValueError("n_months must be non-negative.")
    raw = (
        float(config.base_annual_rate)
        + float(config.rate_sensitivity) * float(rate_shock_bps) / 10000.0
        + float(config.moneyness_sensitivity) * max(float(moneyness), 0.0)
    )
    annual = float(np.clip(raw, float(config.floor), float(config.cap)))
    monthly = 1.0 - (1.0 - annual) ** (1.0 / 12.0)
    return np.full(n_months, monthly, dtype=float)


def persistency_from_monthly_lapse(lapse_monthly_q: np.ndarray) -> np.ndarray:
    q = np.clip(np.asarray(lapse_monthly_q, dtype=float), 0.0, 0.999999)
    return np.cumprod(1.0 - q)


__all__ = ["DynamicLapseConfig", "dynamic_lapse_path", "persistency_from_monthly_lapse"]

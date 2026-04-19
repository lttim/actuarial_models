"""Feature flags for the portfolio (multi-policy) runner."""

from __future__ import annotations

import os


def portfolio_v1_enabled() -> bool:
    """Return True when ``ANNUITY_MODEL_PORTFOLIO_V1=1`` is set in the environment."""
    return os.environ.get("ANNUITY_MODEL_PORTFOLIO_V1", "").strip() == "1"


__all__ = ["portfolio_v1_enabled"]

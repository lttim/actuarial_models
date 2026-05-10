"""Feature flags for the portfolio (multi-policy) runner and Streamlit sidebar."""

from __future__ import annotations

import os
from collections.abc import Mapping
from pathlib import Path


def _repo_root() -> Path:
    """Directory containing ``portfolio_config.py`` (the ``annuity_model/`` tree root)."""
    return Path(__file__).resolve().parent


def portfolio_disable_file_path() -> Path:
    """Path to the local opt-out marker (gitignored), same as ``run_pricing_ui.sh`` uses."""
    return _repo_root() / ".disable-portfolio-v1"


def portfolio_v1_enabled() -> bool:
    """Return whether portfolio UI/CLI features are enabled.

    Mirrors ``run_pricing_ui.sh`` / ``run_pricing_ui.bat`` so behavior does not depend
    on whether Streamlit was started via those launchers or via ``streamlit run``:

    1. If ``annuity_model/.disable-portfolio-v1`` exists → **disabled** (local opt-out
       wins over environment, same as the shell wrappers).
    2. Else if ``ANNUITY_MODEL_PORTFOLIO_V1`` is a truthy token (``1``, ``true``, …)
       → **enabled**.
    3. Else if it is a falsy token (``0``, ``false``, …) → **disabled**.
    4. Else (unset or empty) → **enabled** (local default ON; previously ``streamlit
       run`` alone left this unset and hid Portfolio unintentionally).
    """
    if portfolio_disable_file_path().is_file():
        return False
    raw = os.environ.get("ANNUITY_MODEL_PORTFOLIO_V1", "").strip().lower()
    if raw in ("1", "true", "yes", "on"):
        return True
    return raw not in ("0", "false", "no", "off")


def _session_ui_force_sidebar_key() -> str:
    """Lazy import so ``python -m cli`` does not load ``pricing_run_form_state`` (Streamlit)."""
    from annuity_model.pricing_run_form_state import PORTFOLIO_KEY

    return PORTFOLIO_KEY.UI_FORCE_SIDEBAR


def portfolio_sidebar_visible(session: Mapping[str, object] | None) -> bool:
    """Whether the Pricing Demo sidebar should list **Portfolio (multi-policy)**."""
    if portfolio_v1_enabled():
        return True
    if session is None:
        return False
    return bool(session.get(_session_ui_force_sidebar_key()))


def portfolio_disabled_explanation_markdown() -> str:
    """Human-readable reason ``portfolio_v1_enabled()`` is False (for Streamlit help)."""
    if portfolio_disable_file_path().is_file():
        p = portfolio_disable_file_path()
        return (
            f"**Opt-out file is present:** `{p}`  \n"
            "Remove that file (or rename it) to turn portfolio features back on, "
            "matching `./run_pricing_ui.sh` / `run_pricing_ui.bat` behavior."
        )
    raw = os.environ.get("ANNUITY_MODEL_PORTFOLIO_V1", "").strip()
    if raw.lower() in ("0", "false", "no", "off"):
        return (
            f"**Environment:** `ANNUITY_MODEL_PORTFOLIO_V1={raw!r}` disables portfolio.  \n"
            "Unset it or set `ANNUITY_MODEL_PORTFOLIO_V1=1` (or `true`) to enable."
        )
    return "**Portfolio is disabled** (unexpected configuration)."


__all__ = [
    "portfolio_disable_file_path",
    "portfolio_disabled_explanation_markdown",
    "portfolio_sidebar_visible",
    "portfolio_v1_enabled",
]

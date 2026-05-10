"""App-shell helpers extracted from the legacy Streamlit monolith."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import streamlit as st

from annuity_model.portfolio_config import (
    portfolio_disabled_explanation_markdown,
    portfolio_sidebar_visible,
    portfolio_v1_enabled,
)
from annuity_model.pricing_run_form_state import PORTFOLIO_KEY

from .navigation import section_label, section_options

APP_TITLE = "Pricing Demo"


def configure_pricing_page() -> None:
    """Apply the app-level Streamlit page configuration."""
    st.set_page_config(page_title=APP_TITLE, layout="wide")


def render_sidebar_shell(*, project_root: Path, session_state: Any | None = None) -> str:
    """Render the shared sidebar chrome and return the selected section key."""
    state = st.session_state if session_state is None else session_state

    st.title(APP_TITLE)
    if portfolio_v1_enabled():
        st.session_state.pop(PORTFOLIO_KEY.UI_FORCE_SIDEBAR, None)
        st.caption("Batch / multi-policy: set **Section** (below) to **Portfolio (multi-policy)**.")
    else:
        with st.expander("Portfolio section is off — why?", expanded=False):
            st.markdown(portfolio_disabled_explanation_markdown())
            st.caption(
                "Optional: show the Portfolio page in **Section** for this browser session "
                "only (Streamlit). CLI `portfolio-run` follows the same enablement rules."
            )
            st.checkbox(
                "Show Portfolio (multi-policy) in Section list",
                key=PORTFOLIO_KEY.UI_FORCE_SIDEBAR,
            )

    page = st.radio(
        "Section",
        options=section_options(include_portfolio=portfolio_sidebar_visible(state)),
        format_func=section_label,
    )
    st.divider()
    st.caption(f"Project root: `{project_root}`")
    return str(page)

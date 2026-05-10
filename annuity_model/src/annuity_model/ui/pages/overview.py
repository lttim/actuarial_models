"""Overview page for the pricing demo UI."""

from __future__ import annotations

import streamlit as st

from annuity_model.product_registry import (
    get_product_capabilities,
    product_label,
    product_options_for_ui,
)

from ..navigation import overview_section_labels
from ..widgets.product_badges import render_product_readiness_summary


def dynamic_overview_features() -> list[str]:
    """Build feature bullets from live product registry metadata."""
    options = list(product_options_for_ui())
    available_products = ", ".join(product_label(p) for p in options) if options else "None"
    mc_products = [
        product_label(p) for p in options if get_product_capabilities(p).supports_monte_carlo
    ]
    econ_products = [
        product_label(p) for p in options if get_product_capabilities(p).supports_economic_scenario
    ]
    return [
        f"Supported product run types: {available_products}.",
        "Run-time pricing dispatch is centralized in the product registry adapters.",
        f"Economic scenario controls enabled for: {', '.join(econ_products) if econ_products else 'None'}.",
        f"Monte Carlo pricing enabled for: {', '.join(mc_products) if mc_products else 'None'}.",
        "Yield curve sources: flat rate, zero-curve CSV, or par-yield CSV bootstrapped to zeros.",
        "Mortality sources are product-scoped and configured by registry defaults/options.",
        "ALM tab supports Treasury ladder projection, reinvestment/disinvestment policy controls, and KPI output tied to the active pricing run.",
        "What-if analysis provides before/after/impact views across pricing and ALM dimensions.",
        "Excel replicator export includes parity-oriented workbook output with optional MC and ALM snapshots.",
        "Embedded unit-test dashboard is available from the Unit Tests section.",
    ]


def render_overview() -> None:
    """Render the Overview page."""
    st.header("Model overview")
    st.markdown(
        "This workspace runs the pricing and projection engine with product adapters, "
        "scenario analysis, and Excel parity checks."
    )
    st.caption(
        "Overview content is generated from the product registry and shared section metadata "
        "to reduce documentation drift after model updates."
    )

    st.subheader("Current feature set")
    for i, feat in enumerate(dynamic_overview_features(), start=1):
        st.markdown(f"{i}. {feat}")

    render_product_readiness_summary(st)

    st.subheader("Workspace sections")
    st.markdown(
        "Use the sidebar to navigate: "
        + " | ".join(f"**{name}**" for name in overview_section_labels())
        + "."
    )

"""SPIA Streamlit AppTest smoke test.

Loads ``pricing_ui.py`` end-to-end via ``streamlit.testing.v1.AppTest``,
navigates to the Pricing Run page, selects SPIA, and asserts no
script-level exception was raised. Catches the class of bug where a
session-state key changes on the SPIA branch but no parity / engine
test exercises the UI form path.
"""

from __future__ import annotations

import pytest

from .conftest import (
    assert_no_exceptions,
    load_pricing_ui,
    navigate_to_pricing_run,
    select_product,
)


@pytest.mark.ui
def test_spia_pricing_run_page_renders_without_exceptions(streamlit_apptest_module) -> None:
    AppTest = streamlit_apptest_module
    at = load_pricing_ui(AppTest)

    navigate_to_pricing_run(at)
    assert_no_exceptions(at, context="navigate to Pricing Run page")

    select_product(at, "spia")
    assert_no_exceptions(at, context="select SPIA product")

    # The SPIA branch defines the spia-specific benefit-annual widget.
    # We don't assert on its value (default-seeded by build_run_form_seed_defaults);
    # we assert the widget *exists*, so a future refactor that drops the
    # widget without renaming the key fails this smoke instead of going
    # silent. Number-input keys are exposed via the AppTest number_input list.
    benefit_keys = [n.key for n in at.number_input if n.key == "run_spia_benefit_annual"]
    assert benefit_keys == ["run_spia_benefit_annual"], (
        "SPIA Pricing Run page is missing the 'run_spia_benefit_annual' "
        "number_input. Either the widget was removed/renamed without "
        "updating pricing_run_form_state.PRICING_RUN_NUMBER_INPUT_KEYS, "
        "or the SPIA branch failed to render."
    )

    # Cross-product widget that must always render irrespective of product.
    issue_age_keys = [n.key for n in at.number_input if n.key == "run_issue_age"]
    assert issue_age_keys, (
        "Pricing Run page is missing the cross-product 'run_issue_age' "
        "number_input -- the contract section did not render."
    )

"""Term Life Streamlit AppTest smoke test.

See ``test_apptest_spia.py`` for the rationale -- this is the Term Life
analog.
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
def test_term_pricing_run_page_renders_without_exceptions(streamlit_apptest_module) -> None:
    AppTest = streamlit_apptest_module
    at = load_pricing_ui(AppTest)

    navigate_to_pricing_run(at)
    assert_no_exceptions(at, context="navigate to Pricing Run page")

    select_product(at, "term_life")
    assert_no_exceptions(at, context="select Term Life product")

    # Term-specific widget keys -- catches the regression where the
    # term_years / premium_mode / benefit_timing widgets drop their
    # binding (the original P0 bug that prompted the hardening plan).
    expected_term_keys = {
        "run_term_benefit_annual",
        "run_term_monthly_premium",
    }
    actual_term_keys = {n.key for n in at.number_input}
    missing = expected_term_keys - actual_term_keys
    assert not missing, (
        f"Term Life Pricing Run page is missing number_input(s): {missing!r}. "
        "Either the widget was removed without updating "
        "pricing_run_form_state.PRICING_RUN_NUMBER_INPUT_KEYS, or the "
        "Term branch failed to render."
    )

    # Cross-product widget that must render regardless of product.
    issue_age_keys = [n.key for n in at.number_input if n.key == "run_issue_age"]
    assert issue_age_keys, (
        "Pricing Run page is missing the cross-product 'run_issue_age' "
        "number_input -- the contract section did not render."
    )

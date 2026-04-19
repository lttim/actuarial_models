"""Indexed UL Streamlit AppTest smoke test.

See ``test_apptest_spia.py`` for the rationale -- this is the Indexed UL
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
def test_iul_pricing_run_page_renders_without_exceptions(streamlit_apptest_module) -> None:
    AppTest = streamlit_apptest_module
    at = load_pricing_ui(AppTest)

    navigate_to_pricing_run(at)
    assert_no_exceptions(at, context="navigate to Pricing Run page")

    select_product(at, "indexed_ul")
    assert_no_exceptions(at, context="select Indexed UL product")

    expected_keys = {
        "run_iul_face_amount",
        "run_iul_single_premium",
        "run_iul_premium_load",
        "run_iul_monthly_expense",
        "run_iul_participation",
        "run_iul_cap",
        "run_iul_floor",
    }
    actual_keys = {n.key for n in at.number_input}
    missing = expected_keys - actual_keys
    assert not missing, (
        f"Indexed UL Pricing Run page is missing number_input(s): {missing!r}. "
        "Either the widget was removed without updating "
        "pricing_run_form_state.PRICING_RUN_NUMBER_INPUT_KEYS, or the "
        "Indexed UL branch failed to render."
    )

    issue_age_keys = [n.key for n in at.number_input if n.key == "run_issue_age"]
    assert issue_age_keys, (
        "Pricing Run page is missing the cross-product 'run_issue_age' "
        "number_input -- the contract section did not render."
    )

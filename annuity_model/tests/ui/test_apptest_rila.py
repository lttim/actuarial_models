"""RILA Streamlit AppTest smoke test.

See ``test_apptest_spia.py`` for the rationale -- this is the RILA
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
def test_rila_pricing_run_page_renders_without_exceptions(streamlit_apptest_module) -> None:
    AppTest = streamlit_apptest_module
    at = load_pricing_ui(AppTest)

    navigate_to_pricing_run(at)
    assert_no_exceptions(at, context="navigate to Pricing Run page")

    select_product(at, "rila")
    assert_no_exceptions(at, context="select RILA product")

    # RILA-specific knobs: participation, cap, floor, rider fee. These
    # keys are listed in pricing_run_form_state.PRICING_RUN_NUMBER_INPUT_KEYS;
    # if any disappear from the rendered form, the per-product UI bridge
    # has silently broken.
    expected_rila_keys = {
        "run_rila_participation",
        "run_rila_cap",
        "run_rila_floor",
        "run_rila_rider_fee",
    }
    actual_keys = {n.key for n in at.number_input}
    missing = expected_rila_keys - actual_keys
    assert not missing, (
        f"RILA Pricing Run page is missing number_input(s): {missing!r}. "
        "Either the widget was removed without updating "
        "pricing_run_form_state.PRICING_RUN_NUMBER_INPUT_KEYS, or the "
        "RILA branch failed to render."
    )

    issue_age_keys = [n.key for n in at.number_input if n.key == "run_issue_age"]
    assert issue_age_keys, (
        "Pricing Run page is missing the cross-product 'run_issue_age' "
        "number_input -- the contract section did not render."
    )

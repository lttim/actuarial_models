"""Portfolio manual-entry defaults stay aligned with Pricing Run seed defaults."""

from __future__ import annotations

import pandas as pd

from annuity_model.inforce_io import load_policy_inputs_from_csv_from_dataframe
from annuity_model.pricing_run_form_state import (
    RUN_KEY,
    build_run_form_seed_defaults,
    default_inforce_scratch_row,
)
from annuity_model.product_registry import ProductType, parse_term_length_label_to_years


def _seeds(pt: ProductType) -> dict:
    return build_run_form_seed_defaults(
        product_default=pt.value,
        saved_inputs={},
        meta={},
        default_product_type=pt,
    )


def test_default_inforce_spia_matches_pricing_seeds() -> None:
    seeds = _seeds(ProductType.SPIA)
    row = default_inforce_scratch_row(ProductType.SPIA)
    assert row["benefit_annual"] == seeds[RUN_KEY.SPIA_BENEFIT_ANNUAL]
    assert row["issue_age"] == seeds[RUN_KEY.ISSUE_AGE]
    assert row["sex"] == seeds[RUN_KEY.SEX]
    assert row["death_benefit"] is None


def test_default_inforce_term_matches_pricing_seeds() -> None:
    seeds = _seeds(ProductType.TERM_LIFE)
    row = default_inforce_scratch_row(ProductType.TERM_LIFE)
    assert row["death_benefit"] == seeds[RUN_KEY.TERM_BENEFIT_ANNUAL]
    assert row["monthly_premium"] == seeds[RUN_KEY.TERM_MONTHLY_PREMIUM]
    assert row["term_years"] == parse_term_length_label_to_years(str(seeds[RUN_KEY.TERM_LENGTH]))


def test_default_inforce_myga_and_va_match_pricing_seeds() -> None:
    seeds_myga = _seeds(ProductType.MYGA)
    row_myga = default_inforce_scratch_row(ProductType.MYGA)
    assert row_myga["single_premium"] == seeds_myga[RUN_KEY.MYGA_SINGLE_PREMIUM]
    assert row_myga["declared_rate_annual"] == seeds_myga[RUN_KEY.MYGA_DECLARED_RATE]
    assert row_myga["guarantee_years"] == seeds_myga[RUN_KEY.MYGA_GUARANTEE_YEARS]

    seeds_va = _seeds(ProductType.VARIABLE_ANNUITY)
    row_va = default_inforce_scratch_row(ProductType.VARIABLE_ANNUITY)
    assert row_va["single_premium"] == seeds_va[RUN_KEY.VA_SINGLE_PREMIUM]
    assert row_va["me_charge_annual"] == seeds_va[RUN_KEY.VA_ME_CHARGE]
    assert row_va["horizon_years"] == seeds_va[RUN_KEY.VA_HORIZON_YEARS]
    assert row_va["gmdb_basis"] == "return_of_premium"


def test_default_inforce_preserve_keeps_identity_across_product_switch() -> None:
    row = default_inforce_scratch_row(
        ProductType.SPIA,
        preserve={"policy_id": "Z9", "issue_age": 55, "sex": "female"},
    )
    assert row["policy_id"] == "Z9"
    assert row["issue_age"] == 55
    assert row["sex"] == "female"


def test_wide_row_round_trips_inforce_parser_for_each_product() -> None:
    """Assembled wide frame (like the manual UI) parses for every implemented product."""
    for pt in (
        ProductType.SPIA,
        ProductType.TERM_LIFE,
        ProductType.RILA,
        ProductType.MYGA,
        ProductType.FIA,
        ProductType.VARIABLE_ANNUITY,
        ProductType.WHOLE_LIFE,
        ProductType.UNIVERSAL_LIFE,
        ProductType.INDEXED_UL,
        ProductType.VARIABLE_UL,
    ):
        row = default_inforce_scratch_row(pt)
        row["policy_id"] = f"test-{pt.value}"
        df = pd.DataFrame([row])
        policies = load_policy_inputs_from_csv_from_dataframe(df)
        assert len(policies) == 1
        assert policies[0].product_type == pt

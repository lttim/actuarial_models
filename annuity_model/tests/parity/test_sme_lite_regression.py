"""SME lite regression -- one canonical scenario per implemented product.

This is the **always-on golden gate** the Actuary SME relies on. It is
deliberately tiny (one deterministic scenario per product, ~10 total)
so the gate completes in well under 30 seconds. Depth (sign / band /
sensitivity / closed-form) lives in the existing per-product
``test_<P>_actuarial.py`` files; this file is the top-line snapshot
the SME's evidence pack references.

Refreshing the baseline
-----------------------

After a deliberate methodology change, regenerate the golden JSON::

    UPDATE_GOLDEN_SME=1 pytest tests/parity/test_sme_lite_regression.py

Then commit the updated ``tests/parity/golden/sme/sme_baseline.json``
along with a paragraph in ``docs/model_change_log.md`` explaining why
each field's value moved.

Perf budget
-----------

A session-scoped autouse fixture asserts the whole module completes in
under 30 seconds (wall clock). Crossing the budget fails the suite
with a "prune scenarios or move Monte-Carlo out of the lite tier"
message. This keeps the gate cheap as the project grows.
"""

from __future__ import annotations

import json
import os
import time
from pathlib import Path
from typing import Any

import numpy as np
import pytest

from parity_constants import SME_LITE_MC_TOL, SME_LITE_TOL  # noqa: F401

# Imports for product engines. Kept inside the test functions where
# possible so a missing engine doesn't break collection of unrelated
# scenarios. Top-level imports only for shared types.
import pricing_projection as sp

GOLDEN_PATH = Path(__file__).parent / "golden" / "sme" / "sme_baseline.json"
PERF_BUDGET_SECONDS = 30.0

pytestmark = [pytest.mark.parity, pytest.mark.sme_smoke]


# ---------------------------------------------------------------------------
# Shared assumption helpers (kept tiny: yield curve + synthetic mortality).
# ---------------------------------------------------------------------------


def _flat_yc(rate: float = 0.04) -> sp.YieldCurve:
    return sp.YieldCurve.from_flat_rate(rate)


def _synthetic_mortality() -> sp.MortalityTableQx:
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.005 + ages * 1e-5, 1e-6, 0.4)
    return sp.MortalityTableQx(ages, qx)


def _zero_expenses() -> sp.ExpenseAssumptions:
    return sp.ExpenseAssumptions(0.0, 0.0, 0.0)


def _flat_index(n_months: int, level: float = 100.0) -> np.ndarray:
    return np.full(n_months, level, dtype=float)


# ---------------------------------------------------------------------------
# Per-product scenario builders. Each returns a top-line snapshot dict.
# Keep each scenario deterministic and short; this is the lite tier.
# ---------------------------------------------------------------------------


def _scenario_spia() -> dict[str, float]:
    contract = sp.SPIAContract(
        issue_age=65, sex="male", benefit_annual=6_000.0,
    )
    res = sp.price_spia_single_premium(
        contract=contract, yield_curve=_flat_yc(),
        mortality=_synthetic_mortality(),
        horizon_age=85, expenses=_zero_expenses(),
    )
    return {
        "pv_benefit": float(res.pv_benefit),
        "single_premium": float(res.single_premium),
        "expected_total_cashflows_first": float(res.expected_total_cashflows[0]),
        "expected_total_cashflows_last": float(res.expected_total_cashflows[-1]),
    }


def _scenario_term() -> dict[str, float]:
    import term_projection as tp

    contract = tp.TermLifeContract(
        issue_age=35, sex="male", death_benefit=250_000.0,
        monthly_premium=20.0, term_years=20,
    )
    res = tp.price_term_life_level_monthly(
        contract=contract, yield_curve=_flat_yc(),
        mortality=_synthetic_mortality(),
        horizon_age=55,
    )
    return {
        "pv_benefit": float(res.pv_benefit),
        "single_premium": float(res.single_premium),
        "expected_claim_cashflows_last": float(res.expected_claim_cashflows[-1]),
        "expected_premium_cashflows_last": float(res.expected_premium_cashflows[-1]),
    }


def _scenario_rila() -> dict[str, float]:
    import rila_projection as rila

    contract = rila.RILAContract(
        issue_age=60, sex="male",
        participation=0.8, cap=0.07, floor=0.0, rider_fee_annual=0.0,
    )
    n_months = (70 - 60) * 12
    # RILA's single_premium is computed implicitly from the expense PV
    # (see rila_projection.price_rila_single_premium line ~292: the SP
    # closes the equation `denom = 1 - rate - K`). Passing zero expenses
    # produces a degenerate SP=0 / AV=0 / PV=0 snapshot that any engine
    # change would still satisfy; the existing per-product RILA parity
    # test uses the same flat-fee convention shown here.
    res = rila.price_rila_single_premium(
        contract=contract, yield_curve=_flat_yc(),
        mortality=_synthetic_mortality(),
        horizon_age=70,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 25.0),
        index_s0=100.0, index_levels_payment=_flat_index(n_months),
    )
    return {
        "pv_benefit": float(res.pv_benefit),
        "single_premium": float(res.single_premium),
        "account_value_end_month_last": float(res.account_value_end_month[-1]),
    }


def _scenario_myga() -> dict[str, float]:
    import myga_projection as my

    contract = my.MYGAContract(
        issue_age=60, sex="male", single_premium=100_000.0,
        declared_rate_annual=0.045, guarantee_years=5,
    )
    res = my.price_myga_single_premium(
        contract=contract, yield_curve=_flat_yc(0.045),
        mortality=_synthetic_mortality(),
        horizon_age=70, expenses=_zero_expenses(),
    )
    return {
        "pv_benefit": float(res.pv_benefit),
        "single_premium": float(res.single_premium),
        "account_value_end_month_last": float(res.account_value_end_month[-1]),
    }


def _scenario_fia() -> dict[str, float]:
    import fia_projection as fp

    contract = fp.FIAContract(
        issue_age=60, sex="male", single_premium=100_000.0,
        participation=0.8, cap=0.07, floor=0.0, horizon_years=10,
    )
    n_months = 10 * 12
    res = fp.price_fia_single_premium(
        contract=contract, yield_curve=_flat_yc(),
        mortality=_synthetic_mortality(),
        horizon_age=70, expenses=_zero_expenses(),
        index_s0=100.0, index_levels_payment=_flat_index(n_months),
    )
    return {
        "pv_benefit": float(res.pv_benefit),
        "single_premium": float(res.single_premium),
        "account_value_end_month_last": float(res.account_value_end_month[-1]),
    }


def _scenario_va() -> dict[str, float]:
    import va_projection as va

    contract = va.VAContract(
        issue_age=55, sex="male", single_premium=100_000.0,
        me_charge_annual=0.014, horizon_years=20,
    )
    n_months = 20 * 12
    res = va.price_va_single_premium(
        contract=contract, yield_curve=_flat_yc(),
        mortality=_synthetic_mortality(),
        horizon_age=75, expenses=_zero_expenses(),
        index_s0=100.0, index_levels_payment=_flat_index(n_months),
    )
    return {
        "pv_benefit": float(res.pv_benefit),
        "single_premium": float(res.single_premium),
        "account_value_end_month_last": float(res.account_value_end_month[-1]),
    }


def _scenario_wl() -> dict[str, float]:
    import wl_projection as wl

    contract = wl.WLContract(
        issue_age=45, sex="male", smoker_class="nonsmoker", face_amount=250_000.0,
    )
    res = wl.price_wl_single_premium(
        contract=contract, yield_curve=_flat_yc(),
        mortality=_synthetic_mortality(),
        horizon_age=80, expenses=_zero_expenses(),
    )
    return {
        "pv_benefit": float(res.pv_benefit),
        "single_premium": float(res.single_premium),
    }


def _scenario_ul() -> dict[str, float]:
    import ul_projection as ul

    contract = ul.ULContract(
        issue_age=45, sex="male", face_amount=250_000.0, single_premium=25_000.0,
        premium_load_pct=0.06, monthly_expense_charge=7.50, declared_rate_annual=0.04,
    )
    res = ul.price_ul_single_premium(
        contract=contract, yield_curve=_flat_yc(),
        mortality=_synthetic_mortality(),
        horizon_age=80, expenses=_zero_expenses(),
    )
    return {
        "pv_benefit": float(res.pv_benefit),
        "single_premium": float(res.single_premium),
        "account_value_end_month_last": float(res.account_value_end_month[-1]),
    }


def _scenario_iul() -> dict[str, float]:
    import iul_projection as iul

    contract = iul.IULContract(
        issue_age=45, sex="male", face_amount=250_000.0, single_premium=25_000.0,
        premium_load_pct=0.06, monthly_expense_charge=7.50,
        participation=1.0, cap=0.10, floor=0.0,
    )
    n_months = (80 - 45) * 12
    res = iul.price_iul_single_premium(
        contract=contract, yield_curve=_flat_yc(),
        mortality=_synthetic_mortality(),
        horizon_age=80, expenses=_zero_expenses(),
        index_s0=100.0, index_levels_payment=_flat_index(n_months),
    )
    return {
        "pv_benefit": float(res.pv_benefit),
        "single_premium": float(res.single_premium),
        "account_value_end_month_last": float(res.account_value_end_month[-1]),
    }


def _scenario_vul() -> dict[str, float]:
    import vul_projection as vul

    contract = vul.VULContract(
        issue_age=45, sex="male", face_amount=250_000.0, single_premium=25_000.0,
        premium_load_pct=0.06, monthly_expense_charge=7.50,
    )
    n_months = (80 - 45) * 12
    res = vul.price_vul_single_premium(
        contract=contract, yield_curve=_flat_yc(),
        mortality=_synthetic_mortality(),
        horizon_age=80, expenses=_zero_expenses(),
        index_s0=100.0, index_levels_payment=_flat_index(n_months),
    )
    return {
        "pv_benefit": float(res.pv_benefit),
        "single_premium": float(res.single_premium),
        "account_value_end_month_last": float(res.account_value_end_month[-1]),
    }


def _scenario_portfolio() -> dict[str, float]:
    """Mixed-product portfolio (canonical inforce CSV + default_run_scenario)."""
    from pathlib import Path

    from inforce_io import load_policy_inputs_from_csv
    from portfolio import Portfolio
    from portfolio_runner import run_portfolio
    from portfolio_scenario import default_run_scenario
    from portfolio_summary import portfolio_result_to_summary_dict

    root = Path(__file__).resolve().parents[2]
    policies = load_policy_inputs_from_csv(root / "tests/data/inforce/example_v1/inforce.csv")
    res = run_portfolio(portfolio=Portfolio(policies=policies), scenario=default_run_scenario())
    s = portfolio_result_to_summary_dict(res)
    flat: dict[str, float] = {
        "n_policies": float(s["n_policies"]),
        "total_cf_sum": float(s["total_cf_sum"]),
    }
    for k, v in s["by_product_type"].items():
        flat[f"rollup_cf_sum_{k}"] = float(v["rollup_cf_sum"])
        flat[f"policy_count_{k}"] = float(v["policy_count"])
    return flat


SCENARIO_BUILDERS: dict[str, Any] = {
    "spia": _scenario_spia,
    "term": _scenario_term,
    "rila": _scenario_rila,
    "myga": _scenario_myga,
    "fia": _scenario_fia,
    "va": _scenario_va,
    "wl": _scenario_wl,
    "ul": _scenario_ul,
    "iul": _scenario_iul,
    "vul": _scenario_vul,
    "portfolio": _scenario_portfolio,
}

PRODUCTS = sorted(SCENARIO_BUILDERS.keys())


# ---------------------------------------------------------------------------
# Golden file IO
# ---------------------------------------------------------------------------


def _load_golden() -> dict[str, dict[str, float]]:
    if not GOLDEN_PATH.exists():
        return {}
    return json.loads(GOLDEN_PATH.read_text(encoding="utf-8"))


def _write_golden(golden: dict[str, dict[str, float]]) -> None:
    GOLDEN_PATH.parent.mkdir(parents=True, exist_ok=True)
    GOLDEN_PATH.write_text(
        json.dumps(golden, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )


# ---------------------------------------------------------------------------
# Perf budget enforcement (session-scoped, autouse).
# ---------------------------------------------------------------------------


@pytest.fixture(scope="module", autouse=True)
def _enforce_perf_budget():
    start = time.monotonic()
    yield
    elapsed = time.monotonic() - start
    if elapsed > PERF_BUDGET_SECONDS:
        pytest.fail(
            f"SME lite regression took {elapsed:.1f}s, exceeding the "
            f"{PERF_BUDGET_SECONDS:.0f}s budget. Prune scenarios or move "
            "Monte-Carlo cases out of the lite tier."
        )


# ---------------------------------------------------------------------------
# Parametrized test: one product per parameter.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("product_code", PRODUCTS)
def test_sme_lite_regression(product_code: str) -> None:
    snapshot = SCENARIO_BUILDERS[product_code]()
    update = os.environ.get("UPDATE_GOLDEN_SME") == "1"

    if update:
        golden = _load_golden()
        golden[product_code] = snapshot
        _write_golden(golden)
        # Don't assert in update mode -- the whole point is to refresh.
        return

    golden = _load_golden()
    assert product_code in golden, (
        f"No golden baseline for '{product_code}'. "
        f"Run `UPDATE_GOLDEN_SME=1 pytest {GOLDEN_PATH.parent.parent.name}/"
        f"test_sme_lite_regression.py` to seed it."
    )
    expected = golden[product_code]
    assert snapshot.keys() == expected.keys(), (
        f"Snapshot fields for '{product_code}' do not match the golden file. "
        f"Got {sorted(snapshot.keys())}, expected {sorted(expected.keys())}. "
        f"Refresh with UPDATE_GOLDEN_SME=1 if the schema change is intentional."
    )
    for field, expected_value in expected.items():
        actual_value = snapshot[field]
        diff = abs(actual_value - expected_value)
        assert diff <= SME_LITE_TOL, (
            f"SME lite regression: '{product_code}.{field}' = {actual_value:.6f} "
            f"vs golden {expected_value:.6f}, diff={diff:.6e} > {SME_LITE_TOL:.0e}. "
            f"Investigate the engine first; refresh the golden only with "
            f"UPDATE_GOLDEN_SME=1 after a documented methodology change "
            f"(see docs/model_change_log.md)."
        )

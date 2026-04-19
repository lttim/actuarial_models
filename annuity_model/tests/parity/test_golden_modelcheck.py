"""Golden ModelCheck snapshots: per-product, version-controlled, byte-exact.

What this test gives us
-----------------------

The platform already has parity tests that re-derive every metric from
inputs, prove Python ↔ Excel agreement, and run an Excel formula
validator. That stack catches *math* drift. It does not catch:

* a contract dataclass field being silently re-defaulted (e.g. ``spread=0``
  becoming ``spread=0.001``),
* a yield-curve helper rounding differently after a numpy version bump,
* an "innocent" refactor of an engine that produces the same shape of
  output but a different value at the canonical scenario,
* a per-product *display unit* change (factor format, currency symbol)
  that breaks downstream tooling but no math invariant.

This test does. It pins a small handful of canonical metrics per product
to a JSON file under ``tests/parity/golden/`` and asserts byte-exact
equality on every CI run. The constants are CODEOWNERS-protected (the
parent ``tests/parity/`` directory is already in CODEOWNERS), and a
deliberate update requires setting an environment variable
(``UPDATE_GOLDEN_MODELCHECK=1``) so it cannot happen by accident.

Workflow when a golden change is genuinely intended
---------------------------------------------------

1. Make the substantive change (engine fix, RP-2014 update, etc.) in a
   PR. The golden test will fail; that is the alarm.
2. In the same PR, run::

       UPDATE_GOLDEN_MODELCHECK=1 python -m pytest tests/parity/test_golden_modelcheck.py

   to refresh ``tests/parity/golden/<product>.json``.
3. Re-run the full parity gate (``python -m pytest tests/parity -q``)
   plus the four canonical gates from ``AGENTS.md``.
4. Update ``docs/model_change_log.md`` describing what changed and why
   the golden moved. CODEOWNERS routes the JSON change to the
   parity-critical reviewer.

Tolerance
---------

The atol used here is ``parity_constants.MODELCHECK_TOL`` (currently
``0.0``). That is intentional: ModelCheck reconciliation is the one
contract the platform promises stays bit-exact. Loosening this tolerance
is forbidden by ``tests/test_meta_invariants.py``.
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import numpy as np
import pytest

import parity_constants
import pricing_projection as sp
import rila_projection as rp
import term_projection as tp

GOLDEN_DIR = Path(__file__).parent / "golden"
UPDATE_ENV_VAR = "UPDATE_GOLDEN_MODELCHECK"

ATOL = float(parity_constants.MODELCHECK_TOL)  # currently 0.0; do NOT relax.


# ---------------------------------------------------------------------------
# Canonical scenario per product. These are intentionally NOT random and NOT
# derived from any user input; they are the published "Golden Scenario" for
# each product. Changing any of these constants is the same as changing the
# golden snapshot itself -- treat with the same review discipline.
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class GoldenSnapshot:
    product: str
    inputs_summary: dict[str, Any]
    metrics: dict[str, float]


def _flat_yc(rate: float) -> sp.YieldCurve:
    return sp.YieldCurve.from_flat_rate(rate)


def _synthetic_mortality(start: float, slope: float) -> sp.MortalityTableQx:
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(start + ages * slope, 1e-6, 0.4)
    return sp.MortalityTableQx(ages, qx)


def _spia_snapshot() -> GoldenSnapshot:
    contract = sp.SPIAContract(issue_age=65, sex="male", benefit_annual=100_000.0)
    yc = _flat_yc(0.04)
    mort = _synthetic_mortality(0.005, 1e-5)
    horizon_age = 80
    res = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=horizon_age,
        spread=0.0,
        valuation_year=None,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        expense_annual_inflation=0.0,
    )
    return GoldenSnapshot(
        product="spia",
        inputs_summary=dict(
            issue_age=65,
            sex="male",
            benefit_annual=100_000.0,
            yield_rate=0.04,
            horizon_age=horizon_age,
            mortality="synthetic_qx_0.005+1e-5*age",
            spread=0.0,
            expense_inflation=0.0,
        ),
        metrics={
            "pv_benefit": float(res.pv_benefit),
            "pv_monthly_expenses": float(res.pv_monthly_expenses),
            "single_premium": float(res.single_premium),
            "annuity_factor": float(res.annuity_factor),
            "n_months": int(res.months.shape[0]),
        },
    )


def _term_snapshot() -> GoldenSnapshot:
    contract = tp.TermLifeContract(
        issue_age=40,
        sex="male",
        death_benefit=250_000.0,
        monthly_premium=200.0,
        term_years=20,
        premium_mode="level_monthly",
        benefit_timing="eoy_death",
    )
    yc = _flat_yc(0.04)
    mort = _synthetic_mortality(0.005, 1e-5)
    horizon_age = 60
    res = tp.price_term_life_level_monthly(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=horizon_age,
        spread=0.0,
    )
    return GoldenSnapshot(
        product="term_life",
        inputs_summary=dict(
            issue_age=40,
            sex="male",
            death_benefit=250_000.0,
            monthly_premium=200.0,
            term_years=20,
            premium_mode="level_monthly",
            benefit_timing="eoy_death",
            yield_rate=0.04,
            horizon_age=horizon_age,
            mortality="synthetic_qx_0.005+1e-5*age",
            spread=0.0,
        ),
        metrics={
            "pv_benefit": float(res.pv_benefit),
            "pv_monthly_expenses": float(res.pv_monthly_expenses),
            "single_premium": float(res.single_premium),
            "annuity_factor": float(res.annuity_factor),
            "n_months": int(res.months.shape[0]),
        },
    )


def _rila_snapshot() -> GoldenSnapshot:
    contract = rp.RILAContract(
        issue_age=55,
        sex="male",
        participation=0.85,
        cap=0.09,
        floor=-0.02,
        rider_fee_annual=0.008,
    )
    yc = _flat_yc(0.035)
    mort = _synthetic_mortality(0.008, 2e-5)
    horizon_age = 85
    n_months = int(round((horizon_age - contract.issue_age) * 12))
    rng = np.random.default_rng(42)
    levels = 100.0 * np.cumprod(1.0 + rng.normal(0.004, 0.02, size=n_months))
    res = rp.price_rila_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=horizon_age,
        spread=0.0,
        valuation_year=None,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 25.0),
        index_s0=100.0,
        index_levels_payment=levels,
        expense_annual_inflation=0.01,
    )
    return GoldenSnapshot(
        product="rila",
        inputs_summary=dict(
            issue_age=55,
            sex="male",
            participation=0.85,
            cap=0.09,
            floor=-0.02,
            rider_fee_annual=0.008,
            yield_rate=0.035,
            horizon_age=horizon_age,
            mortality="synthetic_qx_0.008+2e-5*age",
            spread=0.0,
            expense_inflation=0.01,
            index_seed=42,
            index_drift=0.004,
            index_vol=0.02,
        ),
        metrics={
            "pv_benefit": float(res.pv_benefit),
            "pv_monthly_expenses": float(res.pv_monthly_expenses),
            "single_premium": float(res.single_premium),
            "annuity_factor": float(res.annuity_factor),
            "n_months": int(res.months.shape[0]),
        },
    )


SNAPSHOT_BUILDERS: dict[str, Callable[[], GoldenSnapshot]] = {
    "spia": _spia_snapshot,
    "term_life": _term_snapshot,
    "rila": _rila_snapshot,
}


# ---------------------------------------------------------------------------
# Test machinery.
# ---------------------------------------------------------------------------


def _golden_path(product: str) -> Path:
    return GOLDEN_DIR / f"{product}.json"


def _serialise(snap: GoldenSnapshot) -> str:
    """Stable JSON encoding so diffs in PRs are minimal."""
    payload = {
        "product": snap.product,
        "inputs_summary": snap.inputs_summary,
        "metrics": snap.metrics,
    }
    return json.dumps(payload, indent=2, sort_keys=True) + "\n"


def _load_golden(product: str) -> dict[str, Any]:
    path = _golden_path(product)
    if not path.exists():
        return {}
    return json.loads(path.read_text())


def _update_golden(snap: GoldenSnapshot) -> None:
    path = _golden_path(snap.product)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(_serialise(snap))


def _is_update_run() -> bool:
    return os.environ.get(UPDATE_ENV_VAR, "").strip() not in ("", "0", "false", "False")


@pytest.mark.parametrize("product", sorted(SNAPSHOT_BUILDERS), ids=sorted(SNAPSHOT_BUILDERS))
def test_golden_modelcheck_snapshot(product: str) -> None:
    """For each product, the freshly-computed snapshot must equal the golden file.

    On an UPDATE_GOLDEN_MODELCHECK=1 run, this test rewrites the golden
    file instead of asserting. The rewrite is intentionally placed here
    (and gated by an env var, NOT a CLI flag) so it cannot be triggered by
    `pytest -k snapshot --regen` or similar muscle-memory typos.
    """
    snap = SNAPSHOT_BUILDERS[product]()

    if _is_update_run():
        _update_golden(snap)
        pytest.skip(
            f"Updated golden file {_golden_path(product).name} "
            f"({UPDATE_ENV_VAR}=1)."
        )

    golden = _load_golden(product)
    assert golden, (
        f"No golden snapshot exists at {_golden_path(product)}. "
        f"Run `{UPDATE_ENV_VAR}=1 python -m pytest "
        f"tests/parity/test_golden_modelcheck.py -k {product}` once to "
        "create it, then commit the JSON file (CODEOWNERS will route the "
        "review to the parity-critical owner)."
    )

    # Inputs summary is part of the golden file precisely so a future
    # contributor who tweaks the canonical scenario realises they have
    # changed the input set, not just the engine.
    assert golden["product"] == snap.product
    assert golden["inputs_summary"] == snap.inputs_summary, (
        "Canonical inputs_summary drifted from the golden file; you have "
        "changed the *scenario*, not just the engine. If that was intended, "
        f"re-run with {UPDATE_ENV_VAR}=1 and update model_change_log.md."
    )

    # Metric-by-metric comparison so the assertion message points at the
    # exact metric that drifted.
    expected_metrics = golden["metrics"]
    actual_metrics = snap.metrics
    assert set(actual_metrics) == set(expected_metrics), (
        f"Metric set for {product!r} drifted. "
        f"Added: {sorted(set(actual_metrics) - set(expected_metrics))!r}, "
        f"removed: {sorted(set(expected_metrics) - set(actual_metrics))!r}."
    )
    failures: list[str] = []
    for key, expected in expected_metrics.items():
        actual = actual_metrics[key]
        if isinstance(expected, int) and isinstance(actual, int):
            if actual != expected:
                failures.append(f"{key}: golden={expected!r} actual={actual!r}")
            continue
        if abs(float(actual) - float(expected)) > ATOL:
            failures.append(
                f"{key}: golden={expected!r} actual={actual!r} "
                f"diff={float(actual) - float(expected):+.6e} "
                f"(atol={ATOL})"
            )
    assert not failures, (
        f"Golden ModelCheck drift for product={product!r}:\n"
        + "\n".join("  - " + f for f in failures)
        + f"\nIf this drift is intentional, re-run with {UPDATE_ENV_VAR}=1 "
        "to refresh the golden file AND update docs/model_change_log.md."
    )


def test_every_implemented_product_has_a_snapshot_builder() -> None:
    """Adding a new product must come with a golden snapshot, not a TODO."""
    from product_registry import implemented_product_types

    missing = [
        p.value for p in implemented_product_types() if p.value not in SNAPSHOT_BUILDERS
    ]
    assert not missing, (
        f"Products implemented but missing golden snapshot builders: {missing!r}. "
        "Add a `_<name>_snapshot()` function above and a corresponding entry "
        "in SNAPSHOT_BUILDERS, then run with "
        f"{UPDATE_ENV_VAR}=1 to seed the golden JSON file."
    )


def test_no_orphan_golden_files() -> None:
    """Stale golden files (e.g. for a removed product) must be cleaned up."""
    if not GOLDEN_DIR.exists():
        return
    on_disk = {p.stem for p in GOLDEN_DIR.glob("*.json")}
    expected = set(SNAPSHOT_BUILDERS)
    orphans = sorted(on_disk - expected)
    assert not orphans, (
        f"Found golden snapshot files with no matching SNAPSHOT_BUILDERS entry: "
        f"{orphans!r}. Either re-add the product (and its builder) or delete "
        "the stale JSON file."
    )

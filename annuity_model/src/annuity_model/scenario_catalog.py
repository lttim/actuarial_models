"""Named scenario catalog for pricing-demo workbench runs.

The catalog is intentionally small and deterministic: pricing teams need
repeatable stress names, not one-off slider states that disappear after a demo.
Each entry carries enough metadata to show purpose and replay the shock package
in diagnostics or a future durable ledger.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass


@dataclass(frozen=True, slots=True)
class PricingScenario:
    scenario_id: str
    label: str
    description: str
    owner: str
    intended_use: str
    rate_shift_bps: float = 0.0
    spread_shift_bps: float = 0.0
    longevity_improvement_pct: float = 0.0
    expense_multiplier: float = 1.0
    expense_inflation_shift_pct: float = 0.0
    equity_level_multiplier: float = 1.0
    mc_drift_shift_pct: float = 0.0
    mc_vol_multiplier: float = 1.0

    def to_dict(self) -> dict[str, object]:
        return asdict(self)


SCENARIO_CATALOG: tuple[PricingScenario, ...] = (
    PricingScenario(
        scenario_id="base",
        label="Base",
        description="Current pricing-run assumptions with no stress overlays.",
        owner="pricing_actuary",
        intended_use="pricing_base",
    ),
    PricingScenario(
        scenario_id="rates_up_100",
        label="+100 bps Rates",
        description="Parallel upward shift to the selected zero curve.",
        owner="pricing_actuary",
        intended_use="interest_rate_sensitivity",
        rate_shift_bps=100.0,
    ),
    PricingScenario(
        scenario_id="spread_widen_75",
        label="+75 bps Spread",
        description="Credit/liability spread widening overlay.",
        owner="pricing_actuary",
        intended_use="spread_sensitivity",
        spread_shift_bps=75.0,
    ),
    PricingScenario(
        scenario_id="longevity_plus_10",
        label="Longevity Stress",
        description="Mortality rates reduced 10 percent for longer-life exposure.",
        owner="model_risk_review",
        intended_use="mortality_sensitivity",
        longevity_improvement_pct=10.0,
    ),
    PricingScenario(
        scenario_id="equity_downturn",
        label="Equity Downturn",
        description="Indexed path levels down 20 percent, higher volatility, lower drift.",
        owner="capital_markets",
        intended_use="index_linked_stress",
        equity_level_multiplier=0.80,
        mc_drift_shift_pct=-4.0,
        mc_vol_multiplier=1.35,
    ),
    PricingScenario(
        scenario_id="expense_shock",
        label="Expense Shock",
        description="Maintenance and premium-expense load up 25 percent plus inflation add-on.",
        owner="expense_governance",
        intended_use="expense_sensitivity",
        expense_multiplier=1.25,
        expense_inflation_shift_pct=1.0,
    ),
)

_BY_ID: dict[str, PricingScenario] = {s.scenario_id: s for s in SCENARIO_CATALOG}


def list_pricing_scenarios() -> tuple[PricingScenario, ...]:
    return SCENARIO_CATALOG


def get_pricing_scenario(scenario_id: str) -> PricingScenario:
    try:
        return _BY_ID[scenario_id]
    except KeyError as exc:
        known = ", ".join(sorted(_BY_ID))
        raise KeyError(f"Unknown pricing scenario {scenario_id!r}. Known: {known}") from exc


__all__ = ["PricingScenario", "get_pricing_scenario", "list_pricing_scenarios"]

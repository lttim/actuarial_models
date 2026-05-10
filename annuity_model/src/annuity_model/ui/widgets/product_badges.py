"""Product maturity and assumption-evidence badges for the UI."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from annuity_model.products import get_product_definition, iter_product_definitions


@dataclass(frozen=True, slots=True)
class ProductStatus:
    """Display-oriented product readiness metadata."""

    product_type: str
    display_name: str
    maturity_label: str
    assumption_profile: str
    supports_economic_scenario: bool
    supports_monte_carlo: bool


def _product_value(product_type: Any) -> str:
    return str(getattr(product_type, "value", product_type))


def product_statuses() -> tuple[ProductStatus, ...]:
    """Return display-ready product status rows in canonical UI order."""
    rows: list[ProductStatus] = []
    for definition in sorted(iter_product_definitions(), key=lambda item: item.order):
        capabilities = definition.capabilities
        rows.append(
            ProductStatus(
                product_type=_product_value(definition.product_type),
                display_name=definition.display_name,
                maturity_label=definition.maturity_label,
                assumption_profile=definition.assumption_profile,
                supports_economic_scenario=bool(
                    getattr(capabilities, "supports_economic_scenario", False)
                ),
                supports_monte_carlo=bool(getattr(capabilities, "supports_monte_carlo", False)),
            )
        )
    return tuple(rows)


def product_status_for(product_type: Any) -> ProductStatus:
    """Return the status row for one product."""
    definition = get_product_definition(product_type)
    capabilities = definition.capabilities
    return ProductStatus(
        product_type=_product_value(definition.product_type),
        display_name=definition.display_name,
        maturity_label=definition.maturity_label,
        assumption_profile=definition.assumption_profile,
        supports_economic_scenario=bool(getattr(capabilities, "supports_economic_scenario", False)),
        supports_monte_carlo=bool(getattr(capabilities, "supports_monte_carlo", False)),
    )


def badges_for_status(status: ProductStatus) -> tuple[str, ...]:
    """Return short product-readiness badge labels."""
    econ_badge = "Economic scenarios" if status.supports_economic_scenario else "Deterministic"
    mc_badge = "Monte Carlo" if status.supports_monte_carlo else "No Monte Carlo"
    return (
        status.maturity_label,
        f"Assumptions: {status.assumption_profile}",
        econ_badge,
        mc_badge,
    )


def render_product_status_badges(st_mod: Any, product_type: Any) -> None:
    """Render readiness badges for the selected Pricing Run product."""
    status = product_status_for(product_type)
    st_mod.caption(" | ".join(badges_for_status(status)))


def render_product_readiness_summary(st_mod: Any) -> None:
    """Render a compact readiness inventory for the overview page."""
    st_mod.subheader("Product readiness")
    for status in product_statuses():
        st_mod.markdown(
            f"- **{status.display_name}**: " + " · ".join(badges_for_status(status))
        )

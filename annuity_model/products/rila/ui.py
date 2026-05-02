"""RILA adapter, metric formatter, UI config, and pricing-form helpers."""

from __future__ import annotations

from typing import Any

from pricing_run_form_state import RUN_KEY
from product_registry import (
    ProductType,
    _PRICING_METRIC_FORMATTERS,
    get_product_adapter,
    get_product_ui_config,
)

RILA_ADAPTER = get_product_adapter(ProductType.RILA)
rila_metric_formatter = _PRICING_METRIC_FORMATTERS[ProductType.RILA]
rila_ui_config = get_product_ui_config(ProductType.RILA)


def _monthly_amount_schedule(*, amount: float, start_month: int, end_month: int) -> tuple[float, ...]:
    if amount <= 0.0 or end_month < start_month:
        return ()
    months = max(int(end_month), int(start_month))
    out = [0.0] * months
    for month in range(max(1, int(start_month)), int(end_month) + 1):
        out[month - 1] = float(amount)
    return tuple(out)


def _surrender_rates(*, first_year_rate: float, years: int) -> tuple[float, ...]:
    if first_year_rate <= 0.0 or years <= 0:
        return ()
    n = int(years)
    return tuple(float(first_year_rate) * (n - i) / n for i in range(n))


def render_rila_pricing_controls(st_mod: Any, run_number_input_fn: Any) -> None:
    """Render product-specific RILA controls for ``pricing_ui``."""

    base, access, glwb = st_mod.tabs(["Core", "Access", "GLWB"])
    with base:
        r1, r2, r3, r4 = st_mod.columns(4)
        with r1:
            run_number_input_fn(
                "Participation",
                RUN_KEY.RILA_PARTICIPATION,
                default=1.0,
                min_value=0.0,
                max_value=5.0,
                format="%.4f",
            )
        with r2:
            run_number_input_fn(
                "Annual cap",
                RUN_KEY.RILA_CAP,
                default=0.10,
                min_value=-1.0,
                max_value=2.0,
                format="%.4f",
            )
        with r3:
            run_number_input_fn(
                "Annual floor",
                RUN_KEY.RILA_FLOOR,
                default=0.0,
                min_value=-1.0,
                max_value=1.0,
                format="%.4f",
            )
        with r4:
            run_number_input_fn(
                "Rider fee",
                RUN_KEY.RILA_RIDER_FEE,
                default=0.01,
                min_value=0.0,
                max_value=1.0,
                format="%.4f",
            )
        b1, b2, b3 = st_mod.columns(3)
        with b1:
            st_mod.selectbox(
                "Death benefit",
                options=["account_value", "return_of_premium"],
                key=RUN_KEY.RILA_DEATH_BENEFIT_TYPE,
            )
        with b2:
            run_number_input_fn(
                "Buffer allocation",
                RUN_KEY.RILA_BUFFER_WEIGHT,
                default=0.0,
                min_value=0.0,
                max_value=1.0,
                format="%.4f",
            )
        with b3:
            run_number_input_fn(
                "Buffer",
                RUN_KEY.RILA_BUFFER,
                default=0.10,
                min_value=0.0,
                max_value=1.0,
                format="%.4f",
            )
    with access:
        a1, a2, a3, a4 = st_mod.columns(4)
        with a1:
            run_number_input_fn(
                "Monthly withdrawal",
                RUN_KEY.RILA_WITHDRAWAL_AMOUNT,
                default=0.0,
                min_value=0.0,
                step=100.0,
            )
        with a2:
            run_number_input_fn("Start month", RUN_KEY.RILA_WITHDRAWAL_START, default=121, min_value=1, max_value=1200, step=1)
        with a3:
            run_number_input_fn("End month", RUN_KEY.RILA_WITHDRAWAL_END, default=240, min_value=1, max_value=1200, step=1)
        with a4:
            run_number_input_fn(
                "Surrender charge year 1",
                RUN_KEY.RILA_SURRENDER_CHARGE_Y1,
                default=0.07,
                min_value=0.0,
                max_value=1.0,
                format="%.4f",
            )
        run_number_input_fn("Surrender charge years", RUN_KEY.RILA_SURRENDER_CHARGE_YEARS, default=7, min_value=0, max_value=20, step=1)
    with glwb:
        st_mod.checkbox("Enable GLWB", key=RUN_KEY.RILA_GLWB_ENABLED, value=False)
        g1, g2, g3, g4 = st_mod.columns(4)
        with g1:
            run_number_input_fn("GLWB fee", RUN_KEY.RILA_GLWB_FEE, default=0.01, min_value=0.0, max_value=1.0, format="%.4f")
        with g2:
            run_number_input_fn("Roll-up", RUN_KEY.RILA_GLWB_ROLLUP, default=0.05, min_value=0.0, max_value=1.0, format="%.4f")
        with g3:
            run_number_input_fn("Income start month", RUN_KEY.RILA_GLWB_INCOME_START, default=121, min_value=1, max_value=1200, step=1)
        with g4:
            run_number_input_fn("Withdrawal rate", RUN_KEY.RILA_GLWB_WITHDRAWAL_RATE, default=0.05, min_value=0.0, max_value=1.0, format="%.4f")


def build_rila_contract_from_session(
    session_state: Any,
    *,
    issue_age: int,
    sex: str,
):
    import rila_projection as rp
    from policy_features import GLWBRider, MonthlySchedule, SegmentAllocation, SurrenderChargeSchedule

    buffer_weight = float(session_state.get(RUN_KEY.RILA_BUFFER_WEIGHT, 0.0))
    base_weight = max(0.0, 1.0 - buffer_weight)
    participation = float(session_state.get(RUN_KEY.RILA_PARTICIPATION, 1.0))
    cap = float(session_state.get(RUN_KEY.RILA_CAP, 0.10))
    floor = float(session_state.get(RUN_KEY.RILA_FLOOR, 0.0))
    allocations: list[SegmentAllocation] = []
    if base_weight > 0.0:
        allocations.append(SegmentAllocation(weight=base_weight, design="cap_floor", participation=participation, cap=cap, floor=floor))
    if buffer_weight > 0.0:
        allocations.append(
            SegmentAllocation(
                weight=buffer_weight,
                design="buffer",
                participation=participation,
                cap=max(0.0, cap),
                buffer=float(session_state.get(RUN_KEY.RILA_BUFFER, 0.10)),
            )
        )
    glwb = GLWBRider(
        enabled=bool(session_state.get(RUN_KEY.RILA_GLWB_ENABLED, False)),
        fee_annual=float(session_state.get(RUN_KEY.RILA_GLWB_FEE, 0.01)),
        rollup_annual=float(session_state.get(RUN_KEY.RILA_GLWB_ROLLUP, 0.05)),
        income_start_month=int(session_state.get(RUN_KEY.RILA_GLWB_INCOME_START, 121)),
        withdrawal_rate=float(session_state.get(RUN_KEY.RILA_GLWB_WITHDRAWAL_RATE, 0.05)),
    )
    return rp.RILAContract(
        issue_age=int(issue_age),
        sex="male" if sex == "male" else "female",
        participation=participation,
        cap=cap,
        floor=floor,
        rider_fee_annual=float(session_state.get(RUN_KEY.RILA_RIDER_FEE, 0.01)),
        segment_allocations=tuple(allocations),
        withdrawals=MonthlySchedule(
            _monthly_amount_schedule(
                amount=float(session_state.get(RUN_KEY.RILA_WITHDRAWAL_AMOUNT, 0.0)),
                start_month=int(session_state.get(RUN_KEY.RILA_WITHDRAWAL_START, 121)),
                end_month=int(session_state.get(RUN_KEY.RILA_WITHDRAWAL_END, 240)),
            )
        ),
        surrender_charges=SurrenderChargeSchedule(
            _surrender_rates(
                first_year_rate=float(session_state.get(RUN_KEY.RILA_SURRENDER_CHARGE_Y1, 0.07)),
                years=int(session_state.get(RUN_KEY.RILA_SURRENDER_CHARGE_YEARS, 7)),
            )
        ),
        death_benefit_type=str(session_state.get(RUN_KEY.RILA_DEATH_BENEFIT_TYPE, "account_value")),  # type: ignore[arg-type]
        glwb=glwb,
    )


__all__ = [
    "RILA_ADAPTER",
    "build_rila_contract_from_session",
    "rila_metric_formatter",
    "rila_ui_config",
    "render_rila_pricing_controls",
]

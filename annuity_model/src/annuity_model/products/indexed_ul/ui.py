"""INDEXED_UL adapter, metric formatter, UI config, and pricing-form helpers."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Literal, cast

from annuity_model.pricing_run_form_state import RUN_KEY
from annuity_model.product_registry import (
    _IUL_ADAPTER,
    ProductUIConfig,
    _life_single_premium_metrics,
)

INDEXED_UL_ADAPTER = _IUL_ADAPTER
indexed_ul_metric_formatter = _life_single_premium_metrics
indexed_ul_ui_config = ProductUIConfig(
    selected_info_message="Indexed UL (IUL): UL mechanics with annual point-to-point crediting on segment anniversaries.",
    projection_csv_filename="pricing_projection_indexed_ul.csv",
    recalc_workbook_filename="indexed_ul_recalc_model.xlsx",
)

if TYPE_CHECKING:
    from annuity_model.iul_projection import IULContract


def _monthly_amount_schedule(
    *, amount: float, start_month: int, end_month: int
) -> tuple[float, ...]:
    if amount <= 0.0 or end_month < start_month:
        return ()
    out = [0.0] * int(end_month)
    for month in range(max(1, int(start_month)), int(end_month) + 1):
        out[month - 1] = float(amount)
    return tuple(out)


def _surrender_rates(*, first_year_rate: float, years: int) -> tuple[float, ...]:
    if first_year_rate <= 0.0 or years <= 0:
        return ()
    n = int(years)
    return tuple(float(first_year_rate) * (n - i) / n for i in range(n))


def render_indexed_ul_pricing_controls(st_mod: Any, run_number_input_fn: Any) -> None:
    core, access, loan = st_mod.tabs(["Core", "Access", "Loans"])
    with core:
        i1, i2, i3, i4 = st_mod.columns(4)
        with i1:
            run_number_input_fn(
                "Face amount ($)",
                RUN_KEY.IUL_FACE_AMOUNT,
                default=250_000.0,
                min_value=1.0,
                step=10_000.0,
            )
            st_mod.selectbox(
                "Smoker class", options=["nonsmoker", "smoker"], key=RUN_KEY.IUL_SMOKER_CLASS
            )
        with i2:
            run_number_input_fn(
                "Single premium ($)",
                RUN_KEY.IUL_SINGLE_PREMIUM,
                default=25_000.0,
                min_value=1.0,
                step=1_000.0,
            )
            run_number_input_fn(
                "Premium load",
                RUN_KEY.IUL_PREMIUM_LOAD,
                default=0.06,
                min_value=0.0,
                max_value=0.5,
                format="%.4f",
            )
        with i3:
            run_number_input_fn(
                "Monthly expense",
                RUN_KEY.IUL_MONTHLY_EXPENSE,
                default=7.50,
                min_value=0.0,
                step=0.50,
            )
            run_number_input_fn(
                "Participation",
                RUN_KEY.IUL_PARTICIPATION,
                default=1.0,
                min_value=0.0,
                max_value=5.0,
                format="%.4f",
            )
        with i4:
            run_number_input_fn(
                "Annual cap",
                RUN_KEY.IUL_CAP,
                default=0.10,
                min_value=-1.0,
                max_value=2.0,
                format="%.4f",
            )
            run_number_input_fn(
                "Annual floor",
                RUN_KEY.IUL_FLOOR,
                default=0.0,
                min_value=-1.0,
                max_value=1.0,
                format="%.4f",
            )
        p1, p2, p3, p4 = st_mod.columns(4)
        with p1:
            st_mod.selectbox(
                "Death benefit",
                options=["level_face", "return_of_av"],
                key=RUN_KEY.IUL_DEATH_BENEFIT_TYPE,
            )
        with p2:
            run_number_input_fn(
                "Planned premium",
                RUN_KEY.IUL_PLANNED_PREMIUM,
                default=0.0,
                min_value=0.0,
                step=100.0,
            )
        with p3:
            run_number_input_fn(
                "Premium mode months",
                RUN_KEY.IUL_PREMIUM_MODE_MONTHS,
                default=12,
                min_value=1,
                max_value=12,
                step=1,
            )
        with p4:
            run_number_input_fn(
                "Premium end month",
                RUN_KEY.IUL_PREMIUM_END_MONTH,
                default=240,
                min_value=1,
                max_value=1200,
                step=1,
            )
    with access:
        a1, a2, a3, a4 = st_mod.columns(4)
        with a1:
            run_number_input_fn(
                "Monthly withdrawal",
                RUN_KEY.IUL_WITHDRAWAL_AMOUNT,
                default=0.0,
                min_value=0.0,
                step=100.0,
            )
        with a2:
            run_number_input_fn(
                "Start month",
                RUN_KEY.IUL_WITHDRAWAL_START,
                default=121,
                min_value=1,
                max_value=1200,
                step=1,
            )
        with a3:
            run_number_input_fn(
                "End month",
                RUN_KEY.IUL_WITHDRAWAL_END,
                default=240,
                min_value=1,
                max_value=1200,
                step=1,
            )
        with a4:
            run_number_input_fn(
                "Surrender charge year 1",
                RUN_KEY.IUL_SURRENDER_CHARGE_Y1,
                default=0.07,
                min_value=0.0,
                max_value=1.0,
                format="%.4f",
            )
        run_number_input_fn(
            "Surrender charge years",
            RUN_KEY.IUL_SURRENDER_CHARGE_YEARS,
            default=7,
            min_value=0,
            max_value=20,
            step=1,
        )
    with loan:
        l1, l2, l3, l4 = st_mod.columns(4)
        with l1:
            run_number_input_fn(
                "Loan annual rate",
                RUN_KEY.IUL_LOAN_RATE,
                default=0.05,
                min_value=0.0,
                max_value=1.0,
                format="%.4f",
            )
        with l2:
            run_number_input_fn(
                "Monthly loan draw", RUN_KEY.IUL_LOAN_DRAW, default=0.0, min_value=0.0, step=100.0
            )
        with l3:
            run_number_input_fn(
                "Loan draw start",
                RUN_KEY.IUL_LOAN_DRAW_START,
                default=121,
                min_value=1,
                max_value=1200,
                step=1,
            )
        with l4:
            run_number_input_fn(
                "Monthly loan repayment",
                RUN_KEY.IUL_LOAN_REPAY,
                default=0.0,
                min_value=0.0,
                step=100.0,
            )


def build_indexed_ul_contract_from_session(
    session_state: Any,
    *,
    issue_age: int,
    sex: str,
) -> IULContract:
    from annuity_model import iul_projection as iul_proj
    from annuity_model.policy_features import (
        LevelPremiumSchedule,
        LoanTerms,
        MonthlySchedule,
        SurrenderChargeSchedule,
    )

    loan_draw_start = int(session_state.get(RUN_KEY.IUL_LOAN_DRAW_START, 121))
    loan_draw_end = int(session_state.get(RUN_KEY.IUL_LOAN_DRAW_END, loan_draw_start + 119))
    smoker_class = cast(
        Literal["nonsmoker", "smoker"],
        str(session_state.get(RUN_KEY.IUL_SMOKER_CLASS, "nonsmoker")),
    )
    db_type = cast(
        Literal["return_of_av", "level_face"],
        str(session_state.get(RUN_KEY.IUL_DEATH_BENEFIT_TYPE, "level_face")),
    )
    return iul_proj.IULContract(
        issue_age=int(issue_age),
        sex="male" if sex == "male" else "female",
        smoker_class=smoker_class,
        face_amount=float(session_state.get(RUN_KEY.IUL_FACE_AMOUNT, 250_000.0)),
        single_premium=float(session_state.get(RUN_KEY.IUL_SINGLE_PREMIUM, 25_000.0)),
        premium_load_pct=float(session_state.get(RUN_KEY.IUL_PREMIUM_LOAD, 0.06)),
        monthly_expense_charge=float(session_state.get(RUN_KEY.IUL_MONTHLY_EXPENSE, 7.50)),
        planned_premiums=LevelPremiumSchedule(
            modal_premium=float(session_state.get(RUN_KEY.IUL_PLANNED_PREMIUM, 0.0)),
            mode_months=int(session_state.get(RUN_KEY.IUL_PREMIUM_MODE_MONTHS, 12)),
            start_month=1,
            end_month=int(session_state.get(RUN_KEY.IUL_PREMIUM_END_MONTH, 240)),
        ),
        withdrawals=MonthlySchedule(
            _monthly_amount_schedule(
                amount=float(session_state.get(RUN_KEY.IUL_WITHDRAWAL_AMOUNT, 0.0)),
                start_month=int(session_state.get(RUN_KEY.IUL_WITHDRAWAL_START, 121)),
                end_month=int(session_state.get(RUN_KEY.IUL_WITHDRAWAL_END, 240)),
            )
        ),
        loan_terms=LoanTerms(
            annual_rate=float(session_state.get(RUN_KEY.IUL_LOAN_RATE, 0.05)),
            draws=MonthlySchedule(
                _monthly_amount_schedule(
                    amount=float(session_state.get(RUN_KEY.IUL_LOAN_DRAW, 0.0)),
                    start_month=loan_draw_start,
                    end_month=loan_draw_end,
                )
            ),
            repayments=MonthlySchedule(
                _monthly_amount_schedule(
                    amount=float(session_state.get(RUN_KEY.IUL_LOAN_REPAY, 0.0)),
                    start_month=int(
                        session_state.get(RUN_KEY.IUL_LOAN_REPAY_START, loan_draw_end + 1)
                    ),
                    end_month=int(
                        session_state.get(RUN_KEY.IUL_LOAN_REPAY_END, loan_draw_end + 120)
                    ),
                )
            ),
        ),
        surrender_charges=SurrenderChargeSchedule(
            _surrender_rates(
                first_year_rate=float(session_state.get(RUN_KEY.IUL_SURRENDER_CHARGE_Y1, 0.07)),
                years=int(session_state.get(RUN_KEY.IUL_SURRENDER_CHARGE_YEARS, 7)),
            )
        ),
        participation=float(session_state.get(RUN_KEY.IUL_PARTICIPATION, 1.0)),
        cap=float(session_state.get(RUN_KEY.IUL_CAP, 0.10)),
        floor=float(session_state.get(RUN_KEY.IUL_FLOOR, 0.0)),
        db_type=db_type,
    )


__all__ = [
    "INDEXED_UL_ADAPTER",
    "build_indexed_ul_contract_from_session",
    "indexed_ul_metric_formatter",
    "indexed_ul_ui_config",
    "render_indexed_ul_pricing_controls",
]

"""
Unified Streamlit workspace for the SPIA model: overview, configurable pricing run,
interactive charts, and embedded unit-test dashboard.

Run from the annuity_model folder:
    streamlit run spia_ui.py
Or: run_spia_ui.bat (Windows).
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Literal

os.environ.setdefault("MPLBACKEND", "Agg")

import numpy as np
import pandas as pd
import streamlit as st

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import spia_projection as sp
from build_spia_excel_workbook import ExcelPythonSnapshot, build_workbook_from_spec, excel_spec_from_launcher
from test_dashboard import render_unit_tests_page


def _resolve_path(p: str) -> Path:
    path = Path(p.strip())
    if path.is_absolute():
        return path
    return (ROOT / path).resolve()


def _minimal_mortality() -> sp.MortalityTableQx:
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.005 + ages * 1e-5, 1e-6, 0.4)
    return sp.MortalityTableQx(ages, qx)


MortalityMode = Literal["synthetic", "qx_csv", "rp2014_mp2016"]
YieldMode = Literal["flat", "zero_csv", "par_bootstrap"]
ExpenseMode = Literal["csv", "manual"]


def _build_yield_curve(
    mode: YieldMode,
    *,
    flat_rate: float,
    zero_csv: str,
    par_csv: str,
    coupon_freq: int,
) -> sp.YieldCurve:
    if mode == "flat":
        return sp.YieldCurve.from_flat_rate(float(flat_rate))
    if mode == "zero_csv":
        return sp.YieldCurve.load_zero_curve_csv(str(_resolve_path(zero_csv)))
    return sp.YieldCurve.load_par_yield_csv_and_bootstrap(
        str(_resolve_path(par_csv)),
        coupon_freq=int(coupon_freq),
    )


def _build_mortality(
    mode: MortalityMode,
    *,
    qx_csv: str,
    rp_xlsx: str,
    rp_out_csv: str,
    mp_xlsx: str,
    mp_out_csv: str,
) -> tuple[sp.MortalityTableQx | sp.MortalityTableRP2014MP2016, bool]:
    """
    Returns (mortality, needs_valuation_year).
    """
    if mode == "synthetic":
        return _minimal_mortality(), False
    if mode == "qx_csv":
        return sp.MortalityTableQx.load_qx_csv(str(_resolve_path(qx_csv))), False
    base_qx = sp.ensure_rp2014_male_healthy_annuitant_qx_csv(
        rp2014_xlsx_path=str(_resolve_path(rp_xlsx)),
        out_csv_path=str(_resolve_path(rp_out_csv)),
    )
    mp_ages, mp_years, mp_i = sp.ensure_mp2016_male_improvement_csv(
        mp2016_xlsx_path=str(_resolve_path(mp_xlsx)),
        out_csv_path=str(_resolve_path(mp_out_csv)),
    )
    mortality = sp.MortalityTableRP2014MP2016(
        base_qx_2014=base_qx,
        mp2016_ages=mp_ages,
        mp2016_years=mp_years,
        mp2016_i_matrix=mp_i,
        base_year=2014,
    )
    return mortality, True


def _render_overview() -> None:
    st.header("Model overview")
    st.markdown(
        """
This workspace drives **`spia_projection.py`**: a single-life SPIA with **monthly**
benefits (nominal amounts can grow by **S&P proxy return indexation** from a CSV of
monthly index levels), **end-of-period** payments, and cashflows while the annuitant is alive.
**Maintenance expenses** can grow at a separate **fixed annual inflation** rate (monthly compounding).

### Main pieces

1. **`SPIAContract`** — Issue age, sex (metadata for now), annual benefit (starting point for
   the first month’s accrual; further months follow the index scenario), monthly frequency only.

2. **`YieldCurve`** — Continuously compounded zero rates `z(t)`; discount factors
   `DF(t) = exp(-(z(t) + spread) × t)` with log-linear interpolation on DFs inside the
   curve and flat-rate extrapolation beyond endpoints.

3. **Mortality** — Either a static **`MortalityTableQx`** (annual `q_x` by integer age) or
   **`MortalityTableRP2014MP2016`** (RP-2014 base qx in 2014 plus MP-2016 calendar-year
   improvements). RP+MP requires a **valuation year** so month-by-month calendar years are defined.

4. **`ExpenseAssumptions`** — Policy expense at issue, premium expense as a fraction of
   premium (solved iteratively via the closed form), and level monthly expenses while alive.

5. **`price_spia_single_premium`** — Monthly grid to `horizon_age`, survival, discount factors,
   **annuity factor** (level-$1 survival-weighted sum), PV of **indexed** benefits and inflated
   expenses, **single premium**, expected cashflows, **index return series** for charts, and reserves.

### Outputs you can inspect after a run

- Summary metrics (premium, PV benefit, annuity factor, etc.).
- Per-month table: age, survival, discount factor, expected benefit/expense/total, PV increment.
- Charts aligned with `illustrate_spia_projection.py` (survival, PV contributions, cumulative PV,
  nominal flows, reserves).

Use the sidebar to switch to **Run & results** or **Unit tests**.
        """
    )


def _result_dataframe(res: sp.SPIAProjectionResult, contract: sp.SPIAContract) -> pd.DataFrame:
    expected_payment_pv = res.expected_benefit_cashflows * res.discount_factors
    cumulative_pv = np.cumsum(expected_payment_pv)
    return pd.DataFrame(
        {
            "month": res.months,
            "time_years": res.times_years,
            "age_at_payment": res.ages_at_payment,
            "survival": res.survival_to_payment,
            "discount_factor": res.discount_factors,
            "index_level": res.index_level_at_payment,
            "index_simple_return": res.index_simple_return,
            "index_log_return": res.index_log_return,
            "index_cumulative_return": res.index_cumulative_return,
            "benefit_nominal": res.benefit_nominal_scheduled,
            "expense_nominal": res.expense_nominal_scheduled,
            "expected_benefit": res.expected_benefit_cashflows,
            "expected_expense": res.expected_expense_cashflows,
            "expected_total": res.expected_total_cashflows,
            "expected_payment_pv": expected_payment_pv,
            "cumulative_benefit_pv": cumulative_pv,
        }
    )


def _render_charts(res: sp.SPIAProjectionResult, contract: sp.SPIAContract) -> None:
    expected_payment_pv = res.expected_benefit_cashflows * res.discount_factors
    cumulative_pv = np.cumsum(expected_payment_pv)
    ages_r = contract.issue_age + res.reserve_times_years
    ages_pay = res.ages_at_payment

    st.subheader("Illustrations")
    _ret_labels = {
        "simple": "Month-over-month simple return",
        "log": "Month-over-month log return",
        "cumulative": "Cumulative return from month 0",
        "level": "Index level",
    }
    key = st.selectbox(
        "Index return chart (S&P proxy)",
        options=list(_ret_labels.keys()),
        format_func=lambda k: _ret_labels[k],
        key="index_return_chart_choice",
    )
    if key == "simple":
        ret_series = res.index_simple_return
        ylabel = "Simple return"
    elif key == "log":
        ret_series = res.index_log_return
        ylabel = "Log return"
    elif key == "cumulative":
        ret_series = res.index_cumulative_return
        ylabel = "Cumulative return (S_k/S_0 - 1)"
    else:
        ret_series = res.index_level_at_payment
        ylabel = "Index level"

    st.line_chart(pd.DataFrame({"age": ages_pay, "value": ret_series}).set_index("age"))
    st.caption(f"{ylabel} vs attained age at payment (proxy series; not an official index print).")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Survival to payment**")
        st.line_chart(
            pd.DataFrame({"age": res.ages_at_payment, "survival": res.survival_to_payment}).set_index("age")
        )
    with c2:
        st.markdown("**Expected PV per payment**")
        st.line_chart(
            pd.DataFrame({"age": res.ages_at_payment, "pv": expected_payment_pv}).set_index("age")
        )

    c3, c4 = st.columns(2)
    with c3:
        st.markdown("**Cumulative PV of benefits**")
        st.line_chart(
            pd.DataFrame({"age": res.ages_at_payment, "cumulative_pv": cumulative_pv}).set_index("age")
        )
    with c4:
        st.markdown("**Expected nominal benefit vs expense**")
        st.line_chart(
            pd.DataFrame(
                {
                    "age": res.ages_at_payment,
                    "benefit": res.expected_benefit_cashflows,
                    "expense": res.expected_expense_cashflows,
                }
            ).set_index("age")
        )

    st.markdown("**Economic reserve** (benefit + monthly expense, PV roll-forward)")
    st.line_chart(pd.DataFrame({"age": ages_r, "reserve": res.economic_reserve}).set_index("age"))


def _render_run_and_results() -> None:
    st.header("Run & results")

    with st.expander("Contract", expanded=True):
        c1, c2, c3 = st.columns(3)
        issue_age = c1.number_input("Issue age", min_value=0, max_value=120, value=65, step=1)
        sex = c2.selectbox("Sex (metadata)", options=["male", "female"], index=0)
        benefit_annual = c3.number_input("Annual benefit ($)", min_value=0.0, value=100_000.0, step=1_000.0)

    with st.expander("Yield curve", expanded=True):
        y_mode = st.radio(
            "Source",
            options=["flat", "zero_csv", "par_bootstrap"],
            format_func=lambda x: {
                "flat": "Flat zero rate",
                "zero_csv": "Zero curve CSV",
                "par_bootstrap": "Par yields CSV → bootstrap zeros",
            }[x],
            horizontal=True,
        )
        flat_rate = 0.04
        zero_csv = sp.DEFAULT_ZERO_CURVE_CSV
        par_csv = sp.DEFAULT_PAR_CURVE_CSV
        coupon_freq = 2
        if y_mode == "flat":
            flat_rate = st.number_input("Flat continuously compounded zero rate", value=0.04, format="%.4f")
        elif y_mode == "zero_csv":
            zero_csv = st.text_input("Zero curve CSV path", value=sp.DEFAULT_ZERO_CURVE_CSV)
        else:
            par_csv = st.text_input("Par yield CSV path", value=sp.DEFAULT_PAR_CURVE_CSV)
            coupon_freq = st.number_input("Coupon payments per year", min_value=1, value=2, step=1)

    with st.expander("Mortality", expanded=True):
        m_mode = st.radio(
            "Table",
            options=["synthetic", "qx_csv", "rp2014_mp2016"],
            format_func=lambda x: {
                "synthetic": "Synthetic (demo, wide age range)",
                "qx_csv": "Static q_x CSV",
                "rp2014_mp2016": "RP-2014 Healthy Male + MP-2016 (xlsx or cached CSV)",
            }[x],
            horizontal=True,
        )
        qx_csv = sp.DEFAULT_MORTALITY_QX_CSV
        rp_xlsx = sp.DEFAULT_RP2014_XLSX
        rp_out = sp.DEFAULT_RP2014_MALE_HEALTHY_QX_CSV
        mp_xlsx = sp.DEFAULT_MP2016_XLSX
        mp_out = sp.DEFAULT_MP2016_MALE_IMPROVEMENT_CSV
        if m_mode == "qx_csv":
            qx_csv = st.text_input("q_x CSV (columns age, qx)", value=sp.DEFAULT_MORTALITY_QX_CSV)
        elif m_mode == "rp2014_mp2016":
            st.caption("SOA workbooks are optional if matching CSV extracts already exist beside the xlsx paths.")
            rp_xlsx = st.text_input("RP-2014 xlsx", value=sp.DEFAULT_RP2014_XLSX)
            rp_out = st.text_input("RP-2014 healthy male qx cache CSV", value=sp.DEFAULT_RP2014_MALE_HEALTHY_QX_CSV)
            mp_xlsx = st.text_input("MP-2016 xlsx", value=sp.DEFAULT_MP2016_XLSX)
            mp_out = st.text_input("MP-2016 improvement cache CSV", value=sp.DEFAULT_MP2016_MALE_IMPROVEMENT_CSV)

    with st.expander("Expenses & valuation", expanded=True):
        expense_mode = st.radio(
            "Expenses",
            options=["csv", "manual"],
            format_func=lambda x: "Load from CSV" if x == "csv" else "Enter manually",
            horizontal=True,
        )
        expenses_csv = sp.DEFAULT_EXPENSES_CSV
        pol = 0.0
        prem_pct = 0.0
        monthly_ex = 0.0
        if expense_mode == "csv":
            expenses_csv = st.text_input("Expenses CSV path", value=sp.DEFAULT_EXPENSES_CSV)
        else:
            pol = float(st.number_input("Policy expense at issue ($)", value=0.0))
            prem_pct = float(
                st.number_input(
                    "Premium expense (% of single premium)",
                    value=0.0,
                    min_value=0.0,
                    max_value=99.99,
                    help="Enter 2 for 2%. Must stay below 100%.",
                )
            )
            monthly_ex = float(st.number_input("Monthly expense while alive ($)", value=0.0))
        valuation_year = st.number_input(
            "Valuation year (calendar)",
            min_value=1950,
            max_value=2100,
            value=2025,
            help="Used for RP+MP calendar-year mortality; ignored for static/synthetic q_x.",
        )
        horizon_age = st.number_input("Horizon age (stop monthly grid)", min_value=1, max_value=130, value=110)
        spread = st.number_input("Credit spread added to zero rate", value=0.0, format="%.4f")

    with st.expander("Economic scenario (benefit indexation & expense inflation)", expanded=True):
        use_index = st.checkbox(
            "Use S&P 500 proxy CSV for benefit return indexation",
            value=True,
            help="If off, index is flat (zero equity returns); benefits stay level in nominal terms.",
        )
        index_csv = st.text_input(
            "Index scenario CSV (columns: month, sp500_level for months 0..N)",
            value=sp.DEFAULT_SP500_SCENARIO_CSV,
        )
        expense_inflation_pct = st.number_input(
            "Expense annual inflation (%, not tied to S&P)",
            value=2.5,
            min_value=0.0,
            max_value=25.0,
            help="Applied monthly as (1 + annual)^(1/12) to maintenance expenses only.",
        )

    run = st.button("Run pricing", type="primary")

    if run:
        try:
            yc = _build_yield_curve(
                y_mode,  # type: ignore[arg-type]
                flat_rate=flat_rate,
                zero_csv=zero_csv,
                par_csv=par_csv,
                coupon_freq=coupon_freq,
            )
            mort, needs_vy = _build_mortality(
                m_mode,  # type: ignore[arg-type]
                qx_csv=qx_csv,
                rp_xlsx=rp_xlsx,
                rp_out_csv=rp_out,
                mp_xlsx=mp_xlsx,
                mp_out_csv=mp_out,
            )
            vy: int | None = int(valuation_year) if needs_vy else None
            vy_inputs = int(valuation_year)
            idx_path = str(_resolve_path(index_csv)) if use_index else None
            expense_annual_inflation = float(expense_inflation_pct) / 100.0

            contract = sp.SPIAContract(
                issue_age=int(issue_age),
                sex="male" if sex == "male" else "female",
                benefit_annual=float(benefit_annual),
            )

            expenses_arg: sp.ExpenseAssumptions | None = None
            if expense_mode == "manual":
                expenses_arg = sp.ExpenseAssumptions(
                    policy_expense_dollars=pol,
                    premium_expense_rate=prem_pct / 100.0,
                    monthly_expense_dollars=monthly_ex,
                )
                expenses_used = expenses_arg
            else:
                try:
                    expenses_used = sp.ExpenseAssumptions.load_from_csv(str(_resolve_path(expenses_csv)))
                except (FileNotFoundError, ValueError, KeyError):
                    expenses_used = sp.ExpenseAssumptions(0.0, 0.0, 0.0)

            res = sp.price_spia_single_premium(
                contract=contract,
                yield_curve=yc,
                mortality=mort,
                horizon_age=int(horizon_age),
                spread=float(spread),
                valuation_year=vy,
                expenses=expenses_arg,
                expenses_csv_path=str(_resolve_path(expenses_csv)),
                index_scenario_csv_path=idx_path,
                expense_annual_inflation=expense_annual_inflation,
            )
            st.session_state["pricing_res"] = res
            st.session_state["pricing_contract"] = contract
            st.session_state["pricing_err"] = None
            st.session_state["pricing_meta"] = {
                "yield_mode": y_mode,
                "mortality_mode": m_mode,
                "expense_mode": expense_mode,
            }
            st.session_state["pricing_excel_context"] = {
                "contract": contract,
                "yield_curve": yc,
                "mortality": mort,
                "horizon_age": int(horizon_age),
                "spread": float(spread),
                "valuation_year": vy_inputs,
                "expenses": expenses_used,
                "yield_mode": y_mode,
                "mortality_mode": m_mode,
                "expense_mode": expense_mode,
            }
            try:
                spec = excel_spec_from_launcher(
                    contract=contract,
                    yield_curve=yc,
                    mortality=mort,
                    horizon_age=int(horizon_age),
                    spread=float(spread),
                    valuation_year=vy_inputs,
                    expenses=expenses_used,
                    yield_mode_label=str(y_mode),
                    mortality_mode_label=str(m_mode),
                    expense_mode_label=str(expense_mode),
                    index_s0=float(res.index_s0),
                    index_levels_at_payment=res.index_level_at_payment,
                    expense_annual_inflation=float(res.expense_annual_inflation),
                )
                st.session_state["pricing_xlsx_bytes"] = build_workbook_from_spec(
                    spec,
                    out_path=None,
                    python_snapshot=ExcelPythonSnapshot(
                        pv_benefit=float(res.pv_benefit),
                        pv_monthly_expenses=float(res.pv_monthly_expenses),
                        pv_monthly_total=float(res.pv_benefit + res.pv_monthly_expenses),
                        single_premium=float(res.single_premium),
                        annuity_factor=float(res.annuity_factor),
                    ),
                )
                st.session_state.pop("pricing_xlsx_built_error", None)
            except Exception as ex:
                st.session_state["pricing_xlsx_bytes"] = None
                st.session_state["pricing_xlsx_built_error"] = repr(ex)
        except Exception as e:
            st.session_state["pricing_err"] = repr(e)
            st.session_state["pricing_res"] = None
            st.session_state.pop("pricing_excel_context", None)
            st.session_state.pop("pricing_xlsx_bytes", None)
            st.session_state.pop("pricing_xlsx_built_error", None)

    err = st.session_state.get("pricing_err")
    res = st.session_state.get("pricing_res")
    contract_state = st.session_state.get("pricing_contract")

    if err:
        st.error(err)
    if res is not None and contract_state is not None:
        st.success("Pricing completed.")
        meta = st.session_state.get("pricing_meta") or {}

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Single premium", f"${res.single_premium:,.2f}")
        m2.metric("PV benefit", f"${res.pv_benefit:,.2f}")
        m3.metric("PV monthly expenses", f"${res.pv_monthly_expenses:,.2f}")
        m4.metric("Annuity factor", f"{res.annuity_factor:,.6f}")

        st.caption(
            f"Yield: {meta.get('yield_mode')}; mortality: {meta.get('mortality_mode')}; "
            f"expenses: {meta.get('expense_mode')}."
        )

        df = _result_dataframe(res, contract_state)
        st.subheader("Month-by-month projection")
        st.dataframe(df, use_container_width=True, height=360)
        csv_bytes = df.to_csv(index=False).encode("utf-8")
        c_dl1, c_dl2 = st.columns(2)
        with c_dl1:
            st.download_button(
                "Download projection CSV",
                data=csv_bytes,
                file_name="spia_projection.csv",
                mime="text/csv",
            )
        with c_dl2:
            xb = st.session_state.get("pricing_xlsx_bytes")
            if isinstance(xb, bytes) and xb:
                st.download_button(
                    "Download Excel recalculation workbook",
                    data=xb,
                    file_name="spia_recalc_model.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Formula workbook for this run; ModelCheck sheet embeds Python PV/premium to verify Excel recalc vs Inputs (especially spread B9).",
                )
            elif st.session_state.get("pricing_xlsx_built_error"):
                st.caption(f"Excel export unavailable: {st.session_state['pricing_xlsx_built_error']}")

        _render_charts(res, contract_state)


def main() -> None:
    st.set_page_config(page_title="SPIA workspace", layout="wide")
    with st.sidebar:
        st.title("SPIA workspace")
        page = st.radio(
            "Section",
            options=["overview", "run", "tests"],
            format_func=lambda x: {
                "overview": "Overview",
                "run": "Run & results",
                "tests": "Unit tests",
            }[x],
        )
        st.divider()
        st.caption(f"Project root: `{ROOT}`")

    if page == "overview":
        _render_overview()
    elif page == "run":
        _render_run_and_results()
    else:
        render_unit_tests_page(embedded=True)


if __name__ == "__main__":
    main()

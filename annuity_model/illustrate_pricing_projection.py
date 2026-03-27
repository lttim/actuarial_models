"""
Save illustration charts for SPIA projection (with index scenario + expense inflation).

Default: uses treasury zero curve, RP-2014 + MP-2016 when workbooks exist, synthetic index CSV,
and 2.5%/year expense inflation. Index return charts can be saved as PNGs or toggled interactively.

Examples:
    python illustrate_pricing_projection.py
    python illustrate_pricing_projection.py --index-return all
    python illustrate_pricing_projection.py --interactive
"""

from __future__ import annotations

import argparse
import os

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.widgets import RadioButtons

from pricing_projection import (
    DEFAULT_SP500_SCENARIO_CSV,
    ExpenseAssumptions,
    MortalityTableRP2014MP2016,
    SPIAContract,
    YieldCurve,
    ensure_mp2016_male_improvement_csv,
    ensure_rp2014_male_healthy_annuitant_qx_csv,
    load_index_scenario_monthly_csv,
    price_spia_single_premium,
)


def _safe_mkdir(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def _plot_return_panel(ax, ages: np.ndarray, res, mode: str) -> None:
    if mode == "simple":
        y = res.index_simple_return
        title = "S&P proxy: month-over-month simple return"
        ylab = "Return"
    elif mode == "log":
        y = res.index_log_return
        title = "S&P proxy: month-over-month log return"
        ylab = "Log return"
    elif mode == "cumulative":
        y = res.index_cumulative_return
        title = "S&P proxy: cumulative return from month 0"
        ylab = "S_k/S_0 - 1"
    else:
        y = res.index_level_at_payment
        title = "S&P proxy: index level"
        ylab = "Level"
    ax.plot(ages, y, linewidth=1.5)
    ax.set_title(title)
    ax.set_xlabel("Attained age at payment")
    ax.set_ylabel(ylab)
    ax.grid(True, alpha=0.3)


def main() -> None:
    p = argparse.ArgumentParser(description="SPIA illustration plots")
    p.add_argument(
        "--index-return",
        choices=["all", "simple", "log", "cumulative", "level"],
        default="all",
        help="Which index return chart(s) to save (ignored with --interactive).",
    )
    p.add_argument(
        "--interactive",
        action="store_true",
        help="Open one figure with a radio-button toggle between return views.",
    )
    p.add_argument(
        "--expense-inflation-pct",
        type=float,
        default=2.5,
        help="Annual expense inflation percent (maintenance expenses only).",
    )
    p.add_argument(
        "--scenario-csv",
        type=str,
        default=DEFAULT_SP500_SCENARIO_CSV,
        help="Monthly index CSV (month, sp500_level for 0..N). Use empty string for flat index.",
    )
    args = p.parse_args()

    contract = SPIAContract(issue_age=65, sex="male", benefit_annual=100_000.0)
    horizon_age = 110
    spread = 0.0
    valuation_year = 2025
    expense_annual_inflation = float(args.expense_inflation_pct) / 100.0

    zero_curve_csv = "treasury_zero_rate_curve_latest.csv"
    rp2014_xlsx = "rp2014_mort_tab_rates_exposure.xlsx"
    mp2016_xlsx = "mp2016_rates.xlsx"

    out_dir = "illustrations"
    _safe_mkdir(out_dir)

    yc = YieldCurve.load_zero_curve_csv(zero_curve_csv)
    base_qx = ensure_rp2014_male_healthy_annuitant_qx_csv(
        rp2014_xlsx_path=rp2014_xlsx,
        out_csv_path="rp2014_male_healthy_annuitant_qx_2014.csv",
    )
    mp_ages, mp_years, mp_i = ensure_mp2016_male_improvement_csv(
        mp2016_xlsx_path=mp2016_xlsx,
        out_csv_path="mp2016_male_improvement_rates.csv",
    )
    mortality = MortalityTableRP2014MP2016(
        base_qx_2014=base_qx,
        mp2016_ages=mp_ages,
        mp2016_years=mp_years,
        mp2016_i_matrix=mp_i,
        base_year=2014,
    )

    dt = 1.0 / 12.0
    n_months = max(1, int(round((horizon_age - contract.issue_age) / dt)))
    scen_path = args.scenario_csv.strip()
    if scen_path:
        idx_path = scen_path
        try:
            load_index_scenario_monthly_csv(idx_path, n_months=n_months)
        except FileNotFoundError:
            print(f"Scenario not found: {idx_path}, using flat index.")
            idx_path = None
        except ValueError as e:
            print(f"Scenario validation failed ({e}); using flat index.")
            idx_path = None
    else:
        idx_path = None

    res = price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mortality,
        horizon_age=horizon_age,
        spread=spread,
        valuation_year=valuation_year,
        expenses=ExpenseAssumptions(0.0, 0.0, 0.0),
        index_scenario_csv_path=idx_path,
        expense_annual_inflation=expense_annual_inflation,
    )

    ages_pay = res.ages_at_payment
    expected_payment_pv = res.expected_benefit_cashflows * res.discount_factors
    cumulative_pv = np.cumsum(expected_payment_pv)
    ages_at_reserve = contract.issue_age + res.reserve_times_years

    def save_core_plots() -> None:
        plt.figure(figsize=(9, 5))
        plt.plot(ages_pay, res.survival_to_payment, linewidth=2)
        plt.title("SPIA: Survival Probability to Each Monthly Payment Date")
        plt.xlabel("Attained Age at Payment")
        plt.ylabel("P(alive at payment)")
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.savefig(f"{out_dir}/pricing_spia_survival_vs_age.png", dpi=160)
        plt.close()

        plt.figure(figsize=(9, 5))
        plt.plot(ages_pay, expected_payment_pv, linewidth=2)
        plt.title("SPIA: Expected PV Contribution per Monthly Benefit")
        plt.xlabel("Attained Age at Payment")
        plt.ylabel("Expected Present Value ($)")
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.savefig(f"{out_dir}/pricing_spia_expected_pv_contribution_vs_age.png", dpi=160)
        plt.close()

        plt.figure(figsize=(9, 5))
        plt.plot(ages_pay, cumulative_pv, linewidth=2)
        plt.title("SPIA: Cumulative PV of Expected Benefit Payments")
        plt.xlabel("Attained Age at Payment")
        plt.ylabel("Cumulative Present Value ($)")
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.savefig(f"{out_dir}/pricing_spia_cumulative_pv_vs_age.png", dpi=160)
        plt.close()

        plt.figure(figsize=(9, 5))
        plt.plot(ages_pay, res.expected_benefit_cashflows, linewidth=2, color="tab:blue")
        plt.title("SPIA: Expected Monthly Benefit Payments (Nominal)")
        plt.xlabel("Attained Age at Payment")
        plt.ylabel("Expected Benefit Payment ($)")
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.savefig(f"{out_dir}/pricing_spia_expected_benefit_cashflows_vs_age.png", dpi=160)
        plt.close()

        plt.figure(figsize=(9, 5))
        plt.plot(ages_pay, res.expected_expense_cashflows, linewidth=2, color="tab:orange")
        plt.title("SPIA: Expected Monthly Expense Cashflows (Nominal)")
        plt.xlabel("Attained Age at Payment")
        plt.ylabel("Expected Expense Payment ($)")
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.savefig(f"{out_dir}/pricing_spia_expected_expense_cashflows_vs_age.png", dpi=160)
        plt.close()

        plt.figure(figsize=(9, 5))
        plt.plot(ages_at_reserve, res.economic_reserve, linewidth=2, color="tab:green")
        plt.title("SPIA: Projected Economic Reserves (Benefit + Monthly Expenses)")
        plt.xlabel("Attained Age")
        plt.ylabel("Economic Reserve ($, PV at valuation time)")
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.savefig(f"{out_dir}/pricing_spia_economic_reserve_vs_age.png", dpi=160)
        plt.close()

    save_core_plots()

    modes = ["simple", "log", "cumulative", "level"]
    if args.interactive:
        fig, ax = plt.subplots(figsize=(9, 5))
        plt.subplots_adjust(left=0.25)

        def redraw(mode: str) -> None:
            ax.clear()
            _plot_return_panel(ax, ages_pay, res, mode)
            fig.canvas.draw_idle()

        redraw("simple")
        rax = plt.axes((0.02, 0.4, 0.18, 0.2))
        radio = RadioButtons(rax, ("Simple", "Log", "Cumulative", "Level"))
        radio.on_clicked(lambda label: redraw(label.lower()))
        plt.suptitle("Toggle: S&P proxy return view (illustrative series)")
        plt.show()
    else:
        to_save = modes if args.index_return == "all" else [args.index_return]
        for m in to_save:
            fig, ax = plt.subplots(figsize=(9, 5))
            _plot_return_panel(ax, ages_pay, res, m)
            plt.tight_layout()
            plt.savefig(f"{out_dir}/pricing_spia_index_return_{m}.png", dpi=160)
            plt.close()

    print("SPIA illustration (policy):")
    print(f"  Issue age: {contract.issue_age}, benefit annual: {contract.benefit_annual:,.0f}")
    print(f"  Valuation year: {valuation_year}, horizon age: {horizon_age}, spread: {spread}")
    print(f"  Expense inflation (annual): {expense_annual_inflation:.4f}")
    print(f"  Single premium (incl. issue expenses): {res.single_premium:,.2f}")
    print(f"  PV benefits: {res.pv_benefit:,.2f}")
    print(f"  PV monthly expenses: {res.pv_monthly_expenses:,.2f}")
    print(f"  Annuity factor (level $1): {res.annuity_factor:,.6f}")
    print("")
    print("Saved plots under illustrations/ (plus index return PNGs unless --interactive).")


if __name__ == "__main__":
    main()

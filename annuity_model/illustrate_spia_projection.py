import numpy as np
import matplotlib.pyplot as plt

from spia_projection import (
    SPIAContract,
    YieldCurve,
    MortalityTableRP2014MP2016,
    ensure_mp2016_male_improvement_csv,
    ensure_rp2014_male_healthy_annuitant_qx_csv,
    price_spia_single_premium,
    ExpenseAssumptions,
)


def _safe_mkdir(path: str) -> None:
    import os

    os.makedirs(path, exist_ok=True)


def main() -> None:
    # Policy inputs (from conversation)
    contract = SPIAContract(issue_age=65, sex="male", benefit_annual=100_000.0)
    horizon_age = 110
    spread = 0.0
    valuation_year = 2025  # 12/31/2025 valuation date

    # Files in working folder
    zero_curve_csv = "treasury_zero_rate_curve_latest.csv"
    rp2014_xlsx = "rp2014_mort_tab_rates_exposure.xlsx"
    mp2016_xlsx = "mp2016_rates.xlsx"

    out_dir = "illustrations"
    _safe_mkdir(out_dir)

    # Load curves / mortality
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

    # Compute projection
    res = price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mortality,
        horizon_age=horizon_age,
        spread=spread,
        valuation_year=valuation_year,
    )

    dt = 1.0 / 12.0
    b_month = contract.benefit_annual / 12.0
    expected_payment_pv = b_month * res.survival_to_payment * res.discount_factors
    expected_benefit_cashflows = res.expected_benefit_cashflows
    expected_expense_cashflows = res.expected_expense_cashflows
    expected_total_cashflows = res.expected_total_cashflows

    # Economic reserves are stored at t=0 plus each monthly payment date.
    ages_at_reserve = contract.issue_age + res.reserve_times_years

    # Plot 1: survival
    plt.figure(figsize=(9, 5))
    plt.plot(res.ages_at_payment, res.survival_to_payment, linewidth=2)
    plt.title("SPIA: Survival Probability to Each Monthly Payment Date")
    plt.xlabel("Attained Age at Payment")
    plt.ylabel("P(alive at payment)")
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(f"{out_dir}/spia_survival_vs_age.png", dpi=160)
    plt.close()

    # Plot 2: expected PV contribution per payment date
    plt.figure(figsize=(9, 5))
    plt.plot(res.ages_at_payment, expected_payment_pv, linewidth=2)
    plt.title("SPIA: Expected PV Contribution per Monthly Payment")
    plt.xlabel("Attained Age at Payment")
    plt.ylabel("Expected Present Value of Payment ($)")
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(f"{out_dir}/spia_expected_pv_contribution_vs_age.png", dpi=160)
    plt.close()

    # Plot 3: cumulative PV
    cumulative_pv = np.cumsum(expected_payment_pv)
    plt.figure(figsize=(9, 5))
    plt.plot(res.ages_at_payment, cumulative_pv, linewidth=2)
    plt.title("SPIA: Cumulative PV of Expected Payments")
    plt.xlabel("Attained Age at Payment")
    plt.ylabel("Cumulative Present Value ($)")
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(f"{out_dir}/spia_cumulative_pv_vs_age.png", dpi=160)
    plt.close()

    # Plot 4: projected monthly benefit cashflows (expected nominal)
    plt.figure(figsize=(9, 5))
    plt.plot(res.ages_at_payment, expected_benefit_cashflows, linewidth=2, color="tab:blue")
    plt.title("SPIA: Expected Monthly Benefit Payments (Nominal)")
    plt.xlabel("Attained Age at Payment")
    plt.ylabel("Expected Benefit Payment ($)")
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(f"{out_dir}/spia_expected_benefit_cashflows_vs_age.png", dpi=160)
    plt.close()

    # Plot 5: projected monthly expense cashflows (expected nominal)
    plt.figure(figsize=(9, 5))
    plt.plot(res.ages_at_payment, expected_expense_cashflows, linewidth=2, color="tab:orange")
    plt.title("SPIA: Expected Monthly Expense Cashflows (Nominal)")
    plt.xlabel("Attained Age at Payment")
    plt.ylabel("Expected Expense Payment ($)")
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(f"{out_dir}/spia_expected_expense_cashflows_vs_age.png", dpi=160)
    plt.close()

    # Plot 6: projected economic reserves
    plt.figure(figsize=(9, 5))
    plt.plot(ages_at_reserve, res.economic_reserve, linewidth=2, color="tab:green")
    plt.title("SPIA: Projected Economic Reserves (Benefit + Monthly Expenses)")
    plt.xlabel("Attained Age")
    plt.ylabel("Economic Reserve ($, PV at valuation time)")
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(f"{out_dir}/spia_economic_reserve_vs_age.png", dpi=160)
    plt.close()

    # Print a small verification summary
    checkpoints = [65, 70, 75, 80, 85, 90, 100, 110]
    print("SPIA illustration (policy):")
    print(f"  Issue age: {contract.issue_age}, benefit annual: {contract.benefit_annual:,.0f}")
    print(f"  Valuation year: {valuation_year}, horizon age: {horizon_age}, spread: {spread}")
    print(f"  Single premium (incl. issue expenses): {res.single_premium:,.2f}")
    print(f"  PV benefits: {res.pv_benefit:,.2f}")
    print(f"  PV monthly expenses: {res.pv_monthly_expenses:,.2f}")
    print(f"  Annuity factor: {res.annuity_factor:,.6f}")
    print("")
    for age in checkpoints:
        # Find nearest payment date age
        idx = int(np.argmin(np.abs(res.ages_at_payment - age)))
        t = res.times_years[idx]
        print(
            f"  Age~{age}: t={t:.3f}y, "
            f"S(t)={res.survival_to_payment[idx]:.6f}, "
            f"DF(t)={res.discount_factors[idx]:.6f}, "
            f"Expected payment PV=${expected_payment_pv[idx]:,.2f}"
        )
    print("")
    print("Saved plots to:")
    print(f"  {out_dir}/spia_survival_vs_age.png")
    print(f"  {out_dir}/spia_expected_pv_contribution_vs_age.png")
    print(f"  {out_dir}/spia_cumulative_pv_vs_age.png")
    print(f"  {out_dir}/spia_expected_benefit_cashflows_vs_age.png")
    print(f"  {out_dir}/spia_expected_expense_cashflows_vs_age.png")
    print(f"  {out_dir}/spia_economic_reserve_vs_age.png")


if __name__ == "__main__":
    main()


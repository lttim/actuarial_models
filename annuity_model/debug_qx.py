from spia_projection import (
    ensure_rp2014_male_healthy_annuitant_qx_csv,
    ensure_mp2016_male_improvement_csv,
    MortalityTableRP2014MP2016,
)


def main() -> None:
    base_qx = ensure_rp2014_male_healthy_annuitant_qx_csv(
        rp2014_xlsx_path="rp2014_mort_tab_rates_exposure.xlsx",
        out_csv_path="tmp_rp2014_male_healthy_annuitant_qx_2014.csv",
    )
    mp_ages, mp_years, mp_i = ensure_mp2016_male_improvement_csv(
        mp2016_xlsx_path="mp2016_rates.xlsx",
        out_csv_path="tmp_mp2016_male_improvement_rates.csv",
    )

    model = MortalityTableRP2014MP2016(
        base_qx_2014=base_qx,
        mp2016_ages=mp_ages,
        mp2016_years=mp_years,
        mp2016_i_matrix=mp_i,
        base_year=2014,
    )

    ages = [65, 70, 75, 80, 90, 100]
    years = [2025, 2026, 2030]

    print("Base qx (2014) by age:")
    for a in ages:
        print(f"  age {a}: qx_2014={model.base_qx_2014.qx_at_int_age(a):.6f}")

    print("\nImplied qx(age, calendar_year) under current implementation:")
    for y in years:
        print(f"  calendar_year {y}:")
        for a in ages:
            q = model.qx_at_int_age_and_calendar_year(age_int=a, calendar_year=y)
            print(f"    age {a}: qx={q:.6f}")


if __name__ == "__main__":
    main()


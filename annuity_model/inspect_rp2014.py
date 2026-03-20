import pandas as pd

rp = "rp2014_mort_tab_rates_exposure.xlsx"
SHEETS = ["Total Dataset", "White Collar", "Blue Collar", "Bottom Quartile", "Top Quartile"]


def main() -> None:
    for sheet in SHEETS:
        df = pd.read_excel(rp, sheet_name=sheet, header=None, nrows=120)
        # Print a small "fingerprint" of likely header/section labels.
        # Many SOA workbooks have labels in column A/B; we dump the first few columns.
        print(f"\n=== {sheet} ===")
        print("shape", df.shape)

        # Show first 40 rows for first 6 columns.
        cols_to_show = min(6, df.shape[1])
        preview = df.iloc[:40, :cols_to_show]
        print(preview.to_string(index=False, header=False))


if __name__ == "__main__":
    main()


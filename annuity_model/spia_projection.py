"""
Single Premium Immediate Annuity (SPIA) projection scaffolding.

Contract assumptions (set from the conversation):
- Single life
- Monthly payments; annual benefit sets the starting monthly accrual before indexation
- Benefits can grow by monthly **return indexation** from an external index level path (CSV)
- Maintenance expenses can grow by a fixed **annual inflation** rate (monthly compounding), separate from the index
- Payment timing: end of period
- Payments stop at death (i.e., pay at time t_k iff alive at that payment date)

Pricing (single premium):
    Premium = policy_expense + PV(benefits + monthly expenses), grossed up for premium tax,
    where monthly benefits can grow by **return indexation** from an S&P 500 proxy scenario
    (monthly index levels by payment month) and monthly expenses grow by a **fixed annual
    inflation** rate compounded monthly (independent of the equity index).

Interest:
- Uses continuous-compounding "zero rates" z(t).
  Discount factors: DF(t) = exp(-(z(t) + spread) * t)
- Risk-free by default: spread = 0.

US industry mortality (table selection, data must be supplied):
- RP-2014 Healthy Annuitant Male base table
- MP-2016 mortality improvement scale

Important:
- RP-2014 and MP-2016 are proprietary to SOA.
- This code does NOT embed the table rates.
- Provide annual one-year death probabilities q_x (or equivalent) in a CSV.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal, Optional, Sequence, Tuple

import math

import numpy as np
import pandas as pd


Compounding = Literal["continuous"]


DEFAULT_ZERO_CURVE_CSV = "treasury_zero_rate_curve_latest.csv"
DEFAULT_PAR_CURVE_CSV = "treasury_par_yield_curve_latest.csv"
DEFAULT_MORTALITY_QX_CSV = "rp2014_male_annuitant_qx.csv"
DEFAULT_RP2014_XLSX = "rp2014_mort_tab_rates_exposure.xlsx"
DEFAULT_MP2016_XLSX = "mp2016_rates.xlsx"
DEFAULT_RP2014_MALE_HEALTHY_QX_CSV = "rp2014_male_healthy_annuitant_qx_2014.csv"
DEFAULT_MP2016_MALE_IMPROVEMENT_CSV = "mp2016_male_improvement_rates.csv"
DEFAULT_EXPENSES_CSV = "expenses_assumptions_us_placeholders.csv"
DEFAULT_SP500_SCENARIO_CSV = "sp500_scenario_projection_monthly.csv"


def load_index_scenario_monthly_csv(
    path: str,
    *,
    n_months: int,
    month_col: str = "month",
    level_col: str = "sp500_level",
) -> tuple[float, np.ndarray]:
    """
    Load index levels for payment months 0..n_months.

    The CSV must include month=0 and a **contiguous** block of integer months
    ``0..M`` with no gaps (extra months beyond ``n_months`` are ignored).

    If ``M < n_months``, the level at month ``M`` is **held flat** for
    ``M+1..n_months`` (zero index returns on the extended tail). If the file
    already covers ``0..n_months`` or more, no extension is applied.

    Returns
    -------
    s0:
        Index level at month 0.
    levels_payment:
        Shape (n_months,); levels_payment[k-1] is the index at payment month k.
    """
    if n_months < 1:
        raise ValueError("n_months must be positive.")

    df = pd.read_csv(path)
    if month_col not in df.columns or level_col not in df.columns:
        raise ValueError(f"CSV must contain columns '{month_col}' and '{level_col}'.")

    months = df[month_col].to_numpy(dtype=int)
    levels = df[level_col].to_numpy(dtype=float)
    if np.any(levels <= 0.0) or np.any(np.isnan(levels)):
        raise ValueError("Index levels must be finite and strictly positive.")

    order = np.argsort(months)
    months = months[order]
    levels = levels[order]

    month_levels: dict[int, float] = {}
    for m, lev in zip(months.tolist(), levels.tolist()):
        mi = int(m)
        if mi < 0:
            raise ValueError("Index scenario month indices must be non-negative.")
        month_levels[mi] = float(lev)

    if 0 not in month_levels:
        raise ValueError("Index scenario must include month 0.")

    m_max = max(month_levels.keys())
    expected_keys = set(range(m_max + 1))
    if set(month_levels.keys()) != expected_keys:
        missing = sorted(expected_keys - set(month_levels.keys()))
        extra = sorted(set(month_levels.keys()) - expected_keys)
        msg = (
            f"Index scenario months must be contiguous from 0 through {m_max} with no gaps "
            f"(got non-contiguous or duplicate-spanning keys)."
        )
        if missing:
            msg += f" Missing in prefix: {missing[:20]}{'...' if len(missing) > 20 else ''}."
        if extra:
            msg += f" Unexpected: {extra[:20]}{'...' if len(extra) > 20 else ''}."
        raise ValueError(msg)

    if m_max < n_months:
        last = float(month_levels[m_max])
        for m in range(m_max + 1, n_months + 1):
            month_levels[m] = last

    required = range(0, n_months + 1)
    for m in required:
        if m not in month_levels:
            raise ValueError(f"Index scenario internal error: missing month {m} after load/extend.")

    s0 = float(month_levels[0])
    levels_payment = np.zeros(n_months, dtype=float)
    for k in range(1, n_months + 1):
        levels_payment[k - 1] = float(month_levels[k])
    return s0, levels_payment


def flat_index_scenario(n_months: int, *, level: float = 100.0) -> tuple[float, np.ndarray]:
    """Constant index (zero returns) for n_months payments after month-0 level."""
    s0 = float(level)
    return s0, np.full(n_months, s0, dtype=float)


def monthly_rate_from_annual_inflation(annual_inflation: float) -> float:
    """Convert annual CPI-style inflation to an equivalent monthly compound rate."""
    if annual_inflation < -0.9999:
        raise ValueError("annual_inflation must be > -1 (approximately).")
    return float((1.0 + annual_inflation) ** (1.0 / 12.0) - 1.0)


def _benefit_expense_and_index_returns(
    *,
    base_monthly: float,
    monthly_expense: float,
    s0: float,
    levels_payment: np.ndarray,
    expense_annual_inflation: float,
) -> tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray, np.ndarray]:
    """
    Return-indexed benefits, inflation-only expenses, and per-payment index return metrics.

    Benefit month k (1-based payment): b_k = b_{k-1} * (S_k / S_{k-1}) with
    b_1 = base_monthly * (S_1 / S_0).
    """
    n = int(levels_payment.shape[0])
    if n < 1:
        raise ValueError("levels_payment must be non-empty.")

    b = np.zeros(n, dtype=float)
    e = np.zeros(n, dtype=float)
    b[0] = float(base_monthly) * (float(levels_payment[0]) / float(s0))
    e[0] = float(monthly_expense)
    g = monthly_rate_from_annual_inflation(float(expense_annual_inflation)) if expense_annual_inflation else 0.0

    for k in range(1, n):
        if levels_payment[k - 1] <= 0.0:
            raise ValueError("Index levels must stay positive for return indexation.")
        b[k] = b[k - 1] * (float(levels_payment[k]) / float(levels_payment[k - 1]))
        e[k] = e[k - 1] * (1.0 + g)

    simple = np.zeros(n, dtype=float)
    logret = np.zeros(n, dtype=float)
    cumu = np.zeros(n, dtype=float)
    prev = float(s0)
    for k in range(n):
        cur = float(levels_payment[k])
        simple[k] = cur / prev - 1.0
        logret[k] = math.log(cur / prev) if cur > 0.0 and prev > 0.0 else float("nan")
        cumu[k] = cur / float(s0) - 1.0
        prev = cur

    return b, e, simple, logret, cumu


@dataclass(frozen=True)
class YieldCurve:
    """
    Continuous-compounding zero curve.

    Attributes
    ----------
    maturities_years:
        Array of curve node maturities in years (e.g., [0.5, 1.0, 1.5, ...]).
    zero_rates:
        Corresponding continuously-compounded spot/zero rates z(t).
    """

    maturities_years: np.ndarray
    zero_rates: np.ndarray

    @staticmethod
    def from_flat_rate(flat_zero_rate: float) -> "YieldCurve":
        mats = np.array([0.0, 1.0], dtype=float)
        zeros = np.array([flat_zero_rate, flat_zero_rate], dtype=float)
        return YieldCurve(mats, zeros)

    @staticmethod
    def load_zero_curve_csv(
        path: str,
        *,
        maturity_col: str = "maturity_years",
        rate_col: str = "zero_rate",
    ) -> "YieldCurve":
        df = pd.read_csv(path)
        mats = df[maturity_col].to_numpy(dtype=float)
        zeros = df[rate_col].to_numpy(dtype=float)
        order = np.argsort(mats)
        return YieldCurve(mats[order], zeros[order])

    @staticmethod
    def load_par_yield_csv_and_bootstrap(
        path: str,
        *,
        maturity_col: str = "maturity_years",
        par_rate_col: str = "par_yield",
        coupon_freq: int = 2,
        interpolation: Literal["linear"] = "linear",
    ) -> "YieldCurve":
        """
        Bootstrap discount factors from a par-yield curve.

        Inputs note:
        - Treasury par yields are commonly quoted with semiannual coupons.
        - To bootstrap discount factors on the coupon grid, this function will
          linearly interpolate the par yields onto each coupon date.
        - This is sufficient for a scaffold; later you can replace with a
          more faithful bootstrapping approach as needed.
        """
        df = pd.read_csv(path)
        maturities = df[maturity_col].to_numpy(dtype=float)
        par_yields = df[par_rate_col].to_numpy(dtype=float)
        order = np.argsort(maturities)
        maturities = maturities[order]
        par_yields = par_yields[order]

        zero_mats, zero_rates = bootstrap_zero_rates_from_par_yields(
            par_maturities_years=maturities,
            par_yields=par_yields,
            coupon_freq=coupon_freq,
        )
        return YieldCurve(zero_mats, zero_rates)

    def discount_factors(
        self,
        times_years: np.ndarray,
        *,
        spread: float = 0.0,
        compounding: Compounding = "continuous",
    ) -> np.ndarray:
        """
        Convert the stored zero curve into discount factors at arbitrary times.

        Uses log-linear interpolation on discount factors (stable).
        """
        if compounding != "continuous":
            raise ValueError("Only continuous compounding is implemented.")

        t = np.asarray(times_years, dtype=float)
        if t.ndim != 1:
            raise ValueError("times_years must be a 1D array.")

        # Ensure increasing curve times.
        mats = np.asarray(self.maturities_years, dtype=float)
        zeros = np.asarray(self.zero_rates, dtype=float)
        if np.any(np.diff(mats) < 0):
            order = np.argsort(mats)
            mats = mats[order]
            zeros = zeros[order]

        # Compute curve-node discount factors with (optional) spread.
        df_curve = np.exp(-(zeros + spread) * mats)
        log_df_curve = np.log(df_curve)

        # Interpolate on log(DF) within the curve range.
        # For extrapolation, use constant zero-rate beyond endpoints:
        #   DF(t) = exp(-(z_end + spread) * t)
        #
        # This avoids the incorrect behavior where DF would become constant
        # beyond the final node.
        df_t = np.empty_like(t, dtype=float)
        t0 = float(mats[0])
        tN = float(mats[-1])

        mask_low = t <= t0
        mask_high = t >= tN
        mask_mid = ~(mask_low | mask_high)

        if np.any(mask_low):
            df_t[mask_low] = np.exp(-((float(zeros[0]) + spread) * t[mask_low]))
        if np.any(mask_high):
            df_t[mask_high] = np.exp(-((float(zeros[-1]) + spread) * t[mask_high]))
        if np.any(mask_mid):
            log_df_t_mid = np.interp(t[mask_mid], mats, log_df_curve)
            df_t[mask_mid] = np.exp(log_df_t_mid)

        return df_t


def bootstrap_zero_rates_from_par_yields(
    par_maturities_years: Sequence[float],
    par_yields: Sequence[float],
    *,
    coupon_freq: int = 2,
) -> Tuple[np.ndarray, np.ndarray]:
    """
    Bootstrap continuous zero rates from par yields.

    Assumptions (common for a scaffold):
    - Coupon payments at times j / coupon_freq.
    - Par yield is treated as the bond's coupon rate (annual),
      with coupon per period = par_yield / coupon_freq.

    Parameters
    ----------
    par_maturities_years:
        Curve node maturities in years. Typically includes values up to the max horizon.
    par_yields:
        Par yields at those maturities (annualized, decimal).
    coupon_freq:
        Number of coupon payments per year (2 for semiannual).

    Returns
    -------
    zero_mats_years:
        Coupon-grid maturities used for bootstrapping (e.g., every 0.5 years if coupon_freq=2).
    zero_rates:
        Continuous-compounding zero rates at those maturities.
    """
    if coupon_freq <= 0:
        raise ValueError("coupon_freq must be positive.")

    par_maturities = np.asarray(par_maturities_years, dtype=float)
    y = np.asarray(par_yields, dtype=float)
    if par_maturities.shape != y.shape:
        raise ValueError("par_maturities_years and par_yields must have the same length.")

    order = np.argsort(par_maturities)
    par_maturities = par_maturities[order]
    y = y[order]

    max_T = float(np.max(par_maturities))
    period = 1.0 / coupon_freq

    # Build the coupon grid to bootstrap discount factors.
    n_steps = int(round(max_T / period))
    t_nodes = np.arange(1, n_steps + 1, dtype=float) * period
    par_on_nodes = np.interp(t_nodes, par_maturities, y)

    # Discount factors DF(0)=1; compute DF(j*period) sequentially.
    df = np.zeros_like(t_nodes)
    df_prev = []  # store DF values for summations

    for k, T in enumerate(t_nodes, start=1):
        yk = float(par_on_nodes[k - 1])
        c = yk / coupon_freq  # coupon per period assuming coupon rate equals par yield

        # Sum DF at previous coupon dates: sum_{j=1..k-1} DF(j*period)
        sum_prev = float(np.sum(df_prev)) if df_prev else 0.0

        # Par bond price at par:
        # 1 = c * (sum_prev + DF_k) + DF_k
        # => DF_k = (1 - c * sum_prev) / (1 + c)
        df_k = (1.0 - c * sum_prev) / (1.0 + c)

        df[k - 1] = df_k
        df_prev.append(df_k)

    zero_rates = -np.log(df) / t_nodes
    return t_nodes, zero_rates


@dataclass(frozen=True)
class MortalityTableQx:
    """
    Annual one-year death probabilities q_x by integer attained age.

    Attributes
    ----------
    ages:
        Integer ages x for which qx is provided.
    qx:
        One-year death probability q_x for each age in `ages`.
        Interpreted as P(death in [x, x+1) | alive at age x).
    """

    ages: np.ndarray
    qx: np.ndarray

    @staticmethod
    def load_qx_csv(
        path: str,
        *,
        age_col: str = "age",
        qx_col: str = "qx",
        dropna: bool = True,
    ) -> "MortalityTableQx":
        df = pd.read_csv(path)
        if dropna:
            df = df.dropna(subset=[age_col, qx_col])
        ages = df[age_col].to_numpy(dtype=int)
        qx = df[qx_col].to_numpy(dtype=float)
        order = np.argsort(ages)
        return MortalityTableQx(ages[order], qx[order])

    def qx_at_int_age(self, age_int: int) -> float:
        if age_int < int(self.ages[0]) or age_int > int(self.ages[-1]):
            raise ValueError(f"age_int={age_int} outside mortality table range [{self.ages[0]}, {self.ages[-1]}].")
        idx = int(age_int - int(self.ages[0]))
        # This assumes `ages` are contiguous starting at ages[0]. If not contiguous, fall back to lookup.
        if not np.array_equal(self.ages, np.arange(int(self.ages[0]), int(self.ages[0]) + len(self.ages))):
            # Non-contiguous: use dictionary-like lookup via search.
            idx_arr = np.where(self.ages == age_int)[0]
            if len(idx_arr) != 1:
                raise ValueError("Mortality table ages are ambiguous or missing.")
            return float(self.qx[idx_arr[0]])
        return float(self.qx[idx])

    def monthly_survival_to_payment(
        self,
        *,
        issue_age: int,
        n_months: int,
        valuation_year: int | None = None,
    ) -> np.ndarray:
        """
        Compute S(t_k)=P(T >= t_k) on a monthly grid (t_k=k/12).

        Method:
        - Use annual q_x as one-year death probabilities.
        - Convert to annual force-of-mortality per integer age:
              mu_x = -ln(1 - q_x)
        - Assume mu_x constant within the year, so monthly survival over 1/12
          at integer age x is exp(-mu_x / 12).
        """
        if n_months <= 0:
            raise ValueError("n_months must be positive.")

        dt = 1.0 / 12.0
        S = np.ones(n_months, dtype=float)

        log_S = 0.0
        for k in range(1, n_months + 1):
            # Month k interval starts at t_{k-1} and ends at t_k.
            # Attained age at t_{k-1}:
            age_start = issue_age + (k - 1) * dt
            age_int = int(math.floor(age_start))

            qx = self.qx_at_int_age(age_int)
            if not (0.0 <= qx < 1.0):
                raise ValueError(f"qx must satisfy 0 <= qx < 1, got qx={qx} at age={age_int}")

            mu = -math.log(1.0 - qx)  # piecewise-constant hazard within [age_int, age_int+1)
            survival_interval = math.exp(-mu * dt)
            log_S += math.log(survival_interval)
            S[k - 1] = math.exp(log_S)

        return S


def _is_numeric(x: object) -> bool:
    return isinstance(x, (int, float, np.number)) and not pd.isna(x)


def load_rp2014_male_healthy_annuitant_qx_2014(rp2014_xlsx_path: str) -> MortalityTableQx:
    """
    Extract RP-2014 (base year 2014) one-year death probabilities q_x for:
    - Total Dataset; Males
    - Healthy Annuitant
    from the SOA workbook `research-2014-rp-mort-tab-rates-exposure.xlsx`.

    Output format matches this module:
    - ages: integer attained age
    - qx: annual one-year death probability
    """
    df = pd.read_excel(rp2014_xlsx_path, sheet_name="Total Dataset", header=None)

    header_row = None
    age_col = None
    healthy_col = None

    # Search for the header row that contains both "Age" and "Healthy Annuitant".
    for r in range(min(60, df.shape[0])):
        row = df.iloc[r, :].tolist()
        for c, v in enumerate(row):
            if isinstance(v, str) and v.strip() == "Age":
                age_col = c
            if isinstance(v, str) and v.strip() == "Healthy Annuitant":
                healthy_col = c
        if age_col is not None and healthy_col is not None:
            header_row = r
            break

    if header_row is None or age_col is None or healthy_col is None:
        raise ValueError("Could not locate 'Age' and 'Healthy Annuitant' columns in RP-2014 sheet.")

    ages: list[int] = []
    qxs: list[float] = []

    # Collect the first contiguous run of ages up to 120.
    for r in range(header_row + 1, df.shape[0]):
        age_val = df.iat[r, age_col]
        qx_val = df.iat[r, healthy_col]
        if _is_numeric(age_val) and _is_numeric(qx_val):
            age_int = int(age_val)
            qx = float(qx_val)
            ages.append(age_int)
            qxs.append(qx)
            if age_int >= 120:
                break

    if not ages:
        raise ValueError("No RP-2014 Healthy Annuitant Male qx values were extracted.")

    ages_arr = np.array(ages, dtype=int)
    qx_arr = np.array(qxs, dtype=float)

    order = np.argsort(ages_arr)
    return MortalityTableQx(ages=ages_arr[order], qx=qx_arr[order])


def ensure_rp2014_male_healthy_annuitant_qx_csv(
    *,
    rp2014_xlsx_path: str,
    out_csv_path: str,
) -> MortalityTableQx:
    if not pd.io.common.file_exists(out_csv_path):
        mt = load_rp2014_male_healthy_annuitant_qx_2014(rp2014_xlsx_path)
        pd.DataFrame({"age": mt.ages.astype(int), "qx": mt.qx}).to_csv(out_csv_path, index=False)
        return mt
    return MortalityTableQx.load_qx_csv(out_csv_path, age_col="age", qx_col="qx")


def load_mp2016_male_improvement_rates_multiplicative(
    mp2016_xlsx_path: str,
) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    """
    Extract MP-2016 (Male) improvement scale i(age, calendar_year).

    The workbook is arranged so that:
    - one row contains calendar years as column headers
    - one column contains attained ages as row labels
    - interior cells contain improvement values as decimals.

    Returns
    -------
    ages:
        1D array of integer ages.
    years:
        1D array of integer calendar years.
    i_matrix:
        2D array i_matrix[age_index, year_index] with annual improvement values.

    Notes
    -----
    This function treats the MP scale values as *log* mortality improvement rates.
    i.e., it uses: q(year_next) = q(year) * exp(i(age, year)).
    That is a common scaffold choice; if your downstream standard treats them
    as percentage rather than log-improvements, we can adjust.
    """
    df = pd.read_excel(mp2016_xlsx_path, sheet_name="Male", header=None)

    # Find the header row that contains calendar years in columns 1..end.
    years: list[int] = []
    year_cols: list[int] = []
    header_row = None

    for r in range(min(10, df.shape[0])):
        tmp_years = []
        tmp_cols = []
        for c in range(1, df.shape[1]):
            v = df.iat[r, c]
            if _is_numeric(v):
                iv = int(v)
                if 1950 <= iv <= 2100:
                    tmp_years.append(iv)
                    tmp_cols.append(c)
        if len(tmp_years) >= 20:
            years = tmp_years
            year_cols = tmp_cols
            header_row = r
            break

    if header_row is None or not years:
        raise ValueError("Could not locate MP-2016 year header row.")

    ages: list[int] = []
    i_rows: list[list[float]] = []

    for r in range(header_row + 1, df.shape[0]):
        age_val = df.iat[r, 0]
        if _is_numeric(age_val):
            age_int = int(age_val)
            row_rates: list[float] = []
            ok = True
            for c in year_cols:
                v = df.iat[r, c]
                if not _is_numeric(v):
                    ok = False
                    break
                row_rates.append(float(v))
            if ok:
                ages.append(age_int)
                i_rows.append(row_rates)

    if not ages:
        raise ValueError("No MP-2016 improvement rates were extracted.")

    ages_arr = np.array(ages, dtype=int)
    years_arr = np.array(years, dtype=int)
    i_matrix = np.array(i_rows, dtype=float)

    # Ensure ages are sorted (rows may already be ascending but be safe).
    order = np.argsort(ages_arr)
    return ages_arr[order], years_arr, i_matrix[order, :]


def ensure_mp2016_male_improvement_csv(
    *,
    mp2016_xlsx_path: str,
    out_csv_path: str,
) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    if pd.io.common.file_exists(out_csv_path):
        df = pd.read_csv(out_csv_path)
        ages = np.sort(df["age"].unique().astype(int))
        years = np.sort(df["year"].unique().astype(int))
        age_index = {a: i for i, a in enumerate(ages)}
        year_index = {y: j for j, y in enumerate(years)}
        i_matrix = np.zeros((len(ages), len(years)), dtype=float)
        for _, row in df.iterrows():
            i_matrix[age_index[int(row["age"])], year_index[int(row["year"])]] = float(row["improvement_rate"])
        return ages, years, i_matrix

    ages, years, i_matrix = load_mp2016_male_improvement_rates_multiplicative(mp2016_xlsx_path)
    rows = []
    for i, age in enumerate(ages):
        for j, year in enumerate(years):
            rows.append({"age": int(age), "year": int(years[j]), "improvement_rate": float(i_matrix[i, j])})
    pd.DataFrame(rows).to_csv(out_csv_path, index=False)
    return ages, years, i_matrix


@dataclass(frozen=True)
class MortalityTableRP2014MP2016:
    """
    Mortality model with:
    - RP-2014 base qx (Healthy Annuitant, Male, base year 2014)
    - MP-2016 improvement scale applied over calendar years (period mortality).

    This implementation computes annual qx(age, calendar_year) by:
    1. taking RP-2014 base q_x at integer attained age,
    2. applying cumulative MP factors i(age, y) for calendar years y from 2014 through
       min(calendar_year - 1, last year on the MP grid)—matching Excel SUMIFS over the
       published MP table without repeating the terminal column for later years,
    3. scaling: q_x = q_x(2014) * exp(sum of i).
    """

    base_qx_2014: MortalityTableQx
    mp2016_ages: np.ndarray
    mp2016_years: np.ndarray
    mp2016_i_matrix: np.ndarray  # i(age_index, year_index)
    base_year: int = 2014

    def _mp_i(self, age_int: int, calendar_year: int) -> float:
        # Clamp to available ranges (scaffold behavior).
        age_clamped = int(np.clip(age_int, int(self.mp2016_ages[0]), int(self.mp2016_ages[-1])))
        year_clamped = int(np.clip(calendar_year, int(self.mp2016_years[0]), int(self.mp2016_years[-1])))

        # Find nearest indices (years/ages are expected integer grid).
        age_idx = int(age_clamped - int(self.mp2016_ages[0]))
        year_idx = int(year_clamped - int(self.mp2016_years[0]))

        # In case the grid isn't contiguous, fall back to a search.
        if not np.array_equal(self.mp2016_ages, np.arange(int(self.mp2016_ages[0]), int(self.mp2016_ages[0]) + len(self.mp2016_ages))):
            age_idx = int(np.where(self.mp2016_ages == age_clamped)[0][0])
        if not np.array_equal(self.mp2016_years, np.arange(int(self.mp2016_years[0]), int(self.mp2016_years[0]) + len(self.mp2016_years))):
            year_idx = int(np.where(self.mp2016_years == year_clamped)[0][0])

        return float(self.mp2016_i_matrix[age_idx, year_idx])

    def qx_at_int_age_and_calendar_year(self, *, age_int: int, calendar_year: int) -> float:
        # Period-style improvement application (attained age x held constant):
        #   q_x(Y) = q_x(2014) * exp( sum_{y=2014}^{Y-1} i_x(y) )
        #
        # This avoids shifting x backwards when moving to later calendar years,
        # which can understate mortality if base rates are already indexed
        # by attained age.
        base_qx = self.base_qx_2014.qx_at_int_age(age_int)

        # Sum i(age, y) for y = base_year .. calendar_year-1, but only for years that exist
        # on the MP grid. Excel's SUMIFS(MP..., year,">=2014", year,"<"&E) sums table rows
        # and does not repeat the terminal column for later calendar years. Previously,
        # clamping inside `_mp_i` re-applied the last column for every y beyond the grid,
        # which overstated q_x when terminal rates are positive (long horizons vs Excel).
        y_last = int(self.mp2016_years[-1])
        last_y = min(int(calendar_year) - 1, y_last)
        cumulative_log = 0.0
        for y in range(self.base_year, last_y + 1):
            cumulative_log += self._mp_i(age_int, y)

        qx = base_qx * math.exp(cumulative_log)
        return float(min(max(qx, 0.0), 0.999999))

    def monthly_survival_to_payment(
        self,
        *,
        issue_age: int,
        n_months: int,
        valuation_year: int,
    ) -> np.ndarray:
        """
        Compute S(t_k)=P(T >= t_k) on a monthly grid where:
        - t=0 is 12/31/valuation_year
        - payments are at end of each month
        - mortality improvements vary by calendar year using MP-2016
        """
        if n_months <= 0:
            raise ValueError("n_months must be positive.")

        dt = 1.0 / 12.0
        S = np.ones(n_months, dtype=float)

        # Cache qx computations since many months share the same
        # (attained integer age, calendar-year) pair.
        qx_cache: dict[tuple[int, int], float] = {}

        log_S = 0.0
        for k in range(1, n_months + 1):
            m_index = k - 1  # 0 for first month interval
            age_start = issue_age + m_index * dt
            age_int = int(math.floor(age_start))

            # Since t=0 is 12/31/valuation_year, months in (0,1) are in valuation_year+1.
            calendar_year_start = valuation_year + 1 + (m_index // 12)

            key = (age_int, calendar_year_start)
            if key not in qx_cache:
                qx_cache[key] = self.qx_at_int_age_and_calendar_year(
                    age_int=age_int,
                    calendar_year=calendar_year_start,
                )
            qx = qx_cache[key]
            mu = -math.log(1.0 - qx)
            survival_interval = math.exp(-mu * dt)
            log_S += math.log(survival_interval)
            S[k - 1] = math.exp(log_S)

        return S


@dataclass(frozen=True)
class ExpenseAssumptions:
    """
    Expense assumptions for a single premium immediate annuity.

    Convention (used for this scaffold):
    - `policy_expense_dollars` is paid once at issue (t=0).
    - `premium_expense_rate` is applied to the single premium at issue time (t=0).
      The priced premium solves for the circularity created by this charge.
    - `monthly_expense_dollars` is paid at each monthly payment date while the life is
      alive (assumed to stop at death, aligned with benefit payment timing).
    """

    policy_expense_dollars: float
    premium_expense_rate: float
    monthly_expense_dollars: float

    @staticmethod
    def load_from_csv(
        path: str,
        *,
        key_col: str = "key",
        value_col: str = "value",
        unit_col: str = "unit",
    ) -> "ExpenseAssumptions":
        df = pd.read_csv(path)
        if key_col not in df.columns or value_col not in df.columns:
            raise ValueError(f"Expenses CSV must contain columns '{key_col}' and '{value_col}'.")

        def get_value(key: str) -> tuple[float, str]:
            row = df.loc[df[key_col] == key]
            if row.empty:
                raise ValueError(f"Missing expense key '{key}' in {path}.")
            value = float(row.iloc[0][value_col])
            unit = str(row.iloc[0][unit_col]) if unit_col in df.columns else ""
            return value, unit

        policy_expense, _ = get_value("policy_expense_dollars")
        premium_expense_rate, premium_unit = get_value("premium_expense_rate")
        monthly_expense, _ = get_value("monthly_expense_dollars")

        # Convert percent values if needed.
        premium_unit_norm = premium_unit.lower().strip()
        if premium_unit_norm in {"percent", "pct", "%"}:
            # Accept either "0.01" (already decimal) or "1.0" (one percent).
            premium_rate = premium_expense_rate / 100.0 if premium_expense_rate > 1.0 else premium_expense_rate
        else:
            premium_rate = premium_expense_rate

        if premium_rate < 0.0:
            raise ValueError("premium_expense_rate must be >= 0.")

        return ExpenseAssumptions(
            policy_expense_dollars=policy_expense,
            premium_expense_rate=premium_rate,
            monthly_expense_dollars=monthly_expense,
        )


@dataclass(frozen=True)
class SPIAContract:
    issue_age: int
    sex: Literal["male", "female"]
    benefit_annual: float  # annual analogue; first monthly accrual scales with index S1/S0, then return-indexed
    payment_freq_per_year: int = 12
    benefit_timing: Literal["end_of_period"] = "end_of_period"
    payment_cessation: Literal["at_death"] = "at_death"


@dataclass(frozen=True)
class SPIAProjectionResult:
    months: np.ndarray
    times_years: np.ndarray
    ages_at_payment: np.ndarray
    survival_to_payment: np.ndarray
    discount_factors: np.ndarray
    # PV and pricing
    pv_benefit: float  # PV of benefit cashflows only
    pv_monthly_expenses: float  # PV of monthly expense cashflows only
    annuity_factor: float  # sum_k P(alive at t_k)*DF(t_k) (level-$1 survival-weighted DF sum)
    single_premium: float  # priced premium including issue expense load
    # Expected cashflows (nominal, NOT discounted)
    expected_benefit_cashflows: np.ndarray
    expected_expense_cashflows: np.ndarray
    expected_total_cashflows: np.ndarray
    # Economic reserves: PV of remaining benefit+monthly expense at each time
    reserve_times_years: np.ndarray  # includes t=0 point
    economic_reserve: np.ndarray  # length n_months + 1
    # Index / inflation scaffolding (length n_months; aligned with payment months 1..N)
    index_level_at_payment: np.ndarray
    index_simple_return: np.ndarray  # S_k/S_{k-1} - 1 with S_{-1} := S_0 at issue
    index_log_return: np.ndarray
    index_cumulative_return: np.ndarray  # S_k / S_0 - 1
    benefit_nominal_scheduled: np.ndarray  # per-payment benefit if alive (before × survival in expected_*)
    expense_nominal_scheduled: np.ndarray
    expense_annual_inflation: float
    index_s0: float


@dataclass(frozen=True)
class SPIAMonteCarloResult:
    n_sims: int
    single_premium: np.ndarray
    pv_benefit: np.ndarray
    pv_monthly_expenses: np.ndarray
    pv_monthly_total: np.ndarray
    annuity_factor: np.ndarray
    premium_mean: float
    premium_median: float
    premium_p05: float
    premium_p95: float
    pv_benefit_mean: float
    pv_total_mean: float


def simulate_index_levels_gbm(
    *,
    n_sims: int,
    n_months: int,
    s0: float = 100.0,
    annual_drift: float = 0.06,
    annual_vol: float = 0.15,
    seed: int | None = None,
) -> np.ndarray:
    """
    Simulate monthly index levels under a geometric Brownian motion.

    Returns array shape (n_sims, n_months + 1) including month-0 level.
    """
    if n_sims < 1:
        raise ValueError("n_sims must be >= 1.")
    if n_months < 1:
        raise ValueError("n_months must be >= 1.")
    if s0 <= 0.0:
        raise ValueError("s0 must be > 0.")
    if annual_vol < 0.0:
        raise ValueError("annual_vol must be >= 0.")

    dt = 1.0 / 12.0
    rng = np.random.default_rng(seed)
    z = rng.normal(0.0, 1.0, size=(n_sims, n_months))
    drift = (annual_drift - 0.5 * annual_vol * annual_vol) * dt
    shock_scale = annual_vol * math.sqrt(dt)
    log_returns = drift + shock_scale * z
    cum_log = np.cumsum(log_returns, axis=1)

    levels = np.empty((n_sims, n_months + 1), dtype=float)
    levels[:, 0] = float(s0)
    levels[:, 1:] = float(s0) * np.exp(cum_log)
    return levels


def price_spia_single_premium(
    *,
    contract: SPIAContract,
    yield_curve: YieldCurve,
    mortality: MortalityTableQx | MortalityTableRP2014MP2016,
    horizon_age: int = 110,
    spread: float = 0.0,
    valuation_year: int | None = None,
    expenses: ExpenseAssumptions | None = None,
    expenses_csv_path: str = DEFAULT_EXPENSES_CSV,
    index_scenario_csv_path: str | None = None,
    index_s0: float | None = None,
    index_levels_payment: np.ndarray | None = None,
    expense_annual_inflation: float = 0.0,
) -> SPIAProjectionResult:
    """
    Price a SPIA with monthly benefits at month-end conditional on survival.

    Benefits grow by **return indexation** from monthly index levels in ``index_scenario_csv_path``
    (columns ``month``, ``sp500_level`` for months 0..N). If ``index_scenario_csv_path`` is None,
    the index is flat (zero equity returns) and benefits are level in nominal terms.

    Monthly maintenance expenses grow by ``expense_annual_inflation`` (e.g. 0.025 for 2.5%/year),
    compounded monthly, independent of the equity index.
    """
    if contract.payment_freq_per_year != 12:
        raise ValueError("This scaffold currently assumes monthly payments (12 per year).")

    # Monthly grid: t_k = k/12, stop when attained age >= horizon_age.
    dt = 1.0 / 12.0
    n_months = int(round((horizon_age - contract.issue_age) / dt))
    n_months = max(n_months, 1)

    months = np.arange(1, n_months + 1, dtype=int)
    times_years = months * dt
    ages_at_payment = contract.issue_age + times_years

    # Survival to payment: pay if alive at each payment date.
    if valuation_year is None and isinstance(mortality, MortalityTableRP2014MP2016):
        raise ValueError("valuation_year must be provided when using MortalityTableRP2014MP2016.")
    survival = mortality.monthly_survival_to_payment(
        issue_age=contract.issue_age,
        n_months=n_months,
        valuation_year=valuation_year,
    )

    # Discount factors.
    df = yield_curve.discount_factors(times_years, spread=spread)

    # Level monthly benefit amount (base before indexation).
    b_month = contract.benefit_annual / contract.payment_freq_per_year

    annuity_factor = float(np.sum(survival * df))

    # Load expenses if not provided.
    if expenses is None:
        try:
            expenses = ExpenseAssumptions.load_from_csv(expenses_csv_path)
        except (FileNotFoundError, ValueError):
            # If the expenses file is missing or malformed => assume zero expenses.
            expenses = ExpenseAssumptions(
                policy_expense_dollars=0.0,
                premium_expense_rate=0.0,
                monthly_expense_dollars=0.0,
            )

    monthly_expense_amount = float(expenses.monthly_expense_dollars)

    if index_levels_payment is not None:
        if index_scenario_csv_path is not None:
            raise ValueError("Provide either index_scenario_csv_path or index_levels_payment, not both.")
        if index_s0 is None:
            raise ValueError("index_s0 must be provided when index_levels_payment is provided.")
        levels_payment = np.asarray(index_levels_payment, dtype=float)
        if levels_payment.shape != (n_months,):
            raise ValueError(f"index_levels_payment must have shape ({n_months},).")
        if np.any(levels_payment <= 0.0) or not np.isfinite(index_s0):
            raise ValueError("index levels and index_s0 must be finite and strictly positive.")
        s0 = float(index_s0)
    elif index_scenario_csv_path is None:
        s0, levels_payment = flat_index_scenario(n_months)
    else:
        s0, levels_payment = load_index_scenario_monthly_csv(index_scenario_csv_path, n_months=n_months)

    ben_sched, exp_sched, simp_ret, log_ret, cumu_ret = _benefit_expense_and_index_returns(
        base_monthly=b_month,
        monthly_expense=monthly_expense_amount,
        s0=s0,
        levels_payment=levels_payment,
        expense_annual_inflation=float(expense_annual_inflation),
    )

    expected_benefit_cashflows = ben_sched * survival
    expected_expense_cashflows = exp_sched * survival
    expected_total_cashflows = expected_benefit_cashflows + expected_expense_cashflows

    pv_benefit = float(np.sum(expected_benefit_cashflows * df))
    pv_monthly_expenses = float(np.sum(expected_expense_cashflows * df))
    pv_monthly_total_outflows = pv_benefit + pv_monthly_expenses

    # Solve for the single premium under issue-time premium expense load:
    #   premium = policy_expense + pv_monthly_total_outflows + premium_expense_rate * premium
    # => premium * (1 - premium_expense_rate) = policy_expense + pv_monthly_total_outflows
    rate = float(expenses.premium_expense_rate)
    if rate >= 1.0:
        raise ValueError("premium_expense_rate must be < 1.")
    single_premium = float((float(expenses.policy_expense_dollars) + pv_monthly_total_outflows) / (1.0 - rate))

    # Economic reserve: after payment at t_{i+1}, roll forward PV of remaining expected nominal CFs.
    reserve_times_years = np.concatenate(([0.0], times_years))
    economic_reserve = np.zeros(n_months + 1, dtype=float)
    pv_remaining = np.zeros(n_months + 1, dtype=float)
    pv_remaining[n_months] = 0.0
    for i in range(n_months - 1, -1, -1):
        pv_remaining[i] = float(expected_total_cashflows[i] * df[i] + pv_remaining[i + 1])
    economic_reserve[0] = float(pv_remaining[0])

    for i in range(n_months):
        if i + 1 >= n_months:
            economic_reserve[i + 1] = 0.0
            continue
        if survival[i] <= 0.0 or df[i] <= 0.0:
            economic_reserve[i + 1] = 0.0
            continue
        future_pv_at_issue = float(pv_remaining[i + 1])
        economic_reserve[i + 1] = future_pv_at_issue / (survival[i] * df[i])

    return SPIAProjectionResult(
        months=months,
        times_years=times_years,
        ages_at_payment=ages_at_payment,
        survival_to_payment=survival,
        discount_factors=df,
        pv_benefit=pv_benefit,
        pv_monthly_expenses=pv_monthly_expenses,
        annuity_factor=annuity_factor,
        single_premium=single_premium,
        expected_benefit_cashflows=expected_benefit_cashflows,
        expected_expense_cashflows=expected_expense_cashflows,
        expected_total_cashflows=expected_total_cashflows,
        reserve_times_years=reserve_times_years,
        economic_reserve=economic_reserve,
        index_level_at_payment=levels_payment,
        index_simple_return=simp_ret,
        index_log_return=log_ret,
        index_cumulative_return=cumu_ret,
        benefit_nominal_scheduled=ben_sched,
        expense_nominal_scheduled=exp_sched,
        expense_annual_inflation=float(expense_annual_inflation),
        index_s0=float(s0),
    )


def price_spia_single_premium_monte_carlo(
    *,
    contract: SPIAContract,
    yield_curve: YieldCurve,
    mortality: MortalityTableQx | MortalityTableRP2014MP2016,
    horizon_age: int = 110,
    spread: float = 0.0,
    valuation_year: int | None = None,
    expenses: ExpenseAssumptions | None = None,
    expenses_csv_path: str = DEFAULT_EXPENSES_CSV,
    expense_annual_inflation: float = 0.0,
    n_sims: int = 1000,
    annual_drift: float = 0.06,
    annual_vol: float = 0.15,
    seed: int | None = None,
    s0: float = 100.0,
) -> SPIAMonteCarloResult:
    """Run Monte Carlo by simulating index paths and repricing vectorized across paths."""
    dt = 1.0 / 12.0
    n_months = int(round((horizon_age - contract.issue_age) / dt))
    n_months = max(n_months, 1)

    # Reuse deterministic ingredients once (survival, discounting, expenses).
    if valuation_year is None and isinstance(mortality, MortalityTableRP2014MP2016):
        raise ValueError("valuation_year must be provided when using MortalityTableRP2014MP2016.")

    survival = mortality.monthly_survival_to_payment(
        issue_age=contract.issue_age,
        n_months=n_months,
        valuation_year=valuation_year,  # ignored by MortalityTableQx
    )
    months = np.arange(1, n_months + 1, dtype=int)
    times_years = months * dt
    df = yield_curve.discount_factors(times_years, spread=spread)
    annuity_factor = float(np.sum(survival * df))

    if expenses is None:
        try:
            expenses = ExpenseAssumptions.load_from_csv(expenses_csv_path)
        except (FileNotFoundError, ValueError):
            expenses = ExpenseAssumptions(
                policy_expense_dollars=0.0,
                premium_expense_rate=0.0,
                monthly_expense_dollars=0.0,
            )
    rate = float(expenses.premium_expense_rate)
    if rate >= 1.0:
        raise ValueError("premium_expense_rate must be < 1.")

    g = monthly_rate_from_annual_inflation(float(expense_annual_inflation)) if expense_annual_inflation else 0.0
    expense_sched = float(expenses.monthly_expense_dollars) * (1.0 + g) ** np.arange(n_months, dtype=float)
    pv_monthly_expenses_single = float(np.sum(expense_sched * survival * df))

    idx_paths = simulate_index_levels_gbm(
        n_sims=n_sims,
        n_months=n_months,
        s0=s0,
        annual_drift=annual_drift,
        annual_vol=annual_vol,
        seed=seed,
    )

    # Benefits under return-indexation: b_k = base_monthly * S_k / S0.
    s0_eff = idx_paths[:, [0]]
    if np.any(s0_eff <= 0.0):
        raise ValueError("Simulated index levels at month 0 must be strictly positive.")
    base_monthly = float(contract.benefit_annual) / float(contract.payment_freq_per_year)
    benefit_sched = base_monthly * (idx_paths[:, 1:] / s0_eff)  # shape (n_sims, n_months)

    # PV benefit by path; deterministic monthly expense PV is shared across paths.
    weight = survival * df  # shape (n_months,)
    pvb = benefit_sched @ weight
    pve = np.full(n_sims, pv_monthly_expenses_single, dtype=float)
    pvt = pvb + pve

    numerator = float(expenses.policy_expense_dollars) + pvt
    prem = numerator / (1.0 - rate)
    af = np.full(n_sims, annuity_factor, dtype=float)

    return SPIAMonteCarloResult(
        n_sims=int(n_sims),
        single_premium=prem,
        pv_benefit=pvb,
        pv_monthly_expenses=pve,
        pv_monthly_total=pvt,
        annuity_factor=af,
        premium_mean=float(np.mean(prem)),
        premium_median=float(np.median(prem)),
        premium_p05=float(np.percentile(prem, 5.0)),
        premium_p95=float(np.percentile(prem, 95.0)),
        pv_benefit_mean=float(np.mean(pvb)),
        pv_total_mean=float(np.mean(pvt)),
    )


def _example_usage() -> None:
    """
    Example scaffold run.

    This uses:
    - a flat yield curve (placeholder)
    - an arbitrary synthetic mortality table (placeholder)

    Replace both with your sourced inputs for real pricing.
    """
    contract = SPIAContract(issue_age=65, sex="male", benefit_annual=100_000.0)
    # Prefer the seeded latest Treasury curve if available; otherwise fall back.
    try:
        yc = YieldCurve.load_zero_curve_csv(DEFAULT_ZERO_CURVE_CSV)
    except FileNotFoundError:
        yc = YieldCurve.from_flat_rate(0.04)

    # Mortality:
    # If you have SOA workbooks locally, use RP-2014 base + MP-2016 improvements.
    # Otherwise fall back to a static qx CSV or a synthetic placeholder.
    valuation_year = 2025
    try:
        rp_csv = DEFAULT_RP2014_MALE_HEALTHY_QX_CSV
        mp_csv = DEFAULT_MP2016_MALE_IMPROVEMENT_CSV
        base_qx = ensure_rp2014_male_healthy_annuitant_qx_csv(
            rp2014_xlsx_path=DEFAULT_RP2014_XLSX,
            out_csv_path=rp_csv,
        )
        mp_ages, mp_years, mp_i = ensure_mp2016_male_improvement_csv(
            mp2016_xlsx_path=DEFAULT_MP2016_XLSX,
            out_csv_path=mp_csv,
        )
        mortality = MortalityTableRP2014MP2016(
            base_qx_2014=base_qx,
            mp2016_ages=mp_ages,
            mp2016_years=mp_years,
            mp2016_i_matrix=mp_i,
            base_year=2014,
        )
    except FileNotFoundError:
        # Static qx CSV: columns age,qx
        try:
            mortality = MortalityTableQx.load_qx_csv(DEFAULT_MORTALITY_QX_CSV, age_col="age", qx_col="qx")
        except FileNotFoundError:
            # Synthetic qx: placeholder only.
            ages = np.arange(50, 121, dtype=int)
            qx = 0.0005 + 0.00002 * (ages - 50)  # placeholder only
            qx = np.clip(qx, 1e-9, 0.3)
            mortality = MortalityTableQx(ages=ages, qx=qx)

    res = price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mortality,
        horizon_age=110,
        spread=0.0,
        valuation_year=valuation_year,
    )
    print("Example SPIA single premium (placeholder inputs):", res.single_premium)
    print("  PV benefits:", res.pv_benefit)
    print("  PV monthly expenses:", res.pv_monthly_expenses)
    print("  Annuity factor:", res.annuity_factor)


if __name__ == "__main__":
    _example_usage()


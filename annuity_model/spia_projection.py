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

ALMReinvestRule = Literal["hold_cash", "pro_rata"]
ALMDisinvestRule = Literal["shortest_first", "pro_rata"]
ALMRebalancePolicy = Literal["full_target", "liquidity_only"]

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


# --- ALM (deterministic, single-scenario) ---------------------------------


def yield_curve_parallel_bps(yield_curve: YieldCurve, bps: float) -> YieldCurve:
    shift = float(bps) / 10000.0
    return YieldCurve(
        maturities_years=np.asarray(yield_curve.maturities_years, dtype=float).copy(),
        zero_rates=np.asarray(yield_curve.zero_rates, dtype=float).copy() + shift,
    )


def yield_curve_twist_linear_bps(
    yield_curve: YieldCurve,
    *,
    bps_short: float,
    bps_long: float,
    pivot_years: float = 5.0,
) -> YieldCurve:
    """
    Add a piecewise-linear twist in zero rates (in bps) across curve nodes.

    At maturity 0 (short end) the extra bump is ``bps_short``; at ``pivot_years`` the bump is the
    average of short and long; at and beyond ``max(longest node, pivot_years)`` the bump is ``bps_long``.
    Intermediate maturities interpolate linearly within each segment.
    """
    mats = np.asarray(yield_curve.maturities_years, dtype=float)
    if mats.size == 0:
        raise ValueError("yield_curve_twist_linear_bps requires a non-empty yield curve.")
    zeros = np.asarray(yield_curve.zero_rates, dtype=float).copy()
    if zeros.shape != mats.shape:
        raise ValueError("yield_curve maturities_years and zero_rates must have the same shape.")
    p = float(max(pivot_years, 1e-9))
    t_end = float(max(np.max(mats), p, 1e-9))
    mid_bps = 0.5 * (float(bps_short) + float(bps_long))

    def bump_bps_at_t(t: float) -> float:
        if t <= p:
            u = t / p
            return float(bps_short) * (1.0 - u) + mid_bps * u
        u = (t - p) / max(t_end - p, 1e-9)
        u = min(1.0, max(0.0, u))
        return mid_bps * (1.0 - u) + float(bps_long) * u

    for i, t in enumerate(mats):
        zeros[i] += bump_bps_at_t(float(t)) / 10000.0
    return YieldCurve(maturities_years=mats.copy(), zero_rates=zeros)


@dataclass(frozen=True)
class ALMBucketSpec:
    """One investable bucket: cash (tenor 0) or Treasury ZCB ladder slot."""

    name: str
    tenor_years: float  # 0 for cash


@dataclass(frozen=True)
class ALMAllocationSpec:
    """
    Target mix across buckets. ``weights`` aligns with ``buckets``; must sum to 1.
    """

    buckets: Tuple[ALMBucketSpec, ...]
    weights: np.ndarray

    def __post_init__(self) -> None:
        w = np.asarray(self.weights, dtype=float)
        if w.shape != (len(self.buckets),):
            raise ValueError("weights must have one entry per bucket.")
        if np.any(w < -1e-12) or abs(float(np.sum(w)) - 1.0) > 1e-6:
            raise ValueError("ALM weights must be non-negative and sum to 1.")


@dataclass(frozen=True)
class ALMAssumptions:
    allocation: ALMAllocationSpec
    rebalance_band: float
    rebalance_frequency_months: int
    reinvest_rule: ALMReinvestRule
    disinvest_rule: ALMDisinvestRule
    # "full_target": periodic drift-band rebalance back to target weights.
    # "liquidity_only": do not sell bonds to rebalance drift; only disinvest for cash shortfall.
    rebalance_policy: ALMRebalancePolicy = "full_target"
    # Liquidity buffer: cash + bond MV with residual maturity <= this (years) counts as "near-liquid".
    liquidity_near_liquid_years: float = 0.25

    def __post_init__(self) -> None:
        b = float(self.rebalance_band)
        if not math.isfinite(b) or b < 0.0 or b > 1.0:
            raise ValueError("rebalance_band must be finite and in [0, 1].")
        if int(self.rebalance_frequency_months) < 1:
            raise ValueError("rebalance_frequency_months must be >= 1.")
        if self.rebalance_policy not in ("full_target", "liquidity_only"):
            raise ValueError("rebalance_policy must be 'full_target' or 'liquidity_only'.")
        liq = float(self.liquidity_near_liquid_years)
        if not math.isfinite(liq) or liq < 0.0:
            raise ValueError("liquidity_near_liquid_years must be finite and non-negative.")


@dataclass(frozen=True)
class ALMResult:
    """Monthly ALM path (end-of-month after liability and trades)."""

    month_index: np.ndarray  # 0..n-1 paid months
    times_years: np.ndarray
    asset_market_value: np.ndarray
    liability_pv: np.ndarray
    funding_ratio: np.ndarray
    surplus: np.ndarray
    liquidity_buffer_months: np.ndarray
    bucket_asset_mv: np.ndarray  # shape (n_buckets, n_months)
    # Issue-time risk stats: PV01 is PV change from a +1bp parallel move to asset vs liability discounting
    pv01_assets: float
    pv01_liabilities: float
    pv01_net: float
    duration_assets_mac: float
    duration_liabilities_mac: float
    duration_gap: float


def alm_default_allocation_spec() -> ALMAllocationSpec:
    """5% cash + equal-weight Treasury ladder on the remainder (sums to 100%)."""
    w_cash = 0.05
    bond_names = ("1Y", "3Y", "5Y", "10Y", "20Y")
    bond_tenors = (1.0, 3.0, 5.0, 10.0, 20.0)
    n_b = len(bond_names)
    w_each = (1.0 - w_cash) / n_b
    buckets = (ALMBucketSpec("Cash", 0.0),) + tuple(
        ALMBucketSpec(nm, ty) for nm, ty in zip(bond_names, bond_tenors)
    )
    w = np.array([w_cash] + [w_each] * n_b, dtype=float)
    return ALMAllocationSpec(buckets=buckets, weights=w)


def liability_pv_after_paid_months(
    res: SPIAProjectionResult,
    yield_curve: YieldCurve,
    spread: float,
    last_paid_index: int,
    *,
    cashflows: np.ndarray | None = None,
) -> float:
    """
    PV at end of month ``last_paid_index`` (0-based CF index) of remaining expected outflows.

    ``last_paid_index = -1`` -> PV at issue (t=0) of all ``expected_total_cashflows``.
    """
    cf = np.asarray(res.expected_total_cashflows if cashflows is None else cashflows, dtype=float)
    ty = np.asarray(res.times_years, dtype=float)
    if cf.shape != ty.shape:
        raise ValueError("cashflows length must match pricing.times_years.")
    n = cf.shape[0]
    if last_paid_index >= n - 1:
        return 0.0
    if last_paid_index < 0:
        df = yield_curve.discount_factors(ty, spread=spread)
        return float(np.sum(cf * df))
    t_now = float(ty[last_paid_index])
    df_now = float(yield_curve.discount_factors(np.array([t_now]), spread=spread)[0])
    if df_now <= 0.0:
        return float("nan")
    j = slice(last_paid_index + 1, n)
    df_f = yield_curve.discount_factors(ty[j], spread=spread)
    return float(np.sum(cf[j] * (df_f / df_now)))


def _df_rem(yield_curve: YieldCurve, spread: float, t_rem_years: np.ndarray) -> np.ndarray:
    t = np.maximum(np.asarray(t_rem_years, dtype=float), 0.0)
    out = np.ones_like(t, dtype=float)
    mask = t > 1e-15
    if np.any(mask):
        out[mask] = yield_curve.discount_factors(t[mask], spread=spread)
    return out


def _liability_mac_duration_years(
    res: SPIAProjectionResult,
    yield_curve: YieldCurve,
    spread: float,
    *,
    cashflows: np.ndarray | None = None,
) -> float:
    cf = np.asarray(res.expected_total_cashflows if cashflows is None else cashflows, dtype=float)
    ty = np.asarray(res.times_years, dtype=float)
    if cf.shape != ty.shape:
        raise ValueError("cashflows length must match pricing.times_years.")
    df = yield_curve.discount_factors(ty, spread=spread)
    pv = float(np.sum(cf * df))
    if pv <= 0.0:
        return 0.0
    return float(np.sum(ty * cf * df) / pv)


def _alm_micro_reinvest_pro_rata(
    *,
    cash: float,
    faces: np.ndarray,
    t_rem: np.ndarray,
    w: np.ndarray,
    yield_curve: YieldCurve,
    spread: float,
    nominal_tenors: np.ndarray,
) -> tuple[float, np.ndarray]:
    """After maturity credits: move cash above target into bonds pro-rata to bond weights.

    Empty ladder slots (matured to cash) redeploy at each bucket's **nominal** tenor; otherwise
    principal would remain in cash with no working bond slot.
    """
    nb = faces.shape[0]
    nom = np.asarray(nominal_tenors, dtype=float)
    df = _df_rem(yield_curve, spread, t_rem)
    mv = faces * df
    aum = float(cash + np.sum(mv))
    if aum <= 0.0:
        return cash, faces
    w_b = np.asarray(w[1:], dtype=float)
    s = float(np.sum(w_b))
    if s <= 1e-15:
        return cash, faces
    w_b = w_b / s
    cash_tgt = float(w[0] * aum)
    excess = float(cash - cash_tgt)
    if excess <= 1e-6:
        return cash, faces
    faces = np.asarray(faces, dtype=float).copy()
    cash = float(cash)
    for k in range(nb):
        t_use = float(t_rem[k]) if t_rem[k] > 1e-14 else float(nom[k])
        if t_use <= 1e-14:
            continue
        dff = float(_df_rem(yield_curve, spread, np.array([t_use], dtype=float))[0])
        if dff <= 1e-15:
            continue
        d_mv = excess * float(w_b[k])
        faces[k] += d_mv / dff
        cash -= d_mv
        if float(t_rem[k]) <= 1e-14:
            t_rem[k] = float(nom[k])
    return cash, faces


def _alm_disinvest(
    *,
    cash: float,
    faces: np.ndarray,
    t_rem: np.ndarray,
    yield_curve: YieldCurve,
    spread: float,
    need: float,
    rule: ALMDisinvestRule,
) -> tuple[float, np.ndarray]:
    faces = np.asarray(faces, dtype=float).copy()
    cash = float(cash)
    need = float(need)
    if need <= 1e-9:
        return cash, faces
    df = _df_rem(yield_curve, spread, t_rem)
    nb = faces.shape[0]
    remaining = need

    if rule == "pro_rata":
        while remaining > 1e-6:
            df = _df_rem(yield_curve, spread, t_rem)
            mv = faces * df
            mv_tot = float(np.sum(mv))
            if mv_tot <= 1e-9:
                break
            for k in range(nb):
                if mv[k] <= 1e-12:
                    continue
                take = remaining * (mv[k] / mv_tot)
                dff = float(df[k])
                if dff <= 1e-15:
                    continue
                redu = min(float(faces[k]), take / dff)
                faces[k] -= redu
                cash += redu * dff
                remaining -= redu * dff
        return cash, faces

    # shortest_first: by remaining tenor then iterate
    order = np.argsort(t_rem + (faces <= 1e-15) * 1e6)
    for k in order:
        if remaining <= 1e-9:
            break
        if faces[k] <= 1e-15:
            continue
        dff = float(df[k])
        if dff <= 1e-15:
            continue
        redu = min(float(faces[k]), remaining / dff)
        faces[k] -= redu
        cash += redu * dff
        remaining -= redu * dff
        df = _df_rem(yield_curve, spread, t_rem)
    return cash, faces


def _alm_maybe_rebalance(
    *,
    cash: float,
    faces: np.ndarray,
    t_rem: np.ndarray,
    w: np.ndarray,
    yield_curve: YieldCurve,
    spread: float,
    band: float,
    month: int,
    freq: int,
    nominal_tenors: np.ndarray,
) -> tuple[float, np.ndarray]:
    df = _df_rem(yield_curve, spread, t_rem)
    mv = faces * df
    aum = float(cash + np.sum(mv))
    if aum <= 1e-9:
        return cash, faces
    w = np.asarray(w, dtype=float)
    tgt = w * aum
    tgt_mv_bonds = tgt[1:]
    act = np.concatenate(([cash], mv))
    drift = np.max(np.abs(act - tgt)) / aum
    if (month + 1) % freq != 0:
        return cash, faces
    if drift <= float(band):
        return cash, faces

    nom = np.asarray(nominal_tenors, dtype=float)
    faces = np.asarray(faces, dtype=float).copy()
    for k in range(faces.shape[0]):
        tgt_mv = float(tgt_mv_bonds[k])
        if t_rem[k] <= 1e-14:
            T = float(nom[k])
            if T <= 1e-14:
                faces[k] = 0.0
                continue
            dff = float(_df_rem(yield_curve, spread, np.array([T], dtype=float))[0])
            if dff <= 1e-15:
                faces[k] = 0.0
                continue
            faces[k] = tgt_mv / dff
            t_rem[k] = T
            continue
        dff = float(_df_rem(yield_curve, spread, np.array([t_rem[k]], dtype=float))[0])
        if dff <= 1e-15:
            faces[k] = 0.0
            t_rem[k] = 0.0
            continue
        faces[k] = tgt_mv / dff
    cash = float(aum - np.sum(faces * _df_rem(yield_curve, spread, t_rem)))
    return cash, faces


def run_alm_projection(
    *,
    pricing: SPIAProjectionResult,
    yield_curve: YieldCurve,
    spread: float,
    assumptions: ALMAssumptions,
    initial_asset_market_value: float | None = None,
    asset_curve: YieldCurve | None = None,
    liability_curve: YieldCurve | None = None,
    liability_cashflows: np.ndarray | None = None,
) -> ALMResult:
    """
    Roll forward Treasury ladder + cash against ``expected_total_cashflows``.

    Optional ``asset_curve`` / ``liability_curve`` split discounting for bonds vs liability PV
    (default: both use ``yield_curve``). Mark-to-market and liability PV include ``spread``.

    ``liability_cashflows`` overrides ``pricing.expected_total_cashflows`` when provided (same length).
    """
    if initial_asset_market_value is None:
        initial_asset_market_value = float(pricing.single_premium)
    aum0 = float(initial_asset_market_value)
    if aum0 <= 0.0:
        raise ValueError("initial_asset_market_value must be positive.")

    yc_a = asset_curve if asset_curve is not None else yield_curve
    yc_l = liability_curve if liability_curve is not None else yield_curve

    spec = assumptions.allocation
    buckets = spec.buckets
    w = np.asarray(spec.weights, dtype=float)
    n_b = len(buckets) - 1
    if n_b < 1:
        raise ValueError("Need at least one bond bucket besides cash.")
    dt = 1.0 / 12.0

    faces = np.zeros(n_b, dtype=float)
    t_rem = np.zeros(n_b, dtype=float)
    for k in range(n_b):
        ten = float(buckets[k + 1].tenor_years)
        t_rem[k] = max(ten, 0.0)
        mv_t = float(w[k + 1] * aum0)
        if t_rem[k] <= 1e-14:
            continue
        d0 = float(yc_a.discount_factors(np.array([t_rem[k]]), spread=spread)[0])
        if d0 <= 1e-15:
            raise ValueError("Discount factor too small for ALM initialization.")
        faces[k] = mv_t / d0

    cash = float(w[0] * aum0)
    cf = np.asarray(
        pricing.expected_total_cashflows if liability_cashflows is None else liability_cashflows,
        dtype=float,
    )
    if cf.shape != (pricing.expected_total_cashflows.shape[0],):
        raise ValueError("liability_cashflows must match pricing horizon length.")
    ty_full = np.asarray(pricing.times_years, dtype=float)
    init_cash = float(cash)
    init_faces = np.asarray(faces, dtype=float).copy()
    init_t_rem = np.asarray(t_rem, dtype=float).copy()
    n = cf.shape[0]
    band = float(assumptions.rebalance_band)
    freq = max(1, int(assumptions.rebalance_frequency_months))
    near_liq = float(assumptions.liquidity_near_liquid_years)
    nominal_tenors = np.array([float(buckets[k + 1].tenor_years) for k in range(n_b)], dtype=float)

    asset_mv = np.zeros(n, dtype=float)
    liab_pv = np.zeros(n, dtype=float)
    fr = np.zeros(n, dtype=float)
    surp = np.zeros(n, dtype=float)
    liq_buf = np.zeros(n, dtype=float)
    bucket_hist = np.zeros((len(buckets), n), dtype=float)

    for m in range(n):
        # 1) accrue & mature
        t_rem = np.maximum(t_rem - dt, 0.0)
        df = _df_rem(yc_a, spread, t_rem)
        matured = np.where((faces > 1e-15) & (t_rem <= 1e-14), True, False)
        for k in np.flatnonzero(matured):
            cash += float(faces[k])
            faces[k] = 0.0
        t_rem = np.where(faces <= 1e-15, 0.0, t_rem)
        df = _df_rem(yc_a, spread, t_rem)
        mv = faces * df

        if assumptions.reinvest_rule == "pro_rata" and bool(np.any(matured)):
            cash, faces = _alm_micro_reinvest_pro_rata(
                cash=cash,
                faces=faces,
                t_rem=t_rem,
                w=w,
                yield_curve=yc_a,
                spread=spread,
                nominal_tenors=nominal_tenors,
            )
            df = _df_rem(yc_a, spread, t_rem)
            mv = faces * df

        # 2) liability cashflow
        cash -= float(cf[m])

        need = max(0.0, -cash)
        if need > 1e-9:
            cash, faces = _alm_disinvest(
                cash=cash,
                faces=faces,
                t_rem=t_rem,
                yield_curve=yc_a,
                spread=spread,
                need=need,
                rule=assumptions.disinvest_rule,
            )
            df = _df_rem(yc_a, spread, t_rem)
            mv = faces * df

        if assumptions.rebalance_policy == "full_target":
            cash, faces = _alm_maybe_rebalance(
                cash=cash,
                faces=faces,
                t_rem=t_rem,
                w=w,
                yield_curve=yc_a,
                spread=spread,
                band=band,
                month=m,
                freq=freq,
                nominal_tenors=nominal_tenors,
            )
            df = _df_rem(yc_a, spread, t_rem)
            mv = faces * df
        aum_end = float(cash + np.sum(mv))

        L = liability_pv_after_paid_months(pricing, yc_l, spread, m, cashflows=cf)
        asset_mv[m] = aum_end
        liab_pv[m] = L
        if L > 1e-9:
            fr[m] = aum_end / L
        elif aum_end > 0.0:
            fr[m] = float("inf")
        else:
            fr[m] = 0.0
        surp[m] = aum_end - L

        liquid = float(cash)
        for k in range(n_b):
            if t_rem[k] <= near_liq + 1e-12:
                liquid += float(mv[k])
        out_next12 = float(np.mean(cf[m : min(m + 12, n)])) if m < n else 0.0
        liq_buf[m] = liquid / max(out_next12, 1e-6)

        hist_row = np.concatenate(([cash], mv))
        bucket_hist[:, m] = hist_row

    yc_a_b = yield_curve_parallel_bps(yc_a, 1.0)
    yc_l_b = yield_curve_parallel_bps(yc_l, 1.0)
    L0 = liability_pv_after_paid_months(pricing, yc_l, spread, -1, cashflows=cf)
    L0b = liability_pv_after_paid_months(pricing, yc_l_b, spread, -1, cashflows=cf)
    pv01_liab = float(L0b - L0)

    Ab = float(init_cash + np.sum(init_faces * _df_rem(yc_a_b, spread, init_t_rem)))
    A0 = float(init_cash + np.sum(init_faces * _df_rem(yc_a, spread, init_t_rem)))
    pv01_assets = float(Ab - A0)
    pv01_net = float(pv01_assets - pv01_liab)

    d_l = _liability_mac_duration_years(pricing, yc_l, spread, cashflows=cf)
    df_i = _df_rem(yc_a, spread, init_t_rem)
    mv_i = init_faces * df_i
    aum_i = float(init_cash + np.sum(mv_i))
    d_a = float(np.sum(mv_i * init_t_rem) / aum_i) if aum_i > 1e-9 else 0.0
    d_gap = float(d_a - d_l)

    return ALMResult(
        month_index=np.arange(n, dtype=int),
        times_years=ty_full,
        asset_market_value=asset_mv,
        liability_pv=liab_pv,
        funding_ratio=fr,
        surplus=surp,
        liquidity_buffer_months=liq_buf,
        bucket_asset_mv=bucket_hist,
        pv01_assets=pv01_assets,
        pv01_liabilities=pv01_liab,
        pv01_net=pv01_net,
        duration_assets_mac=d_a,
        duration_liabilities_mac=d_l,
        duration_gap=d_gap,
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


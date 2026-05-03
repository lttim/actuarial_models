from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np

import pricing_projection as sp
from _observability import traced
from policy_features import (
    GLWBRider,
    MonthlySchedule,
    SegmentAllocation,
    SurrenderChargeSchedule,
    normalize_segment_allocations,
)
from policy_features import (
    segment_credited_return as policy_segment_credited_return,
)


@dataclass(frozen=True)
class RILAContract:
    """Registered index-linked annuity — accumulation, annual point-to-point crediting (v1)."""

    issue_age: int
    sex: Literal["male", "female"]
    participation: float
    cap: float
    floor: float
    rider_fee_annual: float
    single_premium: float | None = None
    segment_allocations: tuple[SegmentAllocation, ...] = ()
    withdrawals: MonthlySchedule = MonthlySchedule()
    surrender_charges: SurrenderChargeSchedule = SurrenderChargeSchedule()
    death_benefit_type: Literal["account_value", "return_of_premium"] = "account_value"
    glwb: GLWBRider = GLWBRider()
    segment_months: int = 12
    payment_freq_per_year: int = 12

    @property
    def benefit_annual(self) -> float:
        """UI / export compatibility (no SPIA-style annual benefit)."""
        return 0.0


@dataclass(frozen=True)
class RILAProjectionResult:
    months: np.ndarray
    times_years: np.ndarray
    ages_at_payment: np.ndarray
    survival_to_payment: np.ndarray
    discount_factors: np.ndarray
    pv_benefit: float
    pv_monthly_expenses: float
    annuity_factor: float
    single_premium: float
    expected_benefit_cashflows: np.ndarray
    expected_expense_cashflows: np.ndarray
    expected_total_cashflows: np.ndarray
    reserve_times_years: np.ndarray
    economic_reserve: np.ndarray
    index_level_at_payment: np.ndarray
    index_simple_return: np.ndarray
    index_log_return: np.ndarray
    index_cumulative_return: np.ndarray
    benefit_nominal_scheduled: np.ndarray
    expense_nominal_scheduled: np.ndarray
    expense_annual_inflation: float
    index_s0: float
    account_value_end_month: np.ndarray
    segment_credited_rate: np.ndarray
    expected_claim_cashflows: np.ndarray
    withdrawal_cashflows: np.ndarray
    surrender_charge_dollars: np.ndarray
    surrender_value_end_month: np.ndarray
    benefit_base_end_month: np.ndarray
    glwb_withdrawal_cashflows: np.ndarray
    rider_fee_cashflows: np.ndarray


class RILAPricingInfeasibleError(ValueError):
    """
    Raised when the implicit equation for single premium has no positive solution.

    Premium P must satisfy P = (policy_expense + PV(expenses)) / (1 - premium_expense_rate - K)
    where K is PV(death benefits | initial AV = 1$) per dollar of premium. Feasibility requires
    K + premium_expense_rate < 1.
    """

    def __init__(
        self,
        *,
        k_loading: float,
        premium_expense_rate: float,
        detail: str = "",
    ) -> None:
        self.k_loading = float(k_loading)
        self.premium_expense_rate = float(premium_expense_rate)
        s = self.k_loading + self.premium_expense_rate
        msg = (
            "RILA single-premium pricing is infeasible: the present value of death benefits per "
            "$1 of account value loaded into the premium (K) plus the premium expense rate must "
            f"stay below 1. Here K={self.k_loading:.6f}, premium_expense_rate={self.premium_expense_rate:.6f}, "
            f"sum={s:.6f}. Try lower participation or cap, a less bullish index scenario, lower "
            "rider fee, a shorter horizon, or a lower premium expense rate."
        )
        if detail:
            msg = f"{msg} {detail}"
        super().__init__(msg)


@dataclass(frozen=True)
class RILAMonteCarloResult:
    """
    Monte Carlo pricing summary across simulated index paths.

    Per-path arrays (``single_premium``, ``pv_benefit``, ``pv_monthly_expenses``, ``pv_monthly_total``,
    ``annuity_factor``) carry **NaN** for paths that hit
    :class:`RILAPricingInfeasibleError` (i.e., \\(K + r \\geq 1\\) on that path) when
    ``infeasible_path_policy='skip'``. Aggregate statistics are NaN-aware and computed only
    over feasible paths. ``n_feasible + n_infeasible == n_sims``.
    """

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
    n_feasible: int = 0
    n_infeasible: int = 0
    infeasible_max_loading: float = 0.0


def segment_credited_return(*, raw: float, participation: float, cap: float, floor: float) -> float:
    """Per-segment crediting for RILA's annual point-to-point.

    Public name preserved for back-compat (RILA golden JSON depends on
    byte-identical behavior); internally delegates to
    :class:`crediting.AnnualPointToPointCapped` so the strategy is
    shared with FIA / IUL (Section 1.2 of
    ``docs/seven_product_rollout_plan.md``).
    """
    from crediting import AnnualPointToPointCapped

    strategy = AnnualPointToPointCapped(
        participation=float(participation),
        cap=float(cap),
        floor=float(floor),
    )
    return float(strategy.credit_segment(raw_index_return=float(raw)))


def levels_end_by_policy_month(*, s0: float, levels_payment: np.ndarray) -> np.ndarray:
    """L[j] = index at end of policy month j, j=0..n_months."""
    n_months = int(levels_payment.shape[0])
    L = np.zeros(n_months + 1, dtype=float)
    L[0] = float(s0)
    L[1:] = np.asarray(levels_payment, dtype=float)
    return L


def _rila_claims_rel_per_premium_dollar(
    *,
    contract: RILAContract,
    L: np.ndarray,
    death_prob: np.ndarray,
) -> tuple[
    np.ndarray, np.ndarray, np.ndarray, np.ndarray, np.ndarray, np.ndarray, np.ndarray, np.ndarray
]:
    """
    Relative simulation with initial AV=1.

    Returns (claim_rate, av_end, cred_rate_month, withdrawals, surrender_charges,
    surrender_value, benefit_base, glwb_withdrawals).

    cred_rate_month[k] holds annual segment **credited decimal** applied at end of month k+1
    if that month is a crediting anniversary; else 0.
    """
    n_months = int(death_prob.shape[0])
    seg = int(contract.segment_months)
    if seg != 12:
        raise ValueError("RILA v1 only supports segment_months=12.")
    if L.shape[0] != n_months + 1:
        raise ValueError("L must have shape (n_months + 1,).")

    basis = float(contract.single_premium) if contract.single_premium is not None else 1.0
    av = basis
    initial_av = basis
    benefit_base = basis
    claims = np.zeros(n_months, dtype=float)
    av_end = np.zeros(n_months, dtype=float)
    cred_m = np.zeros(n_months, dtype=float)
    withdrawals = np.zeros(n_months, dtype=float)
    surrender_charges = np.zeros(n_months, dtype=float)
    surrender_value = np.zeros(n_months, dtype=float)
    benefit_base_path = np.zeros(n_months, dtype=float)
    glwb_withdrawals = np.zeros(n_months, dtype=float)
    fee_m = float(contract.rider_fee_annual) / 12.0
    if not (0.0 <= fee_m < 1.0):
        raise ValueError("rider_fee_annual must imply monthly fee in [0,1).")
    glwb_fee_m = float(contract.glwb.fee_annual) / 12.0 if contract.glwb.enabled else 0.0
    glwb_roll_m = (
        (1.0 + float(contract.glwb.rollup_annual)) ** (1.0 / 12.0) - 1.0
        if contract.glwb.enabled
        else 0.0
    )
    w_sched = contract.withdrawals.values(n_months)
    surrender_rates = contract.surrender_charges.monthly_rates(n_months)
    allocations = normalize_segment_allocations(
        contract.segment_allocations
        or (
            SegmentAllocation(
                weight=1.0,
                design="cap_floor",
                participation=float(contract.participation),
                cap=float(contract.cap),
                floor=float(contract.floor),
            ),
        )
    )

    for m in range(1, n_months + 1):
        if m >= seg and (m % seg) == 0:
            raw = float(L[m] / L[m - seg] - 1.0)
            cr = float(
                sum(
                    float(a.weight)
                    * policy_segment_credited_return(allocation=a, raw_index_return=raw)
                    for a in allocations
                )
            )
            cred_m[m - 1] = cr
            av *= 1.0 + cr
        if contract.glwb.enabled and m < int(contract.glwb.income_start_month):
            benefit_base *= 1.0 + glwb_roll_m
            if contract.glwb.ratchet and m % 12 == 0:
                benefit_base = max(benefit_base, av)
        if contract.glwb.enabled and m >= int(contract.glwb.income_start_month):
            gwb = min(av, benefit_base * float(contract.glwb.withdrawal_rate) / 12.0)
            av -= gwb
            glwb_withdrawals[m - 1] = gwb
        wd = min(av, float(w_sched[m - 1]))
        av -= wd
        withdrawals[m - 1] = wd
        av *= 1.0 - fee_m
        if glwb_fee_m > 0.0:
            fee_base = max(0.0, benefit_base)
            av = max(0.0, av - fee_base * glwb_fee_m)
        if contract.death_benefit_type == "return_of_premium":
            db = max(av, initial_av)
        else:
            db = av
        claims[m - 1] = float(death_prob[m - 1] * db)
        av_end[m - 1] = float(av)
        surrender_charges[m - 1] = float(av * surrender_rates[m - 1])
        surrender_value[m - 1] = float(max(0.0, av - surrender_charges[m - 1]))
        benefit_base_path[m - 1] = float(benefit_base)
    return (
        claims,
        av_end,
        cred_m,
        withdrawals,
        surrender_charges,
        surrender_value,
        benefit_base_path,
        glwb_withdrawals,
    )


@traced("pricing.rila.deterministic")
def price_rila_single_premium(
    *,
    contract: RILAContract,
    yield_curve: sp.YieldCurve,
    mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
    horizon_age: int = 110,
    spread: float = 0.0,
    valuation_year: int | None = None,
    expenses: sp.ExpenseAssumptions | None = None,
    expenses_csv_path: str = sp.DEFAULT_EXPENSES_CSV,
    index_scenario_csv_path: str | None = None,
    index_s0: float | None = None,
    index_levels_payment: np.ndarray | None = None,
    expense_annual_inflation: float = 0.0,
) -> RILAProjectionResult:
    if contract.payment_freq_per_year != 12:
        raise ValueError("RILA scaffold assumes monthly frequency.")
    if contract.participation < 0.0:
        raise ValueError("participation must be non-negative.")
    if contract.cap < contract.floor:
        raise ValueError("cap must be >= floor.")
    if not (0.0 <= contract.rider_fee_annual <= 1.0):
        raise ValueError("rider_fee_annual must be in [0, 1].")

    dt = 1.0 / 12.0
    n_months = int(round((horizon_age - contract.issue_age) / dt))
    n_months = max(n_months, 1)

    months = np.arange(1, n_months + 1, dtype=int)
    times_years = months * dt
    ages_at_payment = contract.issue_age + times_years

    if valuation_year is None and isinstance(mortality, sp.MortalityTableRP2014MP2016):
        raise ValueError("valuation_year must be provided when using MortalityTableRP2014MP2016.")

    survival = mortality.monthly_survival_to_payment(
        issue_age=contract.issue_age,
        n_months=n_months,
        valuation_year=valuation_year,
    )
    survival_start = np.empty_like(survival)
    survival_start[0] = 1.0
    survival_start[1:] = survival[:-1]
    death_prob = np.clip(survival_start - survival, 0.0, 1.0)

    df = yield_curve.discount_factors(times_years, spread=spread)
    annuity_factor = float(np.sum(survival * df))

    if expenses is None:
        try:
            expenses = sp.ExpenseAssumptions.load_from_csv(expenses_csv_path)
        except (FileNotFoundError, ValueError):
            expenses = sp.ExpenseAssumptions(0.0, 0.0, 0.0)

    if index_levels_payment is not None:
        if index_scenario_csv_path is not None:
            raise ValueError(
                "Provide either index_scenario_csv_path or index_levels_payment, not both."
            )
        if index_s0 is None:
            raise ValueError("index_s0 must be provided when index_levels_payment is provided.")
        levels_payment = np.asarray(index_levels_payment, dtype=float)
        if levels_payment.shape != (n_months,):
            raise ValueError(f"index_levels_payment must have shape ({n_months},).")
        if (
            np.any(levels_payment <= 0.0)
            or np.any(~np.isfinite(levels_payment))
            or not np.isfinite(index_s0)
            or float(index_s0) <= 0.0
        ):
            raise ValueError("index levels and index_s0 must be finite and strictly positive.")
        s0 = float(index_s0)
    elif index_scenario_csv_path is None:
        s0, levels_payment = sp.flat_index_scenario(n_months)
    else:
        s0, levels_payment = sp.load_index_scenario_monthly_csv(
            index_scenario_csv_path, n_months=n_months
        )

    L = levels_end_by_policy_month(s0=s0, levels_payment=levels_payment)
    (
        claims_rel,
        av_end_rel,
        cred_m,
        withdrawal_rel,
        surrender_charge_rel,
        surrender_value_rel,
        benefit_base_rel,
        glwb_withdrawal_rel,
    ) = _rila_claims_rel_per_premium_dollar(
        contract=contract,
        L=L,
        death_prob=death_prob,
    )

    g = (
        sp.monthly_rate_from_annual_inflation(float(expense_annual_inflation))
        if expense_annual_inflation
        else 0.0
    )
    expense_sched = float(expenses.monthly_expense_dollars) * (1.0 + g) ** np.arange(
        n_months, dtype=float
    )
    expected_expense_cashflows = expense_sched * survival

    access_rel = (withdrawal_rel + glwb_withdrawal_rel) * survival_start
    K = float(np.sum((claims_rel + access_rel) * df))
    pv_exp_sched = float(np.sum(expected_expense_cashflows * df))
    rate = float(expenses.premium_expense_rate)
    if rate >= 1.0:
        raise ValueError("premium_expense_rate must be < 1.")
    if contract.single_premium is None:
        denom = 1.0 - rate - K
        if denom <= 1e-12:
            raise RILAPricingInfeasibleError(k_loading=float(K), premium_expense_rate=float(rate))
        single_premium = float((float(expenses.policy_expense_dollars) + pv_exp_sched) / denom)
    else:
        single_premium = float(contract.single_premium)
        if not np.isfinite(single_premium) or single_premium <= 0.0:
            raise ValueError("single_premium must be finite and > 0 when provided.")
    if contract.single_premium is None and (
        not np.isfinite(single_premium) or single_premium <= 0.0
    ):
        raise ValueError(
            "RILA priced single premium is non-positive. With the current implicit premium "
            "formula, the numerator is policy expenses plus the PV of scheduled monthly "
            "expenses; if both are zero (and there is no premium expense loading that "
            "forces a positive premium), the closed-form premium collapses to zero and "
            "all scaled cashflows vanish. Load non-zero :class:`~pricing_projection.ExpenseAssumptions` "
            "(for example via ``ExpenseAssumptions.load_from_csv(pricing_projection.DEFAULT_EXPENSES_CSV)``) "
            "or pass explicit positive policy / monthly expense dollars."
        )

    scale = 1.0 if contract.single_premium is not None else single_premium
    expected_access_cashflows = access_rel * scale
    expected_claim_cashflows = claims_rel * scale
    expected_benefit_cashflows = expected_claim_cashflows + expected_access_cashflows
    expected_total_cashflows = expected_benefit_cashflows + expected_expense_cashflows

    pv_benefit = float(np.sum(expected_benefit_cashflows * df))
    pv_monthly_expenses = float(np.sum(expected_expense_cashflows * df))

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
        economic_reserve[i + 1] = float(pv_remaining[i + 1] / (survival[i] * df[i]))

    simp_ret = np.zeros(n_months, dtype=float)
    log_ret = np.zeros(n_months, dtype=float)
    cumu_ret = np.zeros(n_months, dtype=float)
    prev = float(s0)
    for k in range(n_months):
        cur = float(levels_payment[k])
        simp_ret[k] = cur / prev - 1.0
        log_ret[k] = np.log(cur / prev) if cur > 0.0 and prev > 0.0 else float("nan")
        cumu_ret[k] = cur / float(s0) - 1.0
        prev = cur

    account_value_end_month = av_end_rel * scale
    withdrawal_cashflows = withdrawal_rel * scale * survival_start
    glwb_withdrawal_cashflows = glwb_withdrawal_rel * scale * survival_start
    surrender_charge_dollars = surrender_charge_rel * scale
    surrender_value_end_month = surrender_value_rel * scale
    benefit_base_end_month = benefit_base_rel * scale
    rider_fee_cashflows = (
        benefit_base_rel
        * scale
        * (float(contract.glwb.fee_annual) / 12.0 if contract.glwb.enabled else 0.0)
        * survival_start
    )

    return RILAProjectionResult(
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
        benefit_nominal_scheduled=expected_benefit_cashflows,
        expense_nominal_scheduled=expected_expense_cashflows,
        expense_annual_inflation=float(expense_annual_inflation),
        index_s0=float(s0),
        account_value_end_month=account_value_end_month,
        segment_credited_rate=cred_m,
        expected_claim_cashflows=expected_claim_cashflows,
        withdrawal_cashflows=withdrawal_cashflows,
        surrender_charge_dollars=surrender_charge_dollars,
        surrender_value_end_month=surrender_value_end_month,
        benefit_base_end_month=benefit_base_end_month,
        glwb_withdrawal_cashflows=glwb_withdrawal_cashflows,
        rider_fee_cashflows=rider_fee_cashflows,
    )


def liability_path_from_rila_projection(pricing: RILAProjectionResult) -> sp.LiabilityPath:
    return sp.LiabilityPath(
        times_years=np.asarray(pricing.times_years, dtype=float),
        expected_total_cashflows=np.asarray(pricing.expected_total_cashflows, dtype=float),
    )


# Register with the liability-path dispatch so the engine core can route
# `run_alm_projection_from_pricing_result` without an isinstance chain.
from liability_dispatch import register_liability_path_converter  # noqa: E402

register_liability_path_converter("RILAProjectionResult", liability_path_from_rila_projection)


@traced("pricing.rila.monte_carlo")
def price_rila_single_premium_monte_carlo(
    *,
    contract: RILAContract,
    yield_curve: sp.YieldCurve,
    mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
    horizon_age: int = 110,
    spread: float = 0.0,
    valuation_year: int | None = None,
    expenses: sp.ExpenseAssumptions | None = None,
    expenses_csv_path: str = sp.DEFAULT_EXPENSES_CSV,
    expense_annual_inflation: float = 0.0,
    n_sims: int = 500,
    annual_drift: float = 0.06,
    annual_vol: float = 0.15,
    seed: int | None = None,
    s0: float = 100.0,
    infeasible_path_policy: Literal["skip", "raise"] = "skip",
) -> RILAMonteCarloResult:
    """
    Monte Carlo price RILA across simulated GBM index paths.

    ``infeasible_path_policy``:
      - ``"skip"`` (default): paths where the implicit single-premium denominator collapses
        (\\(K + r \\geq 1\\)) are recorded as ``NaN`` and counted in
        ``RILAMonteCarloResult.n_infeasible``. Aggregate stats are NaN-aware. If **all** paths
        fail, the worst-case :class:`RILAPricingInfeasibleError` is raised.
      - ``"raise"``: re-raise the first :class:`RILAPricingInfeasibleError` (legacy behavior).
    """
    dt = 1.0 / 12.0
    n_months = int(round((horizon_age - contract.issue_age) / dt))
    n_months = max(n_months, 1)

    if valuation_year is None and isinstance(mortality, sp.MortalityTableRP2014MP2016):
        raise ValueError("valuation_year must be provided when using MortalityTableRP2014MP2016.")

    survival = mortality.monthly_survival_to_payment(
        issue_age=contract.issue_age,
        n_months=n_months,
        valuation_year=valuation_year,
    )
    months = np.arange(1, n_months + 1, dtype=int)
    times_years = months * dt
    df = yield_curve.discount_factors(times_years, spread=spread)
    annuity_factor = float(np.sum(survival * df))

    if expenses is None:
        try:
            expenses = sp.ExpenseAssumptions.load_from_csv(expenses_csv_path)
        except (FileNotFoundError, ValueError):
            expenses = sp.ExpenseAssumptions(0.0, 0.0, 0.0)
    rate = float(expenses.premium_expense_rate)
    if rate >= 1.0:
        raise ValueError("premium_expense_rate must be < 1.")

    g = (
        sp.monthly_rate_from_annual_inflation(float(expense_annual_inflation))
        if expense_annual_inflation
        else 0.0
    )
    expense_sched = float(expenses.monthly_expense_dollars) * (1.0 + g) ** np.arange(
        n_months, dtype=float
    )
    expected_expense_cashflows = expense_sched * survival
    pv_monthly_expenses_single = float(np.sum(expected_expense_cashflows * df))

    idx_paths = sp.simulate_index_levels_gbm(
        n_sims=n_sims,
        n_months=n_months,
        s0=s0,
        annual_drift=annual_drift,
        annual_vol=annual_vol,
        seed=seed,
    )

    prem = np.full(n_sims, np.nan, dtype=float)
    pvb = np.full(n_sims, np.nan, dtype=float)
    n_infeasible = 0
    n_feasible = 0
    worst_exc: RILAPricingInfeasibleError | None = None
    worst_loading = 0.0

    for i in range(int(n_sims)):
        path = idx_paths[i, :]
        if path.shape[0] != n_months + 1:
            raise ValueError("unexpected GBM path shape")
        levels_payment = path[1:].astype(float)
        try:
            res = price_rila_single_premium(
                contract=contract,
                yield_curve=yield_curve,
                mortality=mortality,
                horizon_age=horizon_age,
                spread=spread,
                valuation_year=valuation_year,
                expenses=expenses,
                expenses_csv_path=expenses_csv_path,
                index_s0=float(path[0]),
                index_levels_payment=levels_payment,
                expense_annual_inflation=expense_annual_inflation,
            )
        except RILAPricingInfeasibleError as exc:
            if infeasible_path_policy == "raise":
                raise RILAPricingInfeasibleError(
                    k_loading=exc.k_loading,
                    premium_expense_rate=exc.premium_expense_rate,
                    detail=f"(Monte Carlo path index {i} of {n_sims})",
                ) from exc
            n_infeasible += 1
            loading = float(exc.k_loading) + float(exc.premium_expense_rate)
            if loading > worst_loading:
                worst_loading = loading
                worst_exc = exc
            continue
        prem[i] = float(res.single_premium)
        pvb[i] = float(res.pv_benefit)
        n_feasible += 1

    if n_feasible == 0:
        if worst_exc is None:
            raise RuntimeError("RILA Monte Carlo: every path failed but no error captured.")
        raise RILAPricingInfeasibleError(
            k_loading=worst_exc.k_loading,
            premium_expense_rate=worst_exc.premium_expense_rate,
            detail=f"(All {n_sims} Monte Carlo paths were infeasible.)",
        )

    pve = np.full(n_sims, pv_monthly_expenses_single, dtype=float)
    pve[np.isnan(prem)] = np.nan
    pvt = pvb + pve
    af = np.full(n_sims, annuity_factor, dtype=float)
    af[np.isnan(prem)] = np.nan

    return RILAMonteCarloResult(
        n_sims=int(n_sims),
        single_premium=prem,
        pv_benefit=pvb,
        pv_monthly_expenses=pve,
        pv_monthly_total=pvt,
        annuity_factor=af,
        premium_mean=float(np.nanmean(prem)),
        premium_median=float(np.nanmedian(prem)),
        premium_p05=float(np.nanpercentile(prem, 5.0)),
        premium_p95=float(np.nanpercentile(prem, 95.0)),
        pv_benefit_mean=float(np.nanmean(pvb)),
        pv_total_mean=float(np.nanmean(pvt)),
        n_feasible=int(n_feasible),
        n_infeasible=int(n_infeasible),
        infeasible_max_loading=float(worst_loading),
    )

"""Command-line entry points (portfolio runner v1)."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import pricing_projection as sp
from build_portfolio_excel_workbook import build_portfolio_workbook_to_path
from inforce_io import load_policy_inputs_from_csv
from portfolio import Portfolio
from portfolio_config import portfolio_v1_enabled
from portfolio_runner import run_portfolio
from portfolio_summary import portfolio_result_to_summary_dict
from pricing_scenario_materialize import run_scenario_for_portfolio_policies


def _cmd_portfolio_run(args: argparse.Namespace) -> int:
    if not portfolio_v1_enabled():
        print(
            "portfolio-run is disabled: remove annuity_model/.disable-portfolio-v1 if that "
            "opt-out file exists, or set ANNUITY_MODEL_PORTFOLIO_V1 to 1/true; unsetting the "
            "variable also enables portfolio locally (see portfolio_config.portfolio_v1_enabled).",
            file=sys.stderr,
        )
        return 2
    policies = load_policy_inputs_from_csv(args.inforce)
    pol_t = tuple(policies)
    sex_raw = str(getattr(pol_t[0].contract, "sex", "male")).strip().lower()
    sex = "female" if sex_raw == "female" else "male"
    scen = run_scenario_for_portfolio_policies({}, pol_t, sex=sex)  # type: ignore[arg-type]
    alm_asm = sp.alm_engine_baseline_assumptions() if bool(getattr(args, "alm", False)) else None
    try:
        res = run_portfolio(
            portfolio=Portfolio(policies=tuple(policies)),
            scenario=scen,
            alm_assumptions=alm_asm,
            max_workers=max(1, int(args.workers)),
        )
    except ValueError as exc:
        if alm_asm is not None and (
            "initial_asset_market_value" in str(exc) or "single_premium" in str(exc)
        ):
            print(
                "Warning: baseline ALM skipped (aggregate premium not positive). "
                "Writing pricing-only summary.",
                file=sys.stderr,
            )
            res = run_portfolio(
                portfolio=Portfolio(policies=tuple(policies)),
                scenario=scen,
                alm_assumptions=None,
                max_workers=max(1, int(args.workers)),
            )
        else:
            raise
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    summary_path = out_dir / "portfolio_summary.json"
    summary_path.write_text(
        json.dumps(portfolio_result_to_summary_dict(res), indent=2) + "\n",
        encoding="utf-8",
    )
    build_portfolio_workbook_to_path(res, out_dir / "portfolio.xlsx")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="annuity_model.cli")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_run = sub.add_parser("portfolio-run", help="Run portfolio pricing from an inforce CSV.")
    p_run.add_argument("--inforce", type=Path, required=True, help="Path to inforce CSV.")
    p_run.add_argument("--out", type=Path, required=True, help="Output directory.")
    p_run.add_argument(
        "--workers",
        type=int,
        default=1,
        help="Process pool size for per-policy pricing (default 1 = serial).",
    )
    p_run.add_argument(
        "--alm",
        action="store_true",
        help="Run deterministic baseline ALM on the aggregated liability path (see alm_engine_baseline_assumptions).",
    )
    p_run.set_defaults(func=_cmd_portfolio_run)

    args = parser.parse_args(argv)
    return int(args.func(args))


if __name__ == "__main__":
    raise SystemExit(main())

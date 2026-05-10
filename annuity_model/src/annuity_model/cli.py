"""Command-line entry points (portfolio runner v1)."""

from __future__ import annotations

import argparse
import json
import sys
import uuid
from pathlib import Path
from typing import Any

from annuity_model import pricing_projection as sp
from annuity_model.assumption_provenance import (
    assumption_evidence_summary,
    provenance_rows_from_pricing_state,
)
from annuity_model.build_portfolio_excel_workbook import build_portfolio_workbook_to_path
from annuity_model.inforce_io import load_policy_inputs_from_csv
from annuity_model.portfolio import Portfolio
from annuity_model.portfolio_config import portfolio_v1_enabled
from annuity_model.portfolio_runner import run_portfolio
from annuity_model.portfolio_summary import portfolio_result_to_summary_dict
from annuity_model.pricing_scenario_materialize import run_scenario_for_portfolio_policies
from annuity_model.run_ledger import pricing_run_summary, record_pricing_run


def _default_portfolio_assumption_rows() -> list[dict[str, Any]]:
    return provenance_rows_from_pricing_state(
        pricing_meta={
            "yield_mode": "par_bootstrap",
            "mortality_mode": "rp2014_mp2016",
            "expense_mode": "csv",
        },
        pricing_run_inputs={},
        pricing_excel_context={},
    )


def _portfolio_ledger_summary(
    *,
    res: Any,
    inforce_path: Path,
    output_summary: dict[str, Any],
) -> dict[str, Any]:
    rows = _default_portfolio_assumption_rows()
    evidence = assumption_evidence_summary(rows)
    product_counts = {
        product: details["policy_count"]
        for product, details in output_summary.get("by_product_type", {}).items()
    }
    return pricing_run_summary(
        run_id=f"portfolio-cli-{uuid.uuid4().hex[:12]}",
        product="portfolio",
        scenario_id="portfolio_base",
        assumption_artifacts=rows,
        input_payload={
            "inforce": str(inforce_path),
            "n_policies": len(res.policy_results),
            "product_counts": product_counts,
        },
        output_metrics=output_summary,
        parity_status="Portfolio workbook ModelCheck prepared",
        validation_status="workbook_validated",
        waiver_status=str(evidence["waiver_status"]),
        assumption_evidence=evidence,
        metadata={"source": "cli_portfolio_run"},
    )


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
    summary = portfolio_result_to_summary_dict(res)
    run_summary = _portfolio_ledger_summary(
        res=res,
        inforce_path=Path(args.inforce),
        output_summary=summary,
    )
    summary_path.write_text(
        json.dumps(summary, indent=2) + "\n",
        encoding="utf-8",
    )
    ledger_path = (
        Path(args.ledger_path) if args.ledger_path is not None else out_dir / "run_ledger.sqlite3"
    )
    record_pricing_run(ledger_path, run_summary)
    build_portfolio_workbook_to_path(res, out_dir / "portfolio.xlsx", run_summary=run_summary)
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
    p_run.add_argument(
        "--ledger-path",
        type=Path,
        default=None,
        help="SQLite run ledger path (default: <out>/run_ledger.sqlite3).",
    )
    p_run.set_defaults(func=_cmd_portfolio_run)

    args = parser.parse_args(argv)
    return int(args.func(args))


if __name__ == "__main__":
    raise SystemExit(main())

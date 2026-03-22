"""
Generate a synthetic monthly S&P 500 *proxy* index path for SPIA scenario CSV.

This is illustrative only (geometric random walk), not an official index print.
Usage:
    python generate_sp500_scenario_csv.py --months 540 --out sp500_scenario_projection_monthly.csv
"""

from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np


def write_scenario_csv(out_path: Path, n_months: int, *, seed: int = 42, s0: float = 5000.0) -> None:
    rng = np.random.default_rng(seed)
    levels = np.zeros(n_months + 1, dtype=float)
    levels[0] = s0
    # ~8% annual drift, ~16% annual vol as rough equity stylized numbers (monthly steps)
    mu_m = 0.08 / 12.0
    sig_m = 0.16 / (12.0**0.5)
    for k in range(1, n_months + 1):
        levels[k] = levels[k - 1] * float(np.exp(mu_m - 0.5 * sig_m**2 + sig_m * rng.standard_normal()))

    lines = ["month,sp500_level"]
    for m, v in enumerate(levels.tolist()):
        lines.append(f"{m},{v:.6f}")
    out_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--months", type=int, default=540, help="Number of payment months (rows 0..months)")
    p.add_argument("--out", type=Path, default=Path("sp500_scenario_projection_monthly.csv"))
    p.add_argument("--seed", type=int, default=42)
    p.add_argument("--s0", type=float, default=5000.0)
    args = p.parse_args()
    if args.months < 1:
        raise SystemExit("months must be >= 1")
    write_scenario_csv(args.out, args.months, seed=args.seed, s0=args.s0)
    print(f"Wrote {args.out.resolve()} with months 0..{args.months}")


if __name__ == "__main__":
    main()

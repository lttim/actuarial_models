"""Versioned data-artifact registry.

A single source of truth for every CSV / table that the engines and
builders read at runtime. Each artifact is described by:

* **kind** -- the data category (yield curve, mortality table, expense
  assumptions, index scenario, ...). Drives the on-disk folder.
* **version** -- the as-of date or table version
  (e.g. ``2026-03-20`` for a treasury snapshot, ``rp2014`` for a static
  table). Drives the on-disk subfolder. New snapshots get a new version
  folder; we **never** overwrite an existing artifact in-place -- that
  makes parity reruns reproducible and audit trails honest.
* **path** -- the resolved absolute :class:`pathlib.Path`, computed from
  the package directory so callers don't depend on the process CWD.
* **sha256** -- the expected hex digest. The invariant test
  ``test_data_registry_invariants.py`` recomputes this on every CI run
  and fails if a byte changes. This is what catches the "someone edited
  the rate curve in-place to fix a bug" class of silent parity drift.
* **source** -- a short human-readable string describing where the
  artifact came from. Required so an auditor doesn't need to dig
  through git blame to learn what ``sp500_seed_baseline`` means.

The legacy ``DEFAULT_*_CSV`` string constants in
:mod:`pricing_projection` resolve to these registry paths so existing
call sites (``pd.read_csv(DEFAULT_*_CSV)``) keep working without
filesystem layout knowledge.

Adding a new artifact is a 3-step change:

1. Drop the file under ``data/<kind>/<version>/<basename>``.
2. Append a :class:`DataArtifact` entry to :data:`REGISTRY`.
3. Run ``pytest tests/test_data_registry_invariants.py`` -- it will
   tell you the sha256 to paste into the new entry.
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass
from pathlib import Path

PACKAGE_DIR: Path = Path(__file__).resolve().parent
DATA_ROOT: Path = PACKAGE_DIR / "data"


@dataclass(frozen=True)
class DataArtifact:
    name: str
    kind: str
    version: str
    relative_path: str
    sha256: str
    source: str

    @property
    def path(self) -> Path:
        # Built relative to PACKAGE_DIR so tests / Streamlit / docker all
        # resolve the same file regardless of process CWD.
        return PACKAGE_DIR / self.relative_path

    def compute_sha256(self) -> str:
        # Streaming so large CSVs don't have to live in memory. None of
        # ours exceed a few MB today, but the SP500 monthly scenario
        # could grow if the projection horizon is extended.
        h = hashlib.sha256()
        with self.path.open("rb") as fh:
            for chunk in iter(lambda: fh.read(65536), b""):
                h.update(chunk)
        return h.hexdigest()


REGISTRY: tuple[DataArtifact, ...] = (
    DataArtifact(
        name="treasury_zero_curve",
        kind="yield_curves",
        version="2026-03-20",
        relative_path="data/yield_curves/2026-03-20/treasury_zero_rate_curve.csv",
        sha256="b4af978afefd57d93c4976de541c0395178929845909b57e615a303d75a4eb6b",
        source=(
            "Treasury par-yield snapshot bootstrapped to zero rates "
            "(continuous compounding). Initial seed snapshot captured "
            "2026-03-20 alongside the matching par-yield CSV; both move "
            "together when the snapshot is refreshed (new version folder)."
        ),
    ),
    DataArtifact(
        name="treasury_par_curve",
        kind="yield_curves",
        version="2026-03-20",
        relative_path="data/yield_curves/2026-03-20/treasury_par_yield_curve.csv",
        sha256="15b7c83bf8a7dc05382028a7c5d05139737dd5ce77980c13cbe7f63e4c5f21a9",
        source=(
            "Treasury par-yield curve, daily snapshot. Source for the "
            "matching zero curve in this folder via the bootstrapper "
            "in pricing_projection.bootstrap_zero_rates_from_par_yields."
        ),
    ),
    DataArtifact(
        name="rp2014_male_healthy_annuitant_qx",
        kind="mortality",
        version="rp2014",
        relative_path="data/mortality/rp2014/rp2014_male_healthy_annuitant_qx_2014.csv",
        sha256="8de65dfc0e819977f61b0e85b88f3167654ee6cc2d12daea5ccff1a2ff816c9c",
        source=(
            "Society of Actuaries RP-2014 Mortality Table -- Male "
            "Healthy Annuitant base qx (2014). Cached from the SOA "
            "xlsx by ensure_rp2014_male_healthy_annuitant_qx_2014_csv()."
        ),
    ),
    DataArtifact(
        name="mp2016_male_improvement_rates",
        kind="mortality",
        version="mp2016",
        relative_path="data/mortality/mp2016/mp2016_male_improvement_rates.csv",
        sha256="a2524256d964e1e8e2cd128926af64c8473deb276eb79c9794ae987b5526a73c",
        source=(
            "Society of Actuaries Mortality Improvement Scale MP-2016 "
            "(Male). Cached from the SOA xlsx by "
            "ensure_mp2016_male_improvement_csv()."
        ),
    ),
    DataArtifact(
        name="expenses_assumptions_us_placeholders",
        kind="expenses",
        version="us_placeholders",
        relative_path="data/expenses/us_placeholders/expenses_assumptions_us_placeholders.csv",
        sha256="09a0398f0932f188e323f233084eb619f474405dbcd160081f54fce937242f75",
        source=(
            "Placeholder US expense assumptions (acquisition, "
            "maintenance, ULAE) used as defaults when the user has not "
            "uploaded a real expense file. NOT a market reference -- "
            "do not cite as actuarial assumption set."
        ),
    ),
    DataArtifact(
        name="sp500_scenario_monthly_seed_baseline",
        kind="index_scenarios",
        version="sp500_seed_baseline",
        relative_path="data/index_scenarios/sp500_seed_baseline/sp500_scenario_projection_monthly.csv",
        sha256="eaff15f61767d6802d33dee675541392cc86e32accbeca237fd456c526cecc44",
        source=(
            "Synthetic SP500 monthly index level scenario (single "
            "deterministic path) used by the RILA replication smoke and "
            "Excel ladder defaults. Generated by "
            "generate_sp500_scenario_csv.py -- regenerating produces a "
            "DIFFERENT scenario; lock the seed in the new version folder."
        ),
    ),
    DataArtifact(
        name="cso_2017_ult_male_nonsmoker_qx",
        kind="mortality",
        version="cso_2017_ult",
        relative_path="data/mortality/cso_2017_ult/cso_2017_ult_male_nonsmoker_qx.csv",
        sha256="de83ac9d9df43f477b1cd869e3e09f7d3a40c65ae73c155590d46e207c04bd07",
        source=(
            "SYNTHETIC PLACEHOLDER -- NOT licensed CSO 2017 Ultimate "
            "data. Gompertz-Makeham approximation generated by "
            "scripts/generate_cso_2017_synthetic.py. Production users "
            "MUST overlay their own licensed CSO file at the same path."
        ),
    ),
    DataArtifact(
        name="cso_2017_ult_female_nonsmoker_qx",
        kind="mortality",
        version="cso_2017_ult",
        relative_path="data/mortality/cso_2017_ult/cso_2017_ult_female_nonsmoker_qx.csv",
        sha256="01ab648629813f652d8af99879f887a03a3174caef4c1b53dcb229a2c48d5c7a",
        source=(
            "SYNTHETIC PLACEHOLDER -- NOT licensed CSO 2017 Ultimate "
            "data. Gompertz-Makeham approximation (female multiplier "
            "0.70 applied to male NS base). Overlay licensed file in "
            "production."
        ),
    ),
    DataArtifact(
        name="cso_2017_ult_male_smoker_qx",
        kind="mortality",
        version="cso_2017_ult",
        relative_path="data/mortality/cso_2017_ult/cso_2017_ult_male_smoker_qx.csv",
        sha256="0f4b83529f97fc7714aed1b6d641a3903ee07bc2110f67a6d0517166725470ad",
        source=(
            "SYNTHETIC PLACEHOLDER -- NOT licensed CSO 2017 Ultimate "
            "data. Gompertz-Makeham approximation (smoker multiplier "
            "2.20 applied to male NS base). Overlay licensed file in "
            "production."
        ),
    ),
    DataArtifact(
        name="cso_2017_ult_female_smoker_qx",
        kind="mortality",
        version="cso_2017_ult",
        relative_path="data/mortality/cso_2017_ult/cso_2017_ult_female_smoker_qx.csv",
        sha256="86c4220a04f584525d765dd68dbfa04a717ab47b5b8829f5f7a9ff90dfefdc0b",
        source=(
            "SYNTHETIC PLACEHOLDER -- NOT licensed CSO 2017 Ultimate "
            "data. Gompertz-Makeham approximation (smoker 2.20 + female "
            "0.70 multipliers applied to male NS base). Overlay licensed "
            "file in production."
        ),
    ),
)


_BY_NAME: dict[str, DataArtifact] = {a.name: a for a in REGISTRY}


def get_artifact(name: str) -> DataArtifact:
    """Return the registry entry for *name* or raise KeyError with a hint."""
    try:
        return _BY_NAME[name]
    except KeyError as e:
        known = ", ".join(sorted(_BY_NAME))
        raise KeyError(
            f"Unknown data artifact {name!r}. Known artifacts: {known}. "
            f"Add a new DataArtifact entry to data_registry.REGISTRY."
        ) from e


def path_str(name: str) -> str:
    """Return the artifact's path as a string (handy for legacy callers)."""
    return str(get_artifact(name).path)

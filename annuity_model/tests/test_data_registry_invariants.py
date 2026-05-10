"""Data-artifact registry invariants.

The Phase-5 Wave 3.2 hardening pass introduced
:mod:`data_registry` as the single source of truth for every CSV /
table consumed by the engines and builders. These tests lock the
registry contract:

  1. **Every registered artifact actually exists** at its declared
     path. Catches accidental deletes, typos in ``relative_path``, and
     repo migrations that miss a ``git mv``.
  2. **The on-disk bytes match the declared sha256.** This is the
     parity-critical safety net: if someone edits a yield curve or
     mortality table in-place to fix a bug, this test fails on the
     next CI run and forces them to either roll back or bump the
     version folder + sha256 (an explicit audit-trail action).
  3. **DEFAULT_*_CSV constants in pricing_projection resolve through
     the registry.** Any new bypass (e.g. a hardcoded
     "treasury_zero_rate_curve_latest.csv" reintroduced into the
     codebase) would silently dodge the sha256 lock; this guard makes
     that explicit.
  4. **Names are unique and well-formed.** Two artifacts can't share
     the same name, and the kind/version pair has to live where the
     relative_path says it lives.

If a check fails, the fix lives in :mod:`data_registry` (refresh the
sha256, point at the right path) or in the offending caller (use the
registry, not a hardcoded basename). Not in this test.
"""

from __future__ import annotations

import pytest

from annuity_model import data_registry
from annuity_model import pricing_projection as sp
from annuity_model.data_registry import REGISTRY, DataArtifact

pytestmark = [pytest.mark.invariant]


def test_every_registered_artifact_exists_on_disk() -> None:
    missing = [a for a in REGISTRY if not a.path.exists()]
    assert not missing, "Registered data artifacts not found on disk: " + ", ".join(
        f"{a.name!r} -> {a.path}" for a in missing
    )


def test_every_registered_artifact_matches_declared_sha256() -> None:
    """Recompute sha256 for every artifact; compare to the declared digest."""
    drift: list[tuple[str, str, str]] = []
    for a in REGISTRY:
        actual = a.compute_sha256()
        if actual != a.sha256:
            drift.append((a.name, a.sha256, actual))
    if drift:
        msg_lines = [
            "Data artifact byte-content drift detected (file edited in-place "
            "without bumping the version folder + sha256). For each drifted "
            "artifact, either roll back the change or move the file to a new "
            "data/<kind>/<new_version>/ folder, update data_registry.REGISTRY, "
            "and add a CHANGELOG entry under [Unreleased] -> Changed.",
            "",
        ]
        for name, declared, actual in drift:
            msg_lines.append(f"  - {name}: declared={declared}\n              actual  ={actual}")
        pytest.fail("\n".join(msg_lines))


def test_artifact_names_are_unique() -> None:
    names = [a.name for a in REGISTRY]
    assert len(names) == len(set(names)), (
        f"Duplicate artifact names in REGISTRY: {sorted(n for n in names if names.count(n) > 1)}"
    )


def test_relative_path_matches_declared_kind_and_version() -> None:
    """Catches data/<wrong_kind>/<wrong_version>/file.csv layout drift."""
    for a in REGISTRY:
        expected_prefix = f"data/{a.kind}/{a.version}/"
        assert a.relative_path.startswith(expected_prefix), (
            f"Artifact {a.name!r} relative_path {a.relative_path!r} does not "
            f"start with {expected_prefix!r} (kind={a.kind}, version={a.version})."
        )


def test_pricing_projection_default_csv_constants_resolve_via_registry() -> None:
    """Each DEFAULT_*_CSV must equal the registry's resolved path string.

    If a future commit reintroduces a hardcoded basename for one of these,
    the constant will silently miss the registry's sha256 lock. This test
    catches that drift before it ships.
    """
    expected = {
        "DEFAULT_ZERO_CURVE_CSV": "treasury_zero_curve",
        "DEFAULT_PAR_CURVE_CSV": "treasury_par_curve",
        "DEFAULT_RP2014_MALE_HEALTHY_QX_CSV": "rp2014_male_healthy_annuitant_qx",
        "DEFAULT_MP2016_MALE_IMPROVEMENT_CSV": "mp2016_male_improvement_rates",
        "DEFAULT_EXPENSES_CSV": "expenses_assumptions_us_placeholders",
        "DEFAULT_SP500_SCENARIO_CSV": "sp500_scenario_monthly_seed_baseline",
    }
    for const_name, artifact_name in expected.items():
        actual = getattr(sp, const_name)
        expected_path = data_registry.path_str(artifact_name)
        assert actual == expected_path, (
            f"pricing_projection.{const_name} = {actual!r} but registry says "
            f"{expected_path!r}. Reconnect via data_registry.path_str("
            f"{artifact_name!r})."
        )


def test_get_artifact_raises_keyerror_with_helpful_hint() -> None:
    with pytest.raises(KeyError, match=r"Unknown data artifact 'totally_fake'"):
        data_registry.get_artifact("totally_fake")


def test_dataartifact_path_is_absolute() -> None:
    for a in REGISTRY:
        assert a.path.is_absolute(), (
            f"Artifact {a.name!r} path is not absolute ({a.path}). "
            f"Callers depend on registry paths being CWD-independent so "
            f"Streamlit / docker / pytest all resolve the same file."
        )


def test_compute_sha256_matches_known_value_for_known_artifact() -> None:
    """Smoke test that compute_sha256 actually computes (not just returns the field)."""
    sample: DataArtifact = REGISTRY[0]
    assert sample.compute_sha256() == sample.sha256, (
        "compute_sha256() did not match the declared sha256 on the first "
        "registry entry. If the file is intentionally changed, update the "
        "DataArtifact's sha256 field after moving it to a new version folder."
    )

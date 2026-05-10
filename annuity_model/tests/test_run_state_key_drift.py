"""One-way ratchet against new raw ``"run_*"`` Streamlit session-state literals.

Background
----------
``pricing_run_form_state.py`` defines the canonical Pricing Run session
keys as class attributes on :class:`pricing_run_form_state.RUN_KEY` and
exposes the string set as :data:`RUN_STATE_KEY_NAMES`. The intent is
that all NEW code references these keys via the symbol
(``RUN_KEY.ISSUE_AGE``) rather than the corresponding literal,
so a typo becomes an ``AttributeError`` and an IDE rename works.

Today's reality is that ``pricing_ui.py`` predates this refactor and
contains ~100 raw literals. Migrating all of them at once would be a
large diff with high regression risk; the ``ui/MIGRATION.md`` plan
splits ``pricing_ui.py`` into per-page modules under ``ui/pages/`` and
naturally retires the literals as it goes.

This test enforces the ratchet during the interim:

  * ``tests/run_state_key_baseline.json`` is the per-file baseline.
  * The actual count for any file MUST NOT exceed its baseline.
  * The actual count MAY drop below its baseline as code migrates.
  * A file NOT in the baseline MUST contain zero canonical literals
    (= no new file is allowed to introduce raw literals).

Updating the baseline
---------------------
After migrating literals (count went down), refresh the baseline::

    UPDATE_RUN_STATE_BASELINE=1 python -m pytest tests/test_run_state_key_drift.py

This rewrites ``tests/run_state_key_baseline.json`` to today's values.
The script will refuse to RAISE any baseline (those edits must be
manual, with reviewer sign-off, justified by an explicit "we need to
keep one more literal in this file because X" rationale). CODEOWNERS
covers the baseline file under the same blanket rule that covers
other process-discipline gates.
"""

from __future__ import annotations

import json
import os
import re
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src"
BASELINE_PATH = REPO_ROOT / "tests" / "run_state_key_baseline.json"
SOURCE_OF_TRUTH = "src/annuity_model/pricing_run_form_state.py"


def _canonical_keys() -> frozenset[str]:
    """Late-binding import so the test fails loudly if the module breaks."""
    import sys

    sys.path.insert(0, str(SRC_ROOT))
    from annuity_model.pricing_run_form_state import RUN_STATE_KEY_NAMES

    return RUN_STATE_KEY_NAMES


def _load_baseline() -> dict[str, int]:
    raw = json.loads(BASELINE_PATH.read_text())
    return {k: v for k, v in raw.items() if not k.startswith("_") and isinstance(v, int)}


def _walk_python_files() -> list[Path]:
    out: list[Path] = []
    for p in sorted(REPO_ROOT.rglob("*.py")):
        rel = str(p.relative_to(REPO_ROOT))
        if "/.venv/" in str(p) or "/__pycache__/" in str(p) or rel.startswith(".venv"):
            continue
        if rel.startswith("build/"):
            continue
        if rel == SOURCE_OF_TRUTH:
            continue
        out.append(p)
    return out


def _count_canonical_literals(path: Path, keys: frozenset[str]) -> int:
    text = path.read_text()
    if "run_" not in text:
        return 0
    pattern = re.compile("|".join(rf'["\']{re.escape(k)}["\']' for k in keys))
    return len(pattern.findall(text))


def _measure_actual() -> dict[str, int]:
    keys = _canonical_keys()
    counts: dict[str, int] = {}
    for p in _walk_python_files():
        n = _count_canonical_literals(p, keys)
        if n > 0:
            counts[str(p.relative_to(REPO_ROOT))] = n
    return counts


def _maybe_update_baseline(actual: dict[str, int]) -> bool:
    """Rewrite the baseline if ``UPDATE_RUN_STATE_BASELINE=1`` is set.

    Refuses to raise any individual file's baseline -- the env-var path
    is for migration credit only. Raising a baseline requires editing
    the JSON file by hand.
    """
    if os.environ.get("UPDATE_RUN_STATE_BASELINE") != "1":
        return False
    baseline = _load_baseline()
    for path, count in actual.items():
        prior = baseline.get(path, 0)
        if count > prior:
            raise AssertionError(
                f"refusing to update: actual {count} > baseline {prior} "
                f"for {path}. The env-var path only ratchets DOWN. "
                "If you genuinely need a higher cap, edit "
                f"{BASELINE_PATH.relative_to(REPO_ROOT)} by hand."
            )
    raw = json.loads(BASELINE_PATH.read_text())
    new_raw = {k: v for k, v in raw.items() if k.startswith("_")}
    new_raw.update(dict(sorted(actual.items())))
    BASELINE_PATH.write_text(json.dumps(new_raw, indent=2) + "\n")
    return True


@pytest.fixture(scope="module")
def actual_counts() -> dict[str, int]:
    return _measure_actual()


@pytest.fixture(scope="module")
def baseline() -> dict[str, int]:
    return _load_baseline()


def test_no_file_exceeds_baseline(actual_counts: dict[str, int], baseline: dict[str, int]) -> None:
    """A regression here means new ``"run_*"`` literals snuck in.

    Fix by replacing them with ``RUN_KEY.<NAME>`` (importing
    :class:`pricing_run_form_state.RUN_KEY`). Do NOT just bump the
    baseline -- the whole point of the ratchet is that the literal
    count goes down over time, never up.
    """
    if _maybe_update_baseline(actual_counts):
        pytest.skip("baseline updated via UPDATE_RUN_STATE_BASELINE=1")
    overshoots: list[str] = []
    for path, count in actual_counts.items():
        prior = baseline.get(path)
        if prior is None:
            overshoots.append(
                f"  {path}: {count} canonical 'run_*' literals (NEW file -- "
                "must use RUN_KEY.<NAME>; NO new files may introduce raw "
                "literals)"
            )
        elif count > prior:
            overshoots.append(
                f"  {path}: {count} > {prior} (baseline). {count - prior} "
                "new literal(s); replace with RUN_KEY.<NAME>."
            )
    assert not overshoots, (
        "Raw 'run_*' session-state literals increased. The ratchet only "
        "permits the count to DECREASE. Use RUN_KEY.<NAME> from "
        "pricing_run_form_state for any new reference.\n" + "\n".join(overshoots)
    )


def test_baseline_files_still_exist(baseline: dict[str, int]) -> None:
    """If a file is in the baseline but no longer exists, drop it."""
    missing = [path for path in baseline if not (REPO_ROOT / path).is_file()]
    assert not missing, (
        "tests/run_state_key_baseline.json references non-existent files: "
        f"{missing}. Remove them from the baseline."
    )


def test_baseline_counts_are_nonnegative(baseline: dict[str, int]) -> None:
    bad = {k: v for k, v in baseline.items() if not isinstance(v, int) or v < 0}
    assert not bad, f"baseline values must be non-negative ints; got: {bad}"


def test_source_of_truth_excluded() -> None:
    """``pricing_run_form_state.py`` is intentionally NOT counted -- the
    literals there ARE the constants. If a future contributor adds it
    to the baseline by accident, fail loudly."""
    raw = json.loads(BASELINE_PATH.read_text())
    assert SOURCE_OF_TRUTH not in raw, (
        f"{SOURCE_OF_TRUTH} must NOT be in the baseline; it owns the canonical RUN_KEY constants."
    )


def test_run_key_namespace_round_trips() -> None:
    """Every name in :data:`RUN_STATE_KEY_NAMES` must be reachable as a
    class attribute on :class:`RUN_KEY` (i.e. the reflection that drives
    the namespace agrees with the public set).
    """
    import sys

    sys.path.insert(0, str(SRC_ROOT))
    from annuity_model.pricing_run_form_state import RUN_KEY, RUN_STATE_KEY_NAMES

    attr_values = {v for v in vars(RUN_KEY).values() if isinstance(v, str)}
    assert attr_values >= RUN_STATE_KEY_NAMES, (
        "RUN_STATE_KEY_NAMES contains entries not declared on RUN_KEY: "
        f"{sorted(RUN_STATE_KEY_NAMES - attr_values)}"
    )
    extras = {v for v in attr_values if v.startswith("run_")} - RUN_STATE_KEY_NAMES
    assert not extras, (
        f"RUN_KEY declares 'run_*' attributes missing from RUN_STATE_KEY_NAMES: {sorted(extras)}"
    )

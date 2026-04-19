"""Drift gate between the active and deferred branch-protection profiles.

The repo ships two branch-protection JSONs:

* ``.github/branch-protection.json`` -- ACTIVE today
  (``required_pull_request_reviews = null`` because there is exactly one
  CODEOWNER and GitHub blocks self-approval).
* ``.github/branch-protection.with-second-reviewer.json`` -- the
  drop-in replacement to apply the day a second CODEOWNER (or
  ``@lttim/actuarial-reviewers``) lands.

The two MUST stay byte-identical everywhere except the
``required_pull_request_reviews`` block. Otherwise the deferred profile
silently degrades over time and, on activation day, weakens an
unrelated gate (status-checks list, linear history, force-push policy,
or conversation resolution) that the maintainer never reviewed.

This test fires whenever either file changes, and pins exactly which
keys are allowed to differ.
"""

from __future__ import annotations

import json
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
ACTIVE = REPO_ROOT / ".github" / "branch-protection.json"
DEFERRED = REPO_ROOT / ".github" / "branch-protection.with-second-reviewer.json"

# Keys that ARE allowed to differ between the two profiles. Today this
# is exactly one key -- the whole point of the deferred profile is to
# turn on PR reviews. Any other intentional divergence requires
# updating this allow-list AND a CHANGELOG entry explaining why.
_ALLOWED_DIFF_KEYS: frozenset[str] = frozenset({"required_pull_request_reviews"})


def _load(path: Path) -> dict[str, object]:
    assert path.is_file(), f"missing {path}"
    return json.loads(path.read_text())


def test_both_profiles_exist() -> None:
    assert ACTIVE.is_file()
    assert DEFERRED.is_file()


def test_active_profile_has_no_pr_reviews_required() -> None:
    """Sanity check: the ACTIVE profile MUST stay null on
    required_pull_request_reviews until the second-CODEOWNER upgrade
    runs, because requiring reviews on a one-owner repo deadlocks every
    PR (GitHub forbids self-approval).
    """
    payload = _load(ACTIVE)
    assert payload["required_pull_request_reviews"] is None, (
        "Active branch-protection profile started requiring PR reviews "
        "before the second CODEOWNER landed. This will deadlock every "
        "PR -- GitHub blocks self-approval on protected branches. See "
        "annuity_model/docs/CODEOWNERS_RATIONALE.md, section "
        "'Second-CODEOWNER upgrade path'."
    )


def test_deferred_profile_actually_requires_reviews() -> None:
    """The deferred profile must require >=1 code-owner review or it
    isn't a real upgrade -- it would be byte-identical to the active
    profile and adding it would be pointless."""
    payload = _load(DEFERRED)
    pr = payload["required_pull_request_reviews"]
    assert isinstance(pr, dict), (
        "Deferred profile must specify required_pull_request_reviews; "
        "got null. The whole purpose of this file is to turn that gate "
        "on once a second reviewer exists."
    )
    assert pr.get("required_approving_review_count", 0) >= 1
    assert pr.get("require_code_owner_reviews") is True


def test_profiles_match_outside_pr_reviews() -> None:
    """Every key OTHER than required_pull_request_reviews must be
    byte-identical. Drift here means the deferred profile would, on
    activation, silently weaken a gate (status checks, linear history,
    force-push policy, conversation resolution, etc.) that the
    maintainer never explicitly reviewed.
    """
    a = _load(ACTIVE)
    b = _load(DEFERRED)

    keys = (set(a) | set(b)) - {"_comment"} - _ALLOWED_DIFF_KEYS
    diff: list[str] = []
    for k in sorted(keys):
        if a.get(k) != b.get(k):
            diff.append(f"  {k}: active={a.get(k)!r} deferred={b.get(k)!r}")
    assert not diff, (
        "Branch-protection profiles drifted outside the allow-listed key "
        f"({sorted(_ALLOWED_DIFF_KEYS)}). Update both files in the same PR "
        "or extend _ALLOWED_DIFF_KEYS with a CHANGELOG entry explaining "
        "why the divergence is intentional. Diff:\n" + "\n".join(diff)
    )


def test_status_check_contexts_lockstep() -> None:
    """Status-check contexts list is the most common source of drift
    (every new workflow job lands as a contexts entry). Pin it
    separately for a clearer error message than the generic
    above test would produce.
    """
    a_ctx = _load(ACTIVE)["required_status_checks"]
    b_ctx = _load(DEFERRED)["required_status_checks"]
    assert a_ctx == b_ctx, (
        "required_status_checks drifted between active and deferred "
        "branch-protection profiles. Sync them in the same PR. "
        f"active={a_ctx!r} deferred={b_ctx!r}"
    )

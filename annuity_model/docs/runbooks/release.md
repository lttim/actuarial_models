# Runbook: cut a release

A release in this codebase = a tagged commit on `main` plus a `CHANGELOG.md`
entry plus a `model_change_log.md` entry if any parity-impacting code or
constant changed.

## Pre-flight (15 minutes)

1. Ensure `main` is green:
   ```bash
   gh run list --branch main --limit 5
   ```
   All five must show `success`. If any is failing, fix it first.

2. Local full-suite green:
   ```bash
   cd annuity_model
   python -m pytest -q
   python scripts/deep_smoke.py
   python scripts/render_parity_contract.py --check
   ```

3. Confirm tolerance constants did not silently change since last tag:
   ```bash
   git diff $(git describe --tags --abbrev=0)..HEAD -- annuity_model/parity_constants.py
   ```
   If non-empty, ensure each change is accompanied by a `model_change_log.md`
   entry.

## Bump the version

The version lives in `annuity_model/pyproject.toml` (`[project] version = ...`).
Use semantic versioning:

* **Patch** (`x.y.Z`): bug fix, no API change, no parity tolerance change.
* **Minor** (`x.Y.0`): new product, new public API, parity tightening.
* **Major** (`X.0.0`): breaking API change, parity loosening, schema break.

```bash
# Edit pyproject.toml
git add annuity_model/pyproject.toml
```

## Update the changelogs

1. **Engineering changelog** (`annuity_model/docs/CHANGELOG.md`):
   * Move everything under `## [Unreleased]` into a new
     `## [<version>] - YYYY-MM-DD` heading.
   * Categories: `Added`, `Changed`, `Fixed`, `Removed`, `Security`.

2. **Model change log** (`annuity_model/docs/model_change_log.md`):
   * Required iff any of these changed: parity tolerance, mortality table,
     yield-curve construction, ALM disinvest/reinvest rule, RILA crediting.
   * Required content per entry: PR link, reviewer, parity-trace before/after
     attached, justification.

## Cut the tag

```bash
git checkout main
git pull --ff-only
git tag -a v<version> -m "Release v<version>"
git push origin v<version>
```

The CI workflow `release.yml` (will be added in P4) builds the wheel,
attaches it to the GitHub release, and publishes the mkdocs site.

## Post-release sanity

1. Bootstrap from scratch on a fresh checkout to confirm the new version
   installs cleanly:
   ```bash
   cd /tmp && git clone <repo> Code_Sandbox_release_check
   cd Code_Sandbox_release_check && bash annuity_model/bootstrap_macos.sh
   ```

2. Smoke the released build:
   ```bash
   python annuity_model/scripts/deep_smoke.py
   ```

3. Reset `## [Unreleased]` in `CHANGELOG.md` for the next development
   cycle and push as `chore: open <next-version> dev cycle`.

## Rollback

If a release goes bad in the field:

1. Revert the tag locally and push the deletion:
   ```bash
   git tag -d v<version>
   git push --delete origin v<version>
   ```
2. Force the docs site to redeploy the previous tag.
3. File a `regression: <summary>` issue and require a `model_change_log.md`
   entry on the rollback PR.

## Branch protection refresh

The required CI status checks are declared in
[`.github/branch-protection.json`](../../../.github/branch-protection.json).
Whenever you add or rename a CI job, re-apply protection so GitHub knows it
must wait on the new context (otherwise PRs will hang forever waiting on a
status that never reports, or worse, merge without it):

```bash
gh api -X PUT repos/:owner/:repo/branches/main/protection \
  --input .github/branch-protection.json
```

Currently required (verify with `gh api repos/:owner/:repo/branches/main/protection`):

* `tests (<os> / py3.11|py3.12)` matrix from `ci.yml`
* `pre-commit (lint + format + mypy)` from `ci.yml`
* `docker build + deep_smoke in container` from `ci.yml`
* `parity + validator (ubuntu / py3.12)` from `parity-gate.yml`
  (always-on PR gate as of P0 hardening 2026-04 — no path filter)
* `build-and-deploy` from `docs.yml`

If a required context goes stale (e.g. job rename), apply the JSON immediately
in the same PR that renames the job.

## Related

* [debug_validator_failure.md](debug_validator_failure.md)
* [investigate_parity_break.md](investigate_parity_break.md)
* `annuity_model/docs/CODEOWNERS_RATIONALE.md`

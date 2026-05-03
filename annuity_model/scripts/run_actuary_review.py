"""run_actuary_review -- collect the evidence pack for the Actuary SME loop.

The Actuary SME (skill at ``annuity_model/.cursor/skills/actuary-sme/SKILL.md``)
runs as a readonly subagent. Before the subagent is invoked, this script
gathers the deterministic evidence the SME needs into a single overwriting
file::

    .cursor/actuary-reviews/_evidence-current.md

The evidence pack is intentionally small (target: <5s wall-clock) because
this script runs at the start of EVERY iteration of the SME's autonomous
fix-and-rereview loop. It does NOT re-run pytest; it READS the cached
last-run results from ``.pytest_cache``. The mid-loop ``pytest tests/parity``
run (managed by the orchestration rule) keeps that cache fresh.

What it collects
----------------

1. Repo state header (UTC timestamp, branch, HEAD commit, scope, iteration,
   prior verdict path).
2. Capped git diff (per-file content truncated; ``--stat`` summary always
   in full).
3. Recent ``docs/model_change_log.md`` entries (last 2).
4. Cached parity test status (from ``.pytest_cache/v/cache/lastfailed``).
5. Full dump of ``actuarial_benchmarks.py`` constants.
6. Detected per-product scope from the changed file paths.

Usage
-----

::

    python scripts/run_actuary_review.py \\
        --scope incremental|full|product:<name>|product:portfolio \\
        --iteration <N> \\
        [--prior-verdict <path>]

Always overwrites ``.cursor/actuary-reviews/_evidence-current.md``.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import re
import subprocess
import sys
from pathlib import Path

_HERE = Path(__file__).resolve()
_ANNUITY_DIR = _HERE.parent.parent
_REPO_ROOT = _ANNUITY_DIR.parent
if str(_ANNUITY_DIR) not in sys.path:
    sys.path.insert(0, str(_ANNUITY_DIR))

import actuarial_benchmarks as ab  # noqa: E402

EVIDENCE_DIR = _REPO_ROOT / ".cursor" / "actuary-reviews"
EVIDENCE_PATH = EVIDENCE_DIR / "_evidence-current.md"
CHANGE_LOG_PATH = _ANNUITY_DIR / "docs" / "model_change_log.md"
PYTEST_CACHE_LASTFAILED = _ANNUITY_DIR / ".pytest_cache" / "v" / "cache" / "lastfailed"
PYTEST_CACHE_NODEIDS = _ANNUITY_DIR / ".pytest_cache" / "v" / "cache" / "nodeids"

# Per-file diff content cap. Always show the first N lines plus a
# truncation marker for any file longer than that. The git --stat
# summary is always rendered in full so the SME still sees overall scope.
PER_FILE_DIFF_CAP_LINES = 80

# Number of trailing model_change_log entries to include.
CHANGE_LOG_TAIL = 2

# Mapping from changed-path fragments to product codes used in the
# Actuary SME workflow.
_PRODUCT_PATH_PATTERNS: tuple[tuple[str, str], ...] = (
    ("pricing_projection.py", "spia"),
    ("term_projection.py", "term"),
    ("rila_projection.py", "rila"),
    ("myga_projection.py", "myga"),
    ("fia_projection.py", "fia"),
    ("va_projection.py", "va"),
    ("wl_projection.py", "wl"),
    ("ul_projection.py", "ul"),
    ("iul_projection.py", "iul"),
    ("vul_projection.py", "vul"),
    ("build_pricing_excel_workbook.py", "spia"),
    ("build_term_excel_workbook.py", "term"),
    ("build_rila_excel_workbook.py", "rila"),
    ("build_myga_excel_workbook.py", "myga"),
    ("build_fia_excel_workbook.py", "fia"),
    ("build_va_excel_workbook.py", "va"),
    ("build_wl_excel_workbook.py", "wl"),
    ("build_ul_excel_workbook.py", "ul"),
    ("build_iul_excel_workbook.py", "iul"),
    ("build_vul_excel_workbook.py", "vul"),
    ("products/whole_life/", "wl"),
    ("products/universal_life/", "ul"),
    ("products/indexed_ul/", "iul"),
    ("products/variable_ul/", "vul"),
    ("products/variable_annuity/", "va"),
    ("products/myga/", "myga"),
    ("products/fia/", "fia"),
    ("test_spia_actuarial.py", "spia"),
    ("test_term_actuarial.py", "term"),
    ("test_rila_actuarial.py", "rila"),
    ("test_myga_actuarial.py", "myga"),
    ("test_fia_actuarial.py", "fia"),
    ("test_va_actuarial.py", "va"),
    ("test_wl_actuarial.py", "wl"),
    ("test_ul_actuarial.py", "ul"),
    ("test_iul_actuarial.py", "iul"),
    ("test_vul_actuarial.py", "vul"),
    ("portfolio.py", "portfolio"),
    ("portfolio_runner.py", "portfolio"),
    ("liability_aggregation.py", "portfolio"),
    ("build_portfolio_excel_workbook.py", "portfolio"),
    ("inforce_io.py", "portfolio"),
    ("inforce_parsers.py", "portfolio"),
    ("portfolio_summary.py", "portfolio"),
    ("portfolio_scenario.py", "portfolio"),
    ("products/spia/inforce.py", "portfolio"),
    ("products/term/inforce.py", "portfolio"),
    ("products/rila/inforce.py", "portfolio"),
    ("products/myga/inforce.py", "portfolio"),
    ("products/fia/inforce.py", "portfolio"),
    ("products/variable_annuity/inforce.py", "portfolio"),
    ("products/whole_life/inforce.py", "portfolio"),
    ("products/universal_life/inforce.py", "portfolio"),
    ("products/indexed_ul/inforce.py", "portfolio"),
    ("products/variable_ul/inforce.py", "portfolio"),
)


def _git(*args: str) -> str:
    """Run a git subcommand from the repo root and return stdout."""
    try:
        result = subprocess.run(
            ["git", *args],
            cwd=_REPO_ROOT,
            capture_output=True,
            text=True,
            check=False,
        )
        return result.stdout
    except FileNotFoundError:
        return ""


def _utc_now() -> str:
    return _dt.datetime.now(_dt.UTC).strftime("%Y-%m-%d %H:%M:%S UTC")


def _branch() -> str:
    out = _git("branch", "--show-current").strip()
    return out or "(detached)"


def _head_short() -> str:
    return _git("rev-parse", "--short", "HEAD").strip() or "(unknown)"


def _diff_stat() -> str:
    return _git("diff", "--stat", "HEAD").strip()


def _untracked_files() -> list[str]:
    out = _git("ls-files", "--others", "--exclude-standard").strip()
    return [p for p in out.splitlines() if p]


def _changed_files() -> list[str]:
    """Return the list of changed-vs-HEAD file paths (modified + untracked)."""
    tracked = _git("diff", "--name-only", "HEAD").strip().splitlines()
    untracked = _untracked_files()
    return sorted({p for p in tracked + untracked if p})


def _per_file_diff(path: str) -> str:
    """Return a capped diff for one file (handles tracked + untracked)."""
    if path in _untracked_files():
        full = _git("diff", "--no-index", "/dev/null", path)
    else:
        full = _git("diff", "HEAD", "--", path)
    lines = full.splitlines()
    if len(lines) <= PER_FILE_DIFF_CAP_LINES:
        return full
    head = "\n".join(lines[:PER_FILE_DIFF_CAP_LINES])
    return f"{head}\n... ({len(lines) - PER_FILE_DIFF_CAP_LINES} more lines)\n"


def _detect_products(changed: list[str]) -> list[str]:
    """Map changed file paths to the set of product codes in scope."""
    detected: set[str] = set()
    for path in changed:
        for fragment, code in _PRODUCT_PATH_PATTERNS:
            if fragment in path:
                detected.add(code)
    return sorted(detected)


_CHANGE_LOG_HEADER_RE = re.compile(r"^## (\d{4}-\d{2}-\d{2}.*?)$", re.MULTILINE)


def _change_log_tail() -> str:
    """Return the last CHANGE_LOG_TAIL `## YYYY-MM-DD` entries (verbatim)."""
    if not CHANGE_LOG_PATH.exists():
        return "(model_change_log.md not found)"
    text = CHANGE_LOG_PATH.read_text(encoding="utf-8")
    matches = list(_CHANGE_LOG_HEADER_RE.finditer(text))
    if not matches:
        return "(no entries detected)"
    selected = matches[-CHANGE_LOG_TAIL:]
    start = selected[0].start()
    return text[start:].rstrip()


_TEST_FUNC_RE = re.compile(r"^(?:async\s+)?def\s+(\w+)\s*\(", re.MULTILINE)


def _is_live_test(nodeid: str) -> bool:
    """Return True iff the test referenced by ``nodeid`` still exists in source.

    ``nodeid`` looks like ``tests/x/y.py::test_func`` or
    ``tests/x/y.py::test_func[param]``. We strip the parametrize tail
    and verify (a) the file exists, (b) it defines a function with the
    matching name. Both pytest's ``lastfailed`` and ``nodeids`` caches
    can carry STALE entries from renamed / removed tests that pytest
    never gets to overwrite (it only updates entries it actually
    collected on the most recent run). Source-existence is the
    definitive check.
    """
    if "::" not in nodeid:
        return False
    file_part, _, rest = nodeid.partition("::")
    if not file_part.endswith(".py"):
        return False
    func_name = rest.split("[", 1)[0].split("::", 1)[0]
    test_file = _ANNUITY_DIR / file_part
    if not test_file.exists():
        return False
    try:
        text = test_file.read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError):
        return False
    return any(m.group(1) == func_name for m in _TEST_FUNC_RE.finditer(text))


def _cached_test_status() -> str:
    """Read pytest's lastfailed cache and summarize.

    pytest's ``lastfailed`` entries are STICKY: a previously-failing test
    that has since been renamed or removed stays in the cache forever
    because pytest never sees it pass. We verify each entry against the
    actual source tree (file exists + function name still defined) and
    split the failures into "live" (real signal -> the SME should look)
    and "stale" (renamed / removed -> ignore). Stricter than the prior
    nodeids cross-check because the nodeids cache itself can be stale.
    """
    if not PYTEST_CACHE_LASTFAILED.exists():
        return (
            "- Cache present: **no**\n"
            "- Hint: run `pytest tests/parity -q` from `annuity_model/` first; "
            "the SME relies on a recent run."
        )
    try:
        data = json.loads(PYTEST_CACHE_LASTFAILED.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return "- Cache present but malformed; rerun `pytest tests/parity -q`."
    if not data:
        return (
            "- Cache present: **yes**\n"
            "- Failed tests in last run: **0**\n"
            "- All last-run tests passed (or no tests have been run yet)."
        )
    all_failed = sorted(data.keys())
    live = [nid for nid in all_failed if _is_live_test(nid)]
    stale = [nid for nid in all_failed if nid not in live]
    live_bullets = "\n".join(f"  - `{nid}`" for nid in live[:20]) or "  - _(none)_"
    live_extra = "" if len(live) <= 20 else f"\n  - ... ({len(live) - 20} more)"
    parts = [
        "- Cache present: **yes**",
        f"- Live failing tests (still in source): **{len(live)}**",
        f"- Live failures:\n{live_bullets}{live_extra}",
    ]
    if stale:
        stale_bullets = "\n".join(f"  - `{nid}`" for nid in stale[:10])
        stale_extra = "" if len(stale) <= 10 else f"\n  - ... ({len(stale) - 10} more)"
        parts.append(f"- Stale entries (renamed / removed tests; safe to ignore): **{len(stale)}**")
        parts.append(f"- Stale nodeids:\n{stale_bullets}{stale_extra}")
    return "\n".join(parts)


def _benchmark_dump(scope: str, products: list[str]) -> str:
    """Dump actuarial_benchmarks constants, optionally filtered by product."""
    rows: list[str] = []
    keys = sorted(ab.__all__)
    if scope.startswith("product:"):
        wanted_prefix = scope.split(":", 1)[1].upper() + "_"
        keys = [k for k in keys if k.startswith(wanted_prefix)]
    elif scope == "incremental" and products:
        wanted_prefixes = tuple(f"{p.upper()}_" for p in products)
        keys = [k for k in keys if k.startswith(wanted_prefixes)]
    if not keys:
        keys = sorted(ab.__all__)
    for key in keys:
        value = getattr(ab, key)
        rows.append(f"- `{key}` = `{value!r}`")
    return "\n".join(rows)


def _spec_doc_pointers(products: list[str]) -> str:
    if not products:
        return "_(no specific product detected; consult per-product spec docs as needed)_"
    rows: list[str] = []
    for p in products:
        if p == "portfolio":
            spec = _ANNUITY_DIR / "docs" / "portfolio_runner_spec.md"
            parity_dir = _ANNUITY_DIR / "tests" / "parity" / "portfolio"
            spec_status = "exists" if spec.exists() else "missing"
            parity_status = "exists" if parity_dir.is_dir() else "missing"
            rows.append(
                f"- **portfolio**: spec=`docs/portfolio_runner_spec.md` ({spec_status}); "
                f"parity suite=`tests/parity/portfolio/` ({parity_status})"
            )
            continue
        spec = _ANNUITY_DIR / "docs" / f"{p}_product_spec.md"
        actuarial_test = _ANNUITY_DIR / "tests" / "parity" / f"test_{p}_actuarial.py"
        spec_status = "exists" if spec.exists() else "missing"
        test_status = "exists" if actuarial_test.exists() else "missing"
        rows.append(
            f"- **{p}**: spec=`docs/{p}_product_spec.md` ({spec_status}); "
            f"actuarial test=`tests/parity/test_{p}_actuarial.py` ({test_status})"
        )
    return "\n".join(rows)


def _build_diff_section(changed: list[str]) -> str:
    if not changed:
        return "_No working-tree changes vs HEAD._"
    parts: list[str] = []
    stat = _diff_stat()
    if stat:
        parts.append("### Stat (vs HEAD)\n```\n" + stat + "\n```")
    parts.append("### Per-file diffs (capped)\n")
    for path in changed:
        diff = _per_file_diff(path).rstrip()
        if not diff:
            continue
        parts.append(f"#### `{path}`\n```diff\n{diff}\n```")
    return "\n".join(parts)


def build_evidence_pack(
    *,
    scope: str,
    iteration: int,
    prior_verdict: str | None,
) -> str:
    changed = _changed_files()
    products = _detect_products(changed)
    return f"""# Actuary SME -- Evidence Pack

- **Generated:** {_utc_now()}
- **Scope:** `{scope}`
- **Iteration:** {iteration}
- **Prior verdict:** `{prior_verdict or "none"}`
- **Branch:** `{_branch()}`
- **HEAD:** `{_head_short()}`
- **Detected products in scope:** {", ".join(products) if products else "_(none detected from diff)_"}

> This pack is overwritten on every iteration of the SME loop. Verdicts are
> the historical record (under `.cursor/actuary-reviews/iter-N-*.md`); this
> file is the working input.

---

## 1. Git diff vs `HEAD`

{_build_diff_section(changed)}

---

## 2. Recent `docs/model_change_log.md` entries (tail)

{_change_log_tail()}

---

## 3. Cached parity-test status (from `.pytest_cache`)

{_cached_test_status()}

---

## 4. Actuarial benchmark constants in scope

{_benchmark_dump(scope, products)}

---

## 5. Per-product spec / actuarial-test pointers

{_spec_doc_pointers(products)}

---

## 6. SME guidance

Read this pack first. Then:

- Read `annuity_model/.cursor/skills/actuary-sme/SKILL.md` (your playbook).
- For each product in scope, read `annuity_model/docs/<P>_product_spec.md`
  and the relevant row in `annuity_model/docs/actuarial_benchmarks.md`.
- For each product in scope, read
  `annuity_model/tests/parity/test_<P>_actuarial.py` to see what the
  per-product test gate is asserting.
- If `iteration > 1`, **also read the prior verdict** at the path in
  the header above before re-judging. Verify each prior finding was
  addressed and populate `prior_findings_resolved[]` in your YAML
  frontmatter.

Emit your verdict per the SKILL.md template (YAML frontmatter +
markdown body). Return ONLY the verdict text -- no preamble, no
postamble.
"""


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        "--scope",
        required=True,
        help='Scope: "incremental", "full", "product:<code>", or "product:portfolio".',
    )
    parser.add_argument(
        "--iteration",
        type=int,
        required=True,
        help="Iteration number within the current loop (starts at 1).",
    )
    parser.add_argument(
        "--prior-verdict",
        default=None,
        help="Path (relative to repo root) to the prior iteration's verdict file.",
    )
    args = parser.parse_args(argv)

    EVIDENCE_DIR.mkdir(parents=True, exist_ok=True)
    pack = build_evidence_pack(
        scope=args.scope,
        iteration=args.iteration,
        prior_verdict=args.prior_verdict,
    )
    EVIDENCE_PATH.write_text(pack, encoding="utf-8")
    sys.stdout.write(f"{EVIDENCE_PATH.relative_to(_REPO_ROOT)}: written\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

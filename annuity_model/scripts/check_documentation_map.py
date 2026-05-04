"""Validate that DOCUMENTATION_MAP.md covers all tracked markdown docs.

Usage (from repo root):
    python annuity_model/scripts/check_documentation_map.py

The check enforces that every tracked `*.md` file is listed in
`DOCUMENTATION_MAP.md` as a bullet with a backticked path, and that the map
does not reference non-tracked markdown files.
"""

from __future__ import annotations

import re

# Reviewed: documentation map validation invokes a fixed local git command with shell=False.
import subprocess  # nosec B404
from pathlib import Path

MAP_PATH = Path("DOCUMENTATION_MAP.md")
PATH_PATTERN = re.compile(r"`([^`]+\.md)`")


def _tracked_markdown_files(repo_root: Path) -> set[str]:
    # Reviewed: fixed git command vector; partial path is acceptable for local developer tooling.
    out = subprocess.check_output(  # nosec B603 B607
        ["git", "ls-files", "*.md"],
        cwd=repo_root,
        text=True,
    )
    return {line.strip() for line in out.splitlines() if line.strip()}


def _mapped_markdown_files(repo_root: Path) -> set[str]:
    map_text = (repo_root / MAP_PATH).read_text()
    return set(PATH_PATTERN.findall(map_text))


def validate_documentation_map(repo_root: Path) -> list[str]:
    errors: list[str] = []
    map_file = repo_root / MAP_PATH
    if not map_file.exists():
        return [f"Missing required documentation map: {MAP_PATH}"]

    tracked = _tracked_markdown_files(repo_root)
    mapped = _mapped_markdown_files(repo_root)

    missing_in_map = sorted(tracked - mapped)
    stale_in_map = sorted(mapped - tracked)

    if missing_in_map:
        errors.append("Markdown files tracked by git but missing in DOCUMENTATION_MAP.md:")
        errors.extend(f"  - {p}" for p in missing_in_map)
    if stale_in_map:
        errors.append("Markdown paths listed in DOCUMENTATION_MAP.md but not tracked by git:")
        errors.extend(f"  - {p}" for p in stale_in_map)
    return errors


def main() -> int:
    repo_root = Path(__file__).resolve().parents[2]
    errors = validate_documentation_map(repo_root)
    if errors:
        print("\n".join(errors))
        return 1
    print("OK: DOCUMENTATION_MAP.md matches tracked markdown files.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

from __future__ import annotations

from pathlib import Path

from scripts.check_documentation_map import validate_documentation_map


def test_documentation_map_covers_tracked_markdown() -> None:
    repo_root = Path(__file__).resolve().parents[2]
    errors = validate_documentation_map(repo_root)
    assert not errors, "\n".join(errors)

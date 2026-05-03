#!/usr/bin/env python3
"""Audit ``st.session_state`` key usage across the Streamlit UI.

Why
---
``annuity_model/ui/MIGRATION.md`` is the plan for splitting
``pricing_ui.py`` (~4,500 LOC) into one Streamlit page per file under
``ui/pages/*``. Step 1 of that plan, written verbatim:

    > 3. **`session_state` discipline.** Every key minted by a page must
    >    be prefixed with the page name (e.g. ``pricing_run__contract_type``).
    >    The first PR of the migration adds a session-state audit script
    >    under ``scripts/``.

This is that script. It does NOT move any code; it produces the raw
inventory the per-page splits will rely on:

* **Per-page key map** -- for each ``_render_<page>`` function in
  ``pricing_ui.py``, the set of ``st.session_state`` keys it reads /
  writes / passes through to a Streamlit widget via ``key=``.
* **Cross-page keys** -- keys touched by more than one page. These are
  the actual landmines the migration must defuse before splitting.
* **Symbol vs literal breakdown** -- which references go through
  :class:`pricing_run_form_state.RUN_KEY` (good) vs raw string literals
  (bad; will fight the rename).

Output
------
Default: human-readable summary on stdout, exit 0.

``--json`` -- emits a machine-readable JSON report; exit 0 unless the
``--fail-on-cross-page`` flag is set, in which case the script returns
1 if any non-allow-listed key is shared across pages.

Usage::

    python scripts/audit_session_state.py
    python scripts/audit_session_state.py --json > audit.json
    python scripts/audit_session_state.py --fail-on-cross-page \\
        --allow-cross-page run_product_type run_issue_age

The third form is what the per-page-split PR's CI will call once
``ui/MIGRATION.md`` Step 3 is unblocked: a fixed allow-list of keys
that are *legitimately* shared (product type, demographic basics) and
zero tolerance for new accidental cross-page keys.
"""

from __future__ import annotations

import argparse
import ast
import json
import sys
from collections import defaultdict
from collections.abc import Iterable
from dataclasses import dataclass, field
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
PRICING_UI = REPO_ROOT / "pricing_ui.py"


@dataclass
class KeyUsage:
    """One observation of a session-state key inside a page function."""

    page: str
    key: str
    line: int
    via: str  # "subscript", "method", "widget_key", or "setdefault"
    is_literal: bool


@dataclass
class AuditReport:
    page_to_keys: dict[str, set[str]] = field(default_factory=lambda: defaultdict(set))
    key_to_pages: dict[str, set[str]] = field(default_factory=lambda: defaultdict(set))
    usages: list[KeyUsage] = field(default_factory=list)
    literal_count: int = 0
    symbol_count: int = 0

    def add(self, usage: KeyUsage) -> None:
        self.usages.append(usage)
        self.page_to_keys[usage.page].add(usage.key)
        self.key_to_pages[usage.key].add(usage.page)
        if usage.is_literal:
            self.literal_count += 1
        else:
            self.symbol_count += 1

    def cross_page_keys(self) -> dict[str, list[str]]:
        return {k: sorted(pages) for k, pages in self.key_to_pages.items() if len(pages) > 1}


def _enclosing_page(node: ast.AST, page_ranges: list[tuple[str, int, int]]) -> str | None:
    line = getattr(node, "lineno", None)
    if line is None:
        return None
    for name, start, end in page_ranges:
        if start <= line <= end:
            return name
    return None


def _walk_page_ranges(tree: ast.Module) -> list[tuple[str, int, int]]:
    """Identify every top-level ``_render_<page>`` function block.

    Returns ``(name, start_line, end_line)`` triples sorted by start.
    The end line is approximated as the last line of the function body
    via ``ast.walk``; close enough for the contiguous one-per-page
    layout in ``pricing_ui.py``.
    """
    out: list[tuple[str, int, int]] = []
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name.startswith("_render_"):
            end = max(
                (getattr(n, "lineno", node.lineno) for n in ast.walk(node)), default=node.lineno
            )
            out.append((node.name, node.lineno, end))
    out.sort(key=lambda t: t[1])
    return out


def _is_session_state_attr(value: ast.AST) -> bool:
    """Match ``st.session_state`` (Attribute) or a bare ``session_state``."""
    return (
        isinstance(value, ast.Attribute)
        and value.attr == "session_state"
        and isinstance(value.value, ast.Name)
        and value.value.id == "st"
    ) or (isinstance(value, ast.Name) and value.id == "session_state")


def _extract_string_or_attr(node: ast.AST) -> tuple[str | None, bool]:
    """Return ``(key_name, is_literal)`` if *node* is a string literal or
    a ``RUN_KEY.X`` attribute; ``(None, False)`` otherwise."""
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return node.value, True
    if isinstance(node, ast.Attribute) and isinstance(node.value, ast.Name):
        # RUN_KEY.X -> resolve to the literal value via the runtime
        # import so we record the *underlying* key, not the symbol path.
        try:
            import sys as _sys

            _sys.path.insert(0, str(REPO_ROOT))
            import pricing_run_form_state as prfs

            holder = getattr(prfs, node.value.id, None)
            if holder is not None:
                val = getattr(holder, node.attr, None)
                if isinstance(val, str):
                    return val, False
        except Exception:
            return None, False
    return None, False


def _audit(tree: ast.Module) -> AuditReport:
    page_ranges = _walk_page_ranges(tree)
    rep = AuditReport()

    for node in ast.walk(tree):
        page = _enclosing_page(node, page_ranges)
        if page is None:
            continue

        # st.session_state["key"]
        if isinstance(node, ast.Subscript) and _is_session_state_attr(node.value):
            key, is_lit = _extract_string_or_attr(node.slice)
            if key:
                rep.add(
                    KeyUsage(
                        page=page, key=key, line=node.lineno, via="subscript", is_literal=is_lit
                    )
                )
            continue

        # st.session_state.get("key", ...) / .setdefault("key", ...) /
        # .pop("key", ...) / .update({...})
        if (
            isinstance(node, ast.Call)
            and isinstance(node.func, ast.Attribute)
            and _is_session_state_attr(node.func.value)
        ):
            if node.func.attr in {"get", "setdefault", "pop"} and node.args:
                key, is_lit = _extract_string_or_attr(node.args[0])
                if key:
                    rep.add(
                        KeyUsage(
                            page=page,
                            key=key,
                            line=node.lineno,
                            via=f"method:{node.func.attr}",
                            is_literal=is_lit,
                        )
                    )
            continue

        # st.<widget>(..., key="...") / RUN_KEY.X
        if isinstance(node, ast.Call):
            if (
                isinstance(node.func, ast.Attribute)
                and isinstance(node.func.value, ast.Name)
                and node.func.value.id in {"st"}
            ):
                for kw in node.keywords:
                    if kw.arg == "key":
                        key, is_lit = _extract_string_or_attr(kw.value)
                        if key:
                            rep.add(
                                KeyUsage(
                                    page=page,
                                    key=key,
                                    line=node.lineno,
                                    via="widget_key",
                                    is_literal=is_lit,
                                )
                            )
            # Custom helpers like run_number_input(..., key=...)
            if isinstance(node.func, ast.Name) and node.func.id.endswith("_input"):
                for kw in node.keywords:
                    if kw.arg == "key":
                        key, is_lit = _extract_string_or_attr(kw.value)
                        if key:
                            rep.add(
                                KeyUsage(
                                    page=page,
                                    key=key,
                                    line=node.lineno,
                                    via="widget_key",
                                    is_literal=is_lit,
                                )
                            )

    return rep


def _print_human(rep: AuditReport, top_n: int = 10) -> None:
    try:
        display_path = PRICING_UI.relative_to(REPO_ROOT)
    except ValueError:
        display_path = PRICING_UI
    print(f"== session_state audit of {display_path} ==\n")
    print(f"Total observations:   {len(rep.usages)}")
    print(f"  via raw literal:    {rep.literal_count}")
    print(f"  via RUN_KEY symbol: {rep.symbol_count}")
    print(f"Unique keys:          {len(rep.key_to_pages)}")
    print(f"Pages observed:       {len(rep.page_to_keys)}\n")

    print("Per-page key counts:")
    for page in sorted(rep.page_to_keys):
        keys = rep.page_to_keys[page]
        print(f"  {page}: {len(keys)} unique keys")
    print()

    cross = rep.cross_page_keys()
    print(f"Cross-page keys ({len(cross)}):")
    for k in sorted(cross, key=lambda kk: -len(cross[kk]))[:top_n]:
        pages = cross[k]
        print(f"  {k}: shared by {len(pages)} pages -> {', '.join(pages)}")
    if len(cross) > top_n:
        print(f"  ... and {len(cross) - top_n} more.")


def _to_json(rep: AuditReport) -> dict[str, object]:
    return {
        "totals": {
            "observations": len(rep.usages),
            "literal": rep.literal_count,
            "symbol": rep.symbol_count,
            "unique_keys": len(rep.key_to_pages),
            "pages": len(rep.page_to_keys),
        },
        "per_page": {page: sorted(keys) for page, keys in sorted(rep.page_to_keys.items())},
        "cross_page": rep.cross_page_keys(),
    }


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--json", action="store_true", help="emit JSON instead of text.")
    parser.add_argument(
        "--fail-on-cross-page",
        action="store_true",
        help=(
            "exit non-zero if any non-allow-listed key is shared across "
            "pages. Use in CI once the migration begins so a new "
            "shared key cannot sneak in."
        ),
    )
    parser.add_argument(
        "--allow-cross-page",
        nargs="*",
        default=[],
        help="keys legitimately shared across pages (e.g. product type).",
    )
    args = parser.parse_args(list(argv) if argv is not None else None)

    if not PRICING_UI.is_file():
        print(f"missing {PRICING_UI}", file=sys.stderr)
        return 2
    tree = ast.parse(PRICING_UI.read_text(), filename=str(PRICING_UI))
    rep = _audit(tree)

    if args.json:
        print(json.dumps(_to_json(rep), indent=2))
    else:
        _print_human(rep)

    if args.fail_on_cross_page:
        bad = {
            k: pages
            for k, pages in rep.cross_page_keys().items()
            if k not in set(args.allow_cross_page)
        }
        if bad:
            print(
                f"\n[audit] FAIL: {len(bad)} cross-page key(s) not in --allow-cross-page:",
                file=sys.stderr,
            )
            for k, pages in bad.items():
                print(f"  {k}: {pages}", file=sys.stderr)
            return 1
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())

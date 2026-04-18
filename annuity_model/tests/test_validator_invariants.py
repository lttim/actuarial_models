"""Validator / structural invariants enforced at the AST and registry level.

These tests prevent the entire class of bug where someone removes a call to
``validate_workbook_or_raise`` "to debug" and forgets to put it back, or adds a
new product to ``ProductType`` without an entry in ``LIABILITY_LAYOUTS``.

If any assertion here fails, the fix is **always** in the offending source
file -- never in this test.
"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

from liability_layouts import LIABILITY_LAYOUTS, liability_layout_for
from product_registry import ProductType, implemented_product_types

PACKAGE_ROOT = Path(__file__).resolve().parent.parent
BUILDER_FILES = sorted(PACKAGE_ROOT.glob("build_*_excel_workbook.py"))

pytestmark = pytest.mark.invariant


# ---------------------------------------------------------------------------
# 1. Every wb.save(...) is preceded by validate_workbook_or_raise(wb).
# ---------------------------------------------------------------------------


def _save_call_function(node: ast.AST) -> ast.AST | None:
    """Yield the enclosing function for a save() call (None if module-level)."""
    return getattr(node, "_parent_function", None)


class _SaveValidatorVisitor(ast.NodeVisitor):
    """Walk a builder module and collect (function, save_call, validate_calls)."""

    def __init__(self) -> None:
        self.violations: list[str] = []
        self._function_stack: list[ast.FunctionDef | ast.AsyncFunctionDef] = []

    def visit_FunctionDef(self, node: ast.FunctionDef) -> None:  # noqa: N802
        self._check_function(node)

    def visit_AsyncFunctionDef(self, node: ast.AsyncFunctionDef) -> None:  # noqa: N802
        self._check_function(node)

    def _check_function(self, fn: ast.FunctionDef | ast.AsyncFunctionDef) -> None:
        save_calls: list[ast.Call] = []
        validate_calls: list[ast.Call] = []
        for sub in ast.walk(fn):
            if not isinstance(sub, ast.Call):
                continue
            attr = getattr(sub.func, "attr", None)
            name = getattr(sub.func, "id", None)
            if attr == "save":
                save_calls.append(sub)
            if name == "validate_workbook_or_raise" or attr == "validate_workbook_or_raise":
                validate_calls.append(sub)
        if not save_calls:
            return
        if not validate_calls:
            self.violations.append(
                f"function `{fn.name}` calls `.save(...)` at line "
                f"{save_calls[0].lineno} but does NOT call "
                "`validate_workbook_or_raise` anywhere in the same function."
            )
            return
        # Validate that *every* save is preceded (in source order) by at least
        # one validate call within the same function.
        first_validate_line = min(c.lineno for c in validate_calls)
        for save in save_calls:
            if save.lineno <= first_validate_line:
                self.violations.append(
                    f"function `{fn.name}` calls `.save(...)` at line "
                    f"{save.lineno} BEFORE the first "
                    f"`validate_workbook_or_raise(...)` at line "
                    f"{first_validate_line}."
                )


@pytest.mark.parametrize("builder", BUILDER_FILES, ids=lambda p: p.name)
def test_every_workbook_save_is_validator_gated(builder: Path) -> None:
    """Every ``wb.save(...)`` in a builder must have a same-function
    ``validate_workbook_or_raise(wb)`` call earlier in the function body."""
    if not builder.exists():
        pytest.skip(f"builder {builder} not present")
    tree = ast.parse(builder.read_text(encoding="utf-8"), filename=str(builder))
    visitor = _SaveValidatorVisitor()
    visitor.visit(tree)
    assert not visitor.violations, (
        f"{builder.name}: validator-gating invariant violated:\n  - "
        + "\n  - ".join(visitor.violations)
    )


# ---------------------------------------------------------------------------
# 2. Every ProductType has a LIABILITY_LAYOUTS entry.
# ---------------------------------------------------------------------------


def test_every_implemented_product_has_a_liability_layout() -> None:
    """Every product registered as *implemented* must carry a layout.

    ``ProductType`` may legitimately list placeholder products (whole_life,
    variable_annuity) that are reserved enum values without a builder. The
    invariant we care about is that everything *implemented* is covered.
    """
    impl_codes = {pt.value for pt in implemented_product_types()}
    layout_codes = set(LIABILITY_LAYOUTS)
    missing = impl_codes - layout_codes
    assert not missing, (
        f"Implemented products without a LIABILITY_LAYOUTS entry: {sorted(missing)}. "
        "Add an entry to liability_layouts.LIABILITY_LAYOUTS."
    )


def test_no_orphan_layouts() -> None:
    """A ``LIABILITY_LAYOUTS`` entry must correspond to a real ProductType.

    Catches typos like ``"spia "`` (trailing space) or stale codes left after
    a product was renamed.
    """
    valid_codes = {pt.value for pt in ProductType}
    orphans = set(LIABILITY_LAYOUTS) - valid_codes
    assert not orphans, (
        f"LIABILITY_LAYOUTS contains codes that are not ProductType values: {sorted(orphans)}"
    )


# ---------------------------------------------------------------------------
# 3. End-to-end builder dispatch coverage.
# ---------------------------------------------------------------------------


def test_product_excel_dispatcher_covers_every_implemented_product() -> None:
    """``build_product_workbook`` must dispatch every implemented ProductType.

    A new product added to the registry without an entry in the dispatcher
    would silently raise ``NotImplementedError`` at runtime. AST-walk the
    dispatcher and verify each implemented ``ProductType`` appears in its
    body.
    """
    src_path = PACKAGE_ROOT / "product_excel.py"
    src = src_path.read_text(encoding="utf-8")
    impl = implemented_product_types()
    missing: list[str] = []
    for pt in impl:
        needle_a = f"ProductType.{pt.name}"
        needle_b = f'"{pt.value}"'
        if needle_a not in src and needle_b not in src:
            missing.append(pt.value)
    assert not missing, (
        f"product_excel.build_product_workbook does not dispatch for: {missing}. "
        "Add a branch (or registry-driven dispatch) so the new product can be exported."
    )

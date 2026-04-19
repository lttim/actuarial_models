"""Meta-invariant: the mypy strict-mode glob actually covers every product.

We replaced the manually-curated ``[[tool.mypy.overrides]]`` strict-list
with a glob over ``products.*.engine`` (and friends) so that adding a
new product does not require a separate "remember to add it to mypy
strict" commit. This test pins down the contract: every product
subpackage that exists on disk must be matched by the glob, AND every
load-bearing core module (legacy flat layout) must still be in the
hand-curated strict list.

If this test fails, you almost certainly:

* Added a new ``annuity_model/products/<name>/`` directory but mypy's
  override block no longer matches it (someone tightened the glob), OR
* Removed a module from the load-bearing strict list without adding a
  replacement (regression in type-coverage), OR
* Renamed a product subpackage and the glob no longer fits.

Fix: re-widen the glob in ``pyproject.toml``, do NOT skip this test.
"""

from __future__ import annotations

import fnmatch
import sys
import tomllib
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
PYPROJECT = REPO_ROOT / "pyproject.toml"
PRODUCTS_DIR = REPO_ROOT / "products"

# Load-bearing core modules: every parity-critical legacy module must
# remain mypy-strict. If a module is renamed / split, update both the
# pyproject override block AND this list.
LOAD_BEARING_CORE: tuple[str, ...] = (
    "pricing_projection",
    "term_projection",
    "rila_projection",
    "alm_excel_ladder",
    "build_pricing_excel_workbook",
    "build_term_excel_workbook",
    "build_rila_excel_workbook",
    "excel_builder_helpers",
    "excel_workbook_validator",
    "product_registry",
    "product_excel",
    "liability_dispatch",
    "liability_layouts",
)

# Per-product files that the glob must cover. Adding a new file here
# requires updating the override block in pyproject.toml at the same time.
PRODUCT_SUBMODULES: tuple[str, ...] = ("engine", "excel", "schema", "ui")


def _load_strict_module_patterns() -> list[str]:
    data = tomllib.loads(PYPROJECT.read_text())
    overrides = data.get("tool", {}).get("mypy", {}).get("overrides", [])
    patterns: list[str] = []
    for block in overrides:
        if not _is_strict_block(block):
            continue
        modules = block.get("module", [])
        if isinstance(modules, str):
            patterns.append(modules)
        else:
            patterns.extend(modules)
    return patterns


def _is_strict_block(block: dict) -> bool:
    """A block is 'strict' if it sets the canonical strict flags."""
    return (
        block.get("disallow_untyped_defs") is True
        and block.get("disallow_incomplete_defs") is True
        and block.get("check_untyped_defs") is True
    )


def _matches_any(name: str, patterns: list[str]) -> bool:
    """Mimic mypy's glob semantics: `*` matches a single dotted segment."""
    return any(fnmatch.fnmatchcase(name, p) for p in patterns)


@pytest.fixture(scope="module")
def strict_patterns() -> list[str]:
    if sys.version_info < (3, 11):
        pytest.skip("tomllib requires Python 3.11+")
    return _load_strict_module_patterns()


@pytest.mark.parametrize("module_name", LOAD_BEARING_CORE)
def test_load_bearing_core_is_strict(strict_patterns: list[str], module_name: str) -> None:
    assert _matches_any(module_name, strict_patterns), (
        f"Load-bearing module {module_name!r} is NOT covered by any strict "
        f"mypy override in pyproject.toml. This is a regression in type "
        f"coverage on a parity-critical module. Add it back to the "
        f"hand-curated `module` list in [[tool.mypy.overrides]]."
    )


def _discover_product_subpackages() -> list[str]:
    if not PRODUCTS_DIR.is_dir():
        return []
    out: list[str] = []
    for child in sorted(PRODUCTS_DIR.iterdir()):
        if not child.is_dir():
            continue
        if child.name.startswith("_") or child.name.startswith("."):
            continue
        if not (child / "__init__.py").is_file():
            continue
        out.append(child.name)
    return out


def test_product_directory_is_nonempty() -> None:
    """Sanity: at least SPIA / Term / RILA must exist."""
    found = _discover_product_subpackages()
    assert found, f"No product subpackages found under {PRODUCTS_DIR}."
    for required in ("spia", "term", "rila"):
        assert required in found, (
            f"Expected products/{required}/ subpackage; found only {found!r}. "
            "If you renamed a product directory, update LOAD_BEARING_CORE / "
            "this assertion accordingly."
        )


@pytest.mark.parametrize("submodule", PRODUCT_SUBMODULES)
def test_per_product_glob_covers_all_products(
    strict_patterns: list[str], submodule: str
) -> None:
    """Every products/<name>/<submodule>.py must be matched by the glob."""
    discovered = _discover_product_subpackages()
    missing: list[str] = []
    for product in discovered:
        # Only check submodules that actually exist on disk -- a product
        # is allowed to omit a shim if it has no Streamlit UI yet, etc.
        path = PRODUCTS_DIR / product / f"{submodule}.py"
        if not path.is_file():
            continue
        dotted = f"products.{product}.{submodule}"
        if not _matches_any(dotted, strict_patterns):
            missing.append(dotted)
    assert not missing, (
        f"The mypy strict override glob does not cover these product "
        f"submodules: {missing!r}. Re-widen the override glob in "
        f"pyproject.toml (the canonical pattern is `products.*.{submodule}`)."
    )


def test_no_redundant_product_modules_in_load_bearing_list() -> None:
    """Don't list `products.spia.engine` twice -- the glob already covers it.

    This guards against a contributor seeing a mypy strict failure on
    products.<x>.engine and "fixing" it by adding the literal name to
    LOAD_BEARING_CORE, which would silently re-introduce the old
    manual-list maintenance burden.
    """
    bad = [m for m in LOAD_BEARING_CORE if m.startswith("products.")]
    assert not bad, (
        f"LOAD_BEARING_CORE must not contain products.* entries (these are "
        f"covered by the glob override). Found: {bad!r}."
    )

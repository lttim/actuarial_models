"""Static-AST guard: every name ``pricing_ui.py`` imports from a local
sibling module MUST actually be defined in that module.

Why this file exists
--------------------
On 2026-04-18 a Streamlit user double-clicked ``run_pricing_ui.command``
and got::

    File "annuity_model/pricing_ui.py", line 66, in <module>
        from annuity_model.product_registry import (
    ImportError: cannot import name 'parse_term_benefit_timing_label'
        from 'product_registry' (.../product_registry.py)

The traceback has no frame inside ``product_registry`` itself --
``parse_term_benefit_timing_label`` *was* defined in the file at HEAD,
but the running interpreter held a **stale cached module** in
``sys.modules`` from an earlier launch (Streamlit's hot-reload re-executes
``pricing_ui.py`` via ``exec(code, module.__dict__)`` but does NOT
re-import its dependencies). When a PR adds a new import line in
``pricing_ui.py`` *and* the new symbol in ``product_registry.py``
together, an already-running Streamlit instance sees the new import
against the old cached module and crashes for every user that had the
app open across the upgrade.

The existing tests do not catch this class of bug:

* ``test_pricing_ui_imports_under_supported_python`` does
  ``import pricing_ui`` -- but pytest's collection already imported
  ``pricing_ui`` via ``tests/test_pricing_ui_term_config.py``, so the
  smoke test is a cache hit.
* ``tests/ui/test_apptest_full_workflow.test_pricing_ui_boots_without_exception``
  uses Streamlit's ``AppTest`` which always starts with a fresh script
  runner -- it would catch a real "name truly missing" bug, but only
  after Streamlit has fully bootstrapped (pulling in 200+ modules first).
* ``tests/test_launcher_invariants.test_shell_launcher_self_check_with_clean_path``
  does spawn a clean subprocess, but it requires the project ``.venv``
  to be present and runs ``import pricing_ui`` indirectly via the bash
  launcher. It is not parameterised per imported name so a failure
  reports "launcher exited 1" rather than "pricing_ui imports name X
  from product_registry that does not exist there".

The check below is the missing layer: a fast, deterministic, static AST
walk that flags any pricing_ui import line whose name is not a top-level
definition of the source module *before* the bug ever ships. It
intentionally does NOT execute ``pricing_ui`` -- that is what the AppTest
gate is for. It only inspects source files, so it cannot be defeated by
``sys.modules`` caching, by lazy imports, by Streamlit's reload
behaviour, or by an out-of-tree ``__pycache__`` poisoning.

If this test fails
------------------
Either:

1. ``pricing_ui.py`` has a stale import line for a symbol that was
   removed/renamed in the source module -- delete or update the import.
2. The source module forgot to define a symbol that ``pricing_ui.py``
   genuinely needs -- add the definition.

In both cases the fix lives in the module the test names, not in this
file. Do NOT relax the guard.
"""

from __future__ import annotations

import ast
import os
import subprocess
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
PACKAGE_ROOT = PROJECT_ROOT / "src" / "annuity_model"
PRICING_UI = PACKAGE_ROOT / "pricing_ui.py"

pytestmark = [pytest.mark.invariant]


# ---------------------------------------------------------------------------
# Which "from X import Y" lines do we audit?
#
# We only audit imports from sibling project modules (same directory as
# pricing_ui.py). Third-party imports (streamlit, numpy, ...) are out of
# scope -- their resolution is the package manager's job, not ours.
# Imports from project SUBPACKAGES (``products.term.ui``, ``tests.*``)
# are also out of scope here -- the per-product subpackage shims have
# their own identity-equality gate in
# ``tests/test_products_subpackage_shims.py``.
# ---------------------------------------------------------------------------


def _local_sibling_module_names() -> set[str]:
    """Return the set of module names that live next to ``pricing_ui.py``."""
    return {path.stem for path in PACKAGE_ROOT.glob("*.py") if path.name != "__init__.py"}


def _toplevel_defined_names(source: str) -> set[str]:
    """Return every name a ``from <module> import X`` could legitimately resolve.

    A name is "defined" at module top-level if it is:

    * a function or async function (``def`` / ``async def``),
    * a class (``class``),
    * a target of an assignment (``X = ...`` or ``X: T = ...``),
    * brought in via ``import X``, ``import X as Y``, or
      ``from M import X`` / ``from M import X as Y``.

    We deliberately do NOT execute the module -- this keeps the check
    free of the very ``sys.modules`` caching the original bug exploited.
    """
    tree = ast.parse(source)
    names: set[str] = set()
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            names.add(node.name)
        elif isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name):
                    names.add(target.id)
                elif isinstance(target, (ast.Tuple, ast.List)):
                    for elt in target.elts:
                        if isinstance(elt, ast.Name):
                            names.add(elt.id)
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
            names.add(node.target.id)
        elif isinstance(node, ast.Import):
            for alias in node.names:
                names.add(alias.asname or alias.name.split(".", 1)[0])
        elif isinstance(node, ast.ImportFrom):
            for alias in node.names:
                if alias.name == "*":
                    # Star imports are opaque to AST analysis; we cannot
                    # confirm or deny that *anything* is re-exported, so
                    # skip rather than risk false positives.
                    continue
                names.add(alias.asname or alias.name)
    return names


def _pricing_ui_local_imports() -> list[tuple[str, str, int]]:
    """Return ``(module, name, lineno)`` triples for every local-module import."""
    src = PRICING_UI.read_text(encoding="utf-8")
    tree = ast.parse(src, filename=str(PRICING_UI))
    locals_ = _local_sibling_module_names()
    triples: list[tuple[str, str, int]] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom):
            if node.level != 0:
                continue
            module = node.module or ""
            module_tail = module.removeprefix("annuity_model.")
            if module_tail not in locals_:
                continue
            for alias in node.names:
                if alias.name == "*":
                    continue
                triples.append((module_tail, alias.name, node.lineno))
    return triples


# ---------------------------------------------------------------------------
# Test 1: every "from <local_module> import X" name resolves statically.
# Parametrised so a regression names the EXACT (module, symbol) pair.
# ---------------------------------------------------------------------------

_LOCAL_IMPORT_TRIPLES = _pricing_ui_local_imports()


def test_pricing_ui_has_local_imports_to_audit() -> None:
    """Sanity check: the AST walk found imports to validate.

    If this fails, either ``pricing_ui.py`` was deleted, or every
    ``from X import Y`` line uses a non-sibling module path -- in which
    case this whole guard is silently a no-op and someone has to fix
    the discovery. We assert >= 3 because the file historically imports
    from at least ``product_registry``, ``product_excel``, and
    ``pricing_run_form_state``.
    """
    assert len(_LOCAL_IMPORT_TRIPLES) >= 3, (
        f"Found only {len(_LOCAL_IMPORT_TRIPLES)} local-module imports "
        f"in {PRICING_UI.name}; the static AST audit cannot work if "
        "discovery returns nothing. Did the file get deleted, renamed, "
        "or refactored to import every dependency through a single "
        "facade module?"
    )


@pytest.mark.parametrize(
    "module,name,lineno",
    _LOCAL_IMPORT_TRIPLES,
    ids=[f"{m}.{n}@L{ln}" for m, n, ln in _LOCAL_IMPORT_TRIPLES],
)
def test_pricing_ui_local_import_name_resolves_statically(
    module: str, name: str, lineno: int
) -> None:
    """``pricing_ui.py`` line *lineno* says ``from {module} import {name}``;
    that name MUST be defined at the top level of ``{module}.py``.

    A failure here means the import will raise ``ImportError`` at runtime
    on a fresh interpreter -- exactly the 2026-04-18 incident this whole
    file exists to prevent.
    """
    module_path = PACKAGE_ROOT / f"{module}.py"
    assert module_path.is_file(), (
        f"{PRICING_UI.name}:{lineno} imports from {module!r} but "
        f"{module_path} does not exist. Either delete the import or "
        "create the module."
    )
    defined = _toplevel_defined_names(module_path.read_text(encoding="utf-8"))
    assert name in defined, (
        f"{PRICING_UI.name}:{lineno}  `from {module} import {name}`  "
        f"FAILS: {module}.py does not define a top-level `{name}` "
        "(no def/class/assignment/re-export with that name). "
        "This is the static analog of the 2026-04-18 ImportError that "
        "Streamlit users hit when hot-reload re-executed pricing_ui.py "
        "against a stale cached product_registry. "
        f"Either remove the import line or add `{name}` to {module}.py."
    )


# ---------------------------------------------------------------------------
# Test 2: end-to-end fresh-subprocess import.
#
# The static AST guard above catches the *direct* import line. A
# transitively-required symbol (e.g. ``products.term.ui`` re-exporting
# something that no longer exists in product_registry) would also crash
# at script load but not be visible at the AST layer of pricing_ui.py
# alone. This test runs ``python -c "import pricing_ui"`` in a clean
# subprocess so pytest's already-cached pricing_ui module CANNOT mask
# the failure.
# ---------------------------------------------------------------------------


def test_pricing_ui_imports_cleanly_in_a_fresh_subprocess() -> None:
    """``import pricing_ui`` must succeed in a brand-new Python process.

    pytest's collection already imports ``pricing_ui`` once (via
    ``tests/test_pricing_ui_term_config.py``), so any in-process
    ``import pricing_ui`` is a cache hit and proves nothing about a
    real launcher invocation. Spawning a subprocess with the project
    ``.venv`` interpreter (or the test interpreter as a fallback)
    forces the full import chain to execute from scratch.

    The launcher's own ``--self-check`` mode has the same property,
    but is gated by ``test_shell_launcher_self_check_with_clean_path``
    which skips when ``.venv`` is missing. This test is the
    interpreter-only complement that runs everywhere.
    """
    venv_python = PROJECT_ROOT / ".venv" / "bin" / "python"
    py = str(venv_python) if venv_python.exists() else sys.executable
    env = os.environ.copy()
    env["PYTHONPATH"] = str(PROJECT_ROOT / "src")

    result = subprocess.run(
        [py, "-c", "import annuity_model.pricing_ui"],
        cwd=str(PROJECT_ROOT),
        env=env,
        capture_output=True,
        text=True,
        timeout=120,
    )
    assert result.returncode == 0, (
        "`python -c 'import pricing_ui'` failed in a fresh subprocess. "
        "This is the exact failure mode a user double-clicking "
        "run_pricing_ui.command would see. "
        f"\nInterpreter: {py}"
        f"\nstdout:\n{result.stdout}"
        f"\nstderr:\n{result.stderr}"
    )

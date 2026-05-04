"""Compatibility shim for ``import annuity_model`` from ``annuity_model/`` cwd.

When Python starts inside this directory, the real package cannot be resolved
because its parent directory is not on ``sys.path``.  This module is found in
that specific working-directory mode and re-exports the package public surface
from ``__init__.py`` while the project remains in the legacy flat layout.
"""

from __future__ import annotations

import runpy
import sys
from pathlib import Path
from types import ModuleType

_PACKAGE_DIR = Path(__file__).resolve().parent
_PACKAGE_DIR_STR = str(_PACKAGE_DIR)
if _PACKAGE_DIR_STR not in sys.path:
    sys.path.insert(0, _PACKAGE_DIR_STR)

_namespace = runpy.run_path(str(_PACKAGE_DIR / "__init__.py"), run_name="_annuity_model_compat")
__all__ = list(_namespace["__all__"])
__doc__ = _namespace.get("__doc__", __doc__)
__path__ = [_PACKAGE_DIR_STR]

for _name in __all__:
    globals()[_name] = _namespace[_name]

for _flat_name in (
    "_logging",
    "excel_workbook_validator",
    "liability_aggregation",
    "liability_layouts",
    "portfolio",
    "portfolio_runner",
    "pricing_projection",
    "product_excel",
    "product_registry",
    "rila_projection",
    "term_projection",
):
    _module = sys.modules.get(_flat_name)
    if isinstance(_module, ModuleType):
        sys.modules.setdefault(f"{__name__}.{_flat_name}", _module)

for _loaded_name, _module in list(sys.modules.items()):
    if _loaded_name == "products" or _loaded_name.startswith("products."):
        sys.modules.setdefault(f"{__name__}.{_loaded_name}", _module)

del _flat_name, _loaded_name, _module, _name, _namespace

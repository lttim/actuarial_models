"""Per-product liability-sheet column layouts for ALM linkage.

The ALM ladder helpers in :mod:`alm_excel_ladder` and
:mod:`build_pricing_excel_workbook` need to know which column on each
product's *Liabilities* sheet holds the expected total cashflow ("ExpTotalCF")
and the discount factor used to PV liabilities.

For SPIA and Term these are columns ``S`` and ``O``. For RILA they are ``M``
and ``O`` (the RILA liability grid has fewer expense columns, so ExpTotalCF
shifts left). Hard-coding the literal letters in three different builders
caused the 2026-03 RILA-S/M bug; centralizing them here gives the validator
one source of truth and lets meta-tests assert that every product type in
the registry has a layout.

This module is intentionally **standalone** (no project imports) so that
:mod:`product_registry` can import it without forming a cycle. Layouts are
keyed by the ``ProductType.value`` string code (``"spia"``, ``"term_life"``,
``"rila"``); :func:`liability_layout_for` accepts either a string or a
``ProductType`` enum member (resolved by ``.value`` at call time).

Adding a new product (FIA, VA-GLWB, ...) means one entry in
:data:`LIABILITY_LAYOUTS` -- the builder reads it via
:func:`liability_layout_for` and never needs literal column letters in code.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class LiabilityLayout:
    """Column letters for the *Liabilities* sheet of a given product.

    Attributes
    ----------
    product_code:
        ``ProductType.value`` (e.g. ``"spia"``) the layout applies to.
    total_cf_col:
        Column letter (e.g. ``"S"``) holding the per-month *expected total
        cashflow* (benefits + expenses, alive-weighted). The ALM ladder reads
        this column for liability disinvestment cash needs.
    discount_col:
        Column letter (e.g. ``"O"``) holding the per-month discount factor
        used to PV future liability cashflows.
    """

    product_code: str
    total_cf_col: str
    discount_col: str

    def __post_init__(self) -> None:
        cf = (self.total_cf_col or "").strip().upper()
        df = (self.discount_col or "").strip().upper()
        if not cf.isalpha() or not df.isalpha() or len(cf) > 2 or len(df) > 2:
            raise ValueError(
                f"LiabilityLayout column letters must be 1-2 alphabetic chars; "
                f"got total_cf_col={self.total_cf_col!r}, discount_col={self.discount_col!r}"
            )
        # Frozen dataclass -- assign normalized values via object.__setattr__.
        object.__setattr__(self, "total_cf_col", cf)
        object.__setattr__(self, "discount_col", df)


# Single source of truth. Keys are ProductType.value strings (kept in sync
# with the enum literals in product_registry.py). The validator and parity
# tests read this map; never duplicate these letters elsewhere in the codebase.
LIABILITY_LAYOUTS: Mapping[str, LiabilityLayout] = {
    "spia": LiabilityLayout(product_code="spia", total_cf_col="S", discount_col="O"),
    "term_life": LiabilityLayout(product_code="term_life", total_cf_col="S", discount_col="O"),
    "rila": LiabilityLayout(product_code="rila", total_cf_col="M", discount_col="O"),
}


def liability_layout_for(product) -> LiabilityLayout:  # noqa: ANN001 -- accepts ProductType OR str
    """Return the registered liability layout for *product*.

    Accepts either a :class:`ProductType` enum member or its
    ``.value`` string code. Resolving by value at call time avoids a circular
    import with :mod:`product_registry`.

    Raises
    ------
    KeyError
        If no layout is registered. The Phase 4 meta-test asserts that every
        ProductType has a layout, so missing entries surface in CI rather
        than at runtime.
    """
    code = getattr(product, "value", product)
    if not isinstance(code, str):
        raise TypeError(
            f"liability_layout_for() expected ProductType or str, got {type(product).__name__}"
        )
    try:
        return LIABILITY_LAYOUTS[code]
    except KeyError as exc:
        raise KeyError(
            f"No LiabilityLayout registered for product code {code!r}. "
            f"Add an entry to LIABILITY_LAYOUTS in liability_layouts.py."
        ) from exc


__all__ = [
    "LIABILITY_LAYOUTS",
    "LiabilityLayout",
    "liability_layout_for",
]

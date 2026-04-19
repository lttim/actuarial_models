"""Parity test: actuarial_parity_kit/AGENTS_template.md vs annuity_model/AGENTS.md.

The kit's ``AGENTS_template.md`` is what every *new* product repo will be
seeded from. If it drifts behind ``annuity_model/AGENTS.md`` -- which is
the canonical "what do I run before claiming a task is done?" doc -- new
repos get spawned without the same gate discipline.

This test enforces that the template stays in lock-step with the canonical
doc on the things that actually block merges:

* the four canonical gate commands (parity / pytest / deep_smoke /
  render_parity_contract --check),
* the explicit "Never widen a tolerance" rule,
* the static workbook-validator rule,
* the parity contract / change log file references,
* the canonical-gates header.

We deliberately do NOT compare the documents byte-for-byte: the template
has placeholders (``[PRODUCT NAME]``) that obviously don't appear in the
SPIA repo. We compare specific *non-negotiable* fragments instead.
"""

from __future__ import annotations

from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
CANONICAL_AGENTS = REPO_ROOT / "annuity_model" / "AGENTS.md"
KIT_TEMPLATE = REPO_ROOT / "actuarial_parity_kit" / "AGENTS_template.md"


# Each entry is (label, fragment). The fragment must appear verbatim in
# both files. Pick fragments tight enough to be diagnostic but loose
# enough that incidental wording tweaks don't trigger spurious failures.
REQUIRED_FRAGMENTS: list[tuple[str, str]] = [
    ("canonical-gates header", "## Before completing any task -- canonical gates"),
    ("single source of truth claim", "*single source of truth*"),
    ("gate 1: parity", "python -m pytest tests/parity -q"),
    ("gate 2: full pytest", "python -m pytest -q"),
    ("gate 3: deep smoke", "python scripts/deep_smoke.py"),
    (
        "gate 4: tolerance contract",
        "python scripts/render_parity_contract.py --check",
    ),
    ("all-four-must-exit-0", "All four must exit 0."),
    ("never-widen-tolerance rule", "**Never widen a tolerance to make a test pass.**"),
    (
        "tolerance routing rule",
        "Tolerance changes route",
    ),
    (
        "validator rule",
        "validate_workbook_or_raise",
    ),
    ("parity contract reference", "model_parity_contract.md"),
    ("model change log reference", "model_change_log.md"),
    ("parity_constants reference", "parity_constants.py"),
]


@pytest.fixture(scope="module")
def canonical_text() -> str:
    return CANONICAL_AGENTS.read_text()


@pytest.fixture(scope="module")
def template_text() -> str:
    return KIT_TEMPLATE.read_text()


@pytest.mark.parametrize("label,fragment", REQUIRED_FRAGMENTS, ids=[r[0] for r in REQUIRED_FRAGMENTS])
def test_canonical_agents_contains_fragment(canonical_text: str, label: str, fragment: str) -> None:
    """Every fragment must already exist in annuity_model/AGENTS.md."""
    assert fragment in canonical_text, (
        f"Canonical AGENTS.md is missing the {label!r} fragment: {fragment!r}. "
        "If you changed AGENTS.md, also update REQUIRED_FRAGMENTS in this test "
        "and AGENTS_template.md to keep them in lock-step."
    )


@pytest.mark.parametrize("label,fragment", REQUIRED_FRAGMENTS, ids=[r[0] for r in REQUIRED_FRAGMENTS])
def test_kit_template_contains_fragment(template_text: str, label: str, fragment: str) -> None:
    """Every fragment must also exist in actuarial_parity_kit/AGENTS_template.md."""
    assert fragment in template_text, (
        f"actuarial_parity_kit/AGENTS_template.md is missing the {label!r} "
        f"fragment: {fragment!r}. The kit is the seed for every NEW product "
        "repo; if a non-negotiable rule is in the canonical AGENTS.md but not "
        "in the template, future repos start out non-compliant. Sync the "
        "template now."
    )


def test_kit_template_keeps_placeholder_marker(template_text: str) -> None:
    """The template must still look like a template (placeholders intact).

    A common failure mode would be someone copying annuity_model/AGENTS.md
    over the template wholesale -- losing the ``[PRODUCT NAME]`` etc.
    placeholders that make the kit usable for new products.
    """
    assert "[PRODUCT NAME]" in template_text, (
        "AGENTS_template.md no longer contains the [PRODUCT NAME] placeholder. "
        "If you intentionally renamed the placeholder, update this assertion."
    )


def test_kit_template_does_not_hardcode_spia(template_text: str) -> None:
    """Sanity: the template must not bake SPIA-specific filenames into the kit.

    Names like ``pricing_projection.py`` are SPIA-internal; they would make
    no sense in a future ULSG/DI/etc. repo seeded from this kit.
    """
    spia_specific = [
        "pricing_projection.py",
        "build_pricing_excel_workbook.py",
        "alm_excel_ladder.py",
        "rila_projection.py",
        "build_rila_excel_workbook.py",
    ]
    leaked = [name for name in spia_specific if name in template_text]
    assert not leaked, (
        f"AGENTS_template.md leaked SPIA-specific filenames: {leaked!r}. The "
        "template should reference placeholder names like [engine].py / "
        "[workbook_builder].py so it is reusable for new product lines."
    )

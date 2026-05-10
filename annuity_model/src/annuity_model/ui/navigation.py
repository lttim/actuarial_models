"""Shared Streamlit navigation metadata for the pricing demo UI."""

from __future__ import annotations

SECTION_LABELS: dict[str, str] = {
    "overview": "Overview",
    "run": "Pricing Run",
    "portfolio": "Portfolio (multi-policy)",
    "workbench": "Pricing Workbench",
    "alm": "ALM",
    "what_if": "What-if Analysis",
    "experience": "Experience Study",
    "excel_replicator": "Excel Replicator",
    "tests": "Unit Tests",
}

SECTION_ORDER: tuple[str, ...] = (
    "overview",
    "run",
    "workbench",
    "alm",
    "what_if",
    "experience",
    "excel_replicator",
    "tests",
)


def section_label(section: str) -> str:
    """Return the display label for a sidebar section key."""
    return SECTION_LABELS[section]


def section_options(*, include_portfolio: bool) -> list[str]:
    """Return sidebar section keys, injecting Portfolio after Pricing Run when enabled."""
    if include_portfolio:
        return [*SECTION_ORDER[:2], "portfolio", *SECTION_ORDER[2:]]
    return list(SECTION_ORDER)


def overview_section_labels() -> list[str]:
    """Return non-overview section labels for the Overview page."""
    return [SECTION_LABELS[key] for key in SECTION_ORDER if key != "overview"]

"""Small page router for the Streamlit pricing app."""

from __future__ import annotations

from collections.abc import Callable, Mapping

PageRenderer = Callable[[], None]


def render_selected_page(
    page: str,
    renderers: Mapping[str, PageRenderer],
    *,
    fallback: PageRenderer,
) -> None:
    """Render the selected page, falling back to the unit-test dashboard."""
    renderers.get(page, fallback)()

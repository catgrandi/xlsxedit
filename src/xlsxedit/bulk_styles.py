"""Style resolution for bulk row writes."""

from __future__ import annotations

from typing import Any

from xlsxedit.styles import Styles

_STYLE_KEYS = frozenset({
    "bold",
    "italic",
    "underline",
    "font_size",
    "font_color",
    "font_name",
    "bg_color",
    "horizontal_align",
    "vertical_align",
    "num_format",
})


def normalize_style_spec(spec: dict[str, Any] | None) -> dict[str, Any]:
    if not spec:
        return {}
    return {k: v for k, v in spec.items() if k in _STYLE_KEYS and v is not None}


def merge_style_specs(
    column_spec: dict[str, Any] | None,
    row_spec: dict[str, Any] | None,
) -> dict[str, Any]:
    merged = normalize_style_spec(column_spec)
    merged.update(normalize_style_spec(row_spec))
    return merged


def spec_cache_key(spec: dict[str, Any]) -> tuple[tuple[str, Any], ...]:
    return tuple(sorted(spec.items()))


class BulkStyleCache:
    """Allocate and cache cellXf indices for inline bulk style specs."""

    def __init__(self, styles: Styles, *, base_xf: int = 0):
        self._styles = styles
        self._base_xf = base_xf
        self._cache: dict[tuple[tuple[str, Any], ...], int] = {}

    def index_for(self, spec: dict[str, Any]) -> int | None:
        normalized = normalize_style_spec(spec)
        if not normalized:
            return None
        key = spec_cache_key(normalized)
        if key not in self._cache:
            self._cache[key] = self._styles.allocate_cell_style(
                base_xf=self._base_xf,
                **normalized,
            )
        return self._cache[key]


def validate_style_sources(
    *,
    template_rows: int | list[int] | None,
    row_styles: list[dict[str, Any]] | None,
) -> None:
    if template_rows is not None and row_styles is not None:
        raise ValueError("use either template_rows or row_styles, not both")


def normalize_template_rows(template_rows: int | list[int] | None) -> list[int]:
    if template_rows is None:
        return []
    if isinstance(template_rows, int):
        return [template_rows]
    return list(template_rows)

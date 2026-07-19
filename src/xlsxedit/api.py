"""Package-level helpers (default template path, etc.)."""

from __future__ import annotations

from pathlib import Path


def _default_xlsx_path() -> Path:
    """Return the path to the built-in default .xlsx package."""
    return Path(__file__).resolve().parent / "templates" / "default.xlsx"

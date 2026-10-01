"""Exception hierarchy for xlsxedit.

All library errors derive from :class:`XlsxeditError`, so callers can write a
single ``except XlsxeditError`` without catching unrelated builtins. Each
subclass also inherits the builtin it replaces (``ValueError``, ``KeyError``,
``RuntimeError``), so existing ``except ValueError`` / ``except KeyError``
handlers keep working.
"""

from __future__ import annotations


class XlsxeditError(Exception):
    """Base class for all xlsxedit errors."""


class WorksheetNotFoundError(XlsxeditError, KeyError):
    """Raised when a worksheet name is not present in the workbook."""


class DuplicateWorksheetError(XlsxeditError, ValueError):
    """Raised when adding or renaming to a worksheet name that already exists."""


class InvalidRangeError(XlsxeditError, ValueError):
    """Raised for a malformed cell address or range string (e.g. ``"A1:C3"``)."""


class InvalidColorError(XlsxeditError, ValueError):
    """Raised for an invalid color value."""


class InvalidImageError(XlsxeditError, ValueError):
    """Raised for unsupported or corrupt image data."""


class MissingPartError(XlsxeditError, RuntimeError):
    """Raised when a required package part is absent."""


class FormulaGroupError(XlsxeditError, ValueError):
    """Raised when an edit would split a shared, array, or data-table formula."""

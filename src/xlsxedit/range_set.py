"""Cell ranges and range lists as SpreadsheetML writes them (``ST_Ref``, ``ST_Sqref``).

``CellRange`` is one area: a cell (``B2``), a range (``A1:C3``), whole columns
(``A:C``) or whole rows (``1:3``), each bound optionally absolute (``$A$1``).
``SheetRangeSet`` is a whitespace-separated list of areas, the form of
``sqref`` attributes and ``xm:sqref`` elements.

Both shift on row or column insertion with Excel's rules: an area before the
insertion point stays, one at or after it moves, and one that straddles it
grows. An area that spans every row, written ``A:A`` or as Excel's
``A1:A1048576``, ignores row insertion, and likewise for columns. Results are
clamped to the grid, and an area pushed entirely past the last row or column
is dropped. An area whose bounds do not change keeps its original text.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Iterable, Iterator
from dataclasses import dataclass, field, replace

from xlsxedit.exceptions import InvalidRangeError
from xlsxedit.oxml.address import MAX_COL, MAX_ROW, col_to_index, index_to_col

CELL = "cell"
RANGE = "range"
COLUMNS = "columns"
ROWS = "rows"

_CELL_END = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)")
# xsd:list items are separated by XML whitespace only, not by other Unicode spaces.
_XML_SPACE = re.compile(r"[ \t\r\n]+")
_COL_END = re.compile(r"(\$?)([A-Za-z]{1,3})")
_ROW_END = re.compile(r"(\$?)([0-9]+)")


def column_number(letters: str) -> int:
    """1-based number of column ``letters`` (``A`` .. ``XFD``); ``InvalidRangeError`` otherwise."""
    if _COL_END.fullmatch(letters) is None or letters.startswith("$"):
        raise InvalidRangeError(f"invalid column: {letters!r}")
    return _col(letters, letters)


def _col(letters: str, source: str) -> int:
    index = col_to_index(letters) + 1
    if index > MAX_COL:
        raise InvalidRangeError(f"column {letters!r} is past XFD in {source!r}")
    return index


def _row(digits: str, source: str) -> int:
    row = int(digits)
    if not 1 <= row <= MAX_ROW:
        raise InvalidRangeError(f"row {row} is outside 1..{MAX_ROW} in {source!r}")
    return row


def _ordered(lo: int, lo_abs: bool, hi: int, hi_abs: bool) -> tuple[int, bool, int, bool]:
    return (lo, lo_abs, hi, hi_abs) if lo <= hi else (hi, hi_abs, lo, lo_abs)


@dataclass(frozen=True)
class CellRange:
    """One rectangular area of a worksheet, with 1-based inclusive bounds.

    ``kind`` records how the area is written (``CELL``, ``RANGE``, ``COLUMNS``
    or ``ROWS``) and ``absolute`` which bounds carry ``$``, in the order
    ``min_col, min_row, max_col, max_row``. ``text`` is the source text, kept
    while the bounds are unchanged.
    """

    min_col: int
    min_row: int
    max_col: int
    max_row: int
    kind: str = RANGE
    absolute: tuple[bool, bool, bool, bool] = (False, False, False, False)
    text: str | None = field(default=None, compare=False)

    @classmethod
    def parse(cls, text: str) -> CellRange:
        """Parse one ``ST_Ref``; raise ``InvalidRangeError`` if malformed or off the grid."""
        parts = text.split(":")
        if len(parts) == 1:
            m = _CELL_END.fullmatch(parts[0])
            if m is not None:
                col, row = _col(m.group(2), text), _row(m.group(4), text)
                flags = (bool(m.group(1)), bool(m.group(3)))
                return cls(col, row, col, row, CELL, flags + flags, text)
        elif len(parts) == 2:
            first, last = parts
            cells = _CELL_END.fullmatch(first), _CELL_END.fullmatch(last)
            if cells[0] and cells[1]:
                (a, b) = cells
                c1, ca1, c2, ca2 = _ordered(
                    _col(a.group(2), text), bool(a.group(1)), _col(b.group(2), text), bool(b.group(1))
                )
                r1, ra1, r2, ra2 = _ordered(
                    _row(a.group(4), text), bool(a.group(3)), _row(b.group(4), text), bool(b.group(3))
                )
                return cls(c1, r1, c2, r2, RANGE, (ca1, ra1, ca2, ra2), text)
            cols = _COL_END.fullmatch(first), _COL_END.fullmatch(last)
            if cols[0] and cols[1]:
                (a, b) = cols
                c1, ca1, c2, ca2 = _ordered(
                    _col(a.group(2), text), bool(a.group(1)), _col(b.group(2), text), bool(b.group(1))
                )
                return cls(c1, 1, c2, MAX_ROW, COLUMNS, (ca1, False, ca2, False), text)
            rows = _ROW_END.fullmatch(first), _ROW_END.fullmatch(last)
            if rows[0] and rows[1]:
                (a, b) = rows
                r1, ra1, r2, ra2 = _ordered(
                    _row(a.group(2), text), bool(a.group(1)), _row(b.group(2), text), bool(b.group(1))
                )
                return cls(1, r1, MAX_COL, r2, ROWS, (False, ra1, False, ra2), text)
        raise InvalidRangeError(f"invalid range: {text!r}")

    @classmethod
    def from_bounds(cls, min_col: int, min_row: int, max_col: int, max_row: int) -> CellRange:
        """A ``CELL`` or ``RANGE`` area over 1-based bounds inside the grid."""
        if not (1 <= min_col <= max_col <= MAX_COL and 1 <= min_row <= max_row <= MAX_ROW):
            raise InvalidRangeError(
                f"bounds ({min_col}, {min_row}, {max_col}, {max_row}) are outside the grid"
            )
        single = min_col == max_col and min_row == max_row
        return cls(min_col, min_row, max_col, max_row, CELL if single else RANGE)

    def __str__(self) -> str:
        return self.text if self.text is not None else self._render()

    def _render(self) -> str:
        ca1, ra1, ca2, ra2 = ("$" if flag else "" for flag in self.absolute)
        if self.kind == COLUMNS:
            return f"{ca1}{index_to_col(self.min_col - 1)}:{ca2}{index_to_col(self.max_col - 1)}"
        if self.kind == ROWS:
            return f"{ra1}{self.min_row}:{ra2}{self.max_row}"
        first = f"{ca1}{index_to_col(self.min_col - 1)}{ra1}{self.min_row}"
        if self.kind == CELL:
            return first
        return f"{first}:{ca2}{index_to_col(self.max_col - 1)}{ra2}{self.max_row}"

    @property
    def spans_all_rows(self) -> bool:
        return self.min_row == 1 and self.max_row == MAX_ROW

    @property
    def spans_all_cols(self) -> bool:
        return self.min_col == 1 and self.max_col == MAX_COL

    @property
    def is_cell(self) -> bool:
        """Whether the area covers exactly one cell, however it is written."""
        return self.min_col == self.max_col and self.min_row == self.max_row

    def contains(self, col: int, row: int) -> bool:
        return self.min_col <= col <= self.max_col and self.min_row <= row <= self.max_row

    def intersects(self, other: CellRange) -> bool:
        return not (
            self.max_col < other.min_col
            or other.max_col < self.min_col
            or self.max_row < other.min_row
            or other.max_row < self.min_row
        )

    def shift_rows(self, at: int, count: int) -> CellRange | None:
        """The area after inserting ``count`` rows before row ``at``.

        ``None`` when the whole area is pushed past the last row.
        """
        if self.spans_all_rows or self.max_row < at:
            return self
        lo = self.min_row + count if self.min_row >= at else self.min_row
        if lo > MAX_ROW:
            return None
        hi = min(self.max_row + count, MAX_ROW)
        if (lo, hi) == (self.min_row, self.max_row):
            return self
        return replace(self, min_row=lo, max_row=hi, text=None)

    def shift_cols(self, at: int, count: int) -> CellRange | None:
        """The area after inserting ``count`` columns before 1-based column ``at``.

        ``None`` when the whole area is pushed past column XFD.
        """
        if self.spans_all_cols or self.max_col < at:
            return self
        lo = self.min_col + count if self.min_col >= at else self.min_col
        if lo > MAX_COL:
            return None
        hi = min(self.max_col + count, MAX_COL)
        if (lo, hi) == (self.min_col, self.max_col):
            return self
        return replace(self, min_col=lo, max_col=hi, text=None)

    def extend_to_row(self, row: int) -> CellRange:
        """The area grown down to ``row`` (clamped to the grid); unchanged if it already reaches it."""
        row = min(row, MAX_ROW)
        if row <= self.max_row:
            return self
        kind = RANGE if self.kind == CELL else self.kind
        return replace(self, max_row=row, kind=kind, text=None)


class SheetRangeSet:
    """A whitespace-separated list of areas (``ST_Sqref``), in order.

    Shifting returns a new set (or this one, when nothing moved) whose text is
    the original text while no area changed. Sets compare by identity.
    """

    __slots__ = ("_ranges", "_text")

    def __init__(self, ranges: Iterable[CellRange] = (), *, text: str | None = None):
        self._ranges = tuple(ranges)
        self._text = text

    @classmethod
    def parse(cls, text: str) -> SheetRangeSet:
        """Parse an ``ST_Sqref``; raise ``InvalidRangeError`` on any malformed area."""
        tokens = _XML_SPACE.split(text.strip(" \t\r\n"))
        return cls((CellRange.parse(token) for token in tokens if token), text=text)

    def __str__(self) -> str:
        if self._text is not None:
            return self._text
        return " ".join(str(r) for r in self._ranges)

    def __repr__(self) -> str:
        return f"SheetRangeSet({str(self)!r})"

    def __iter__(self) -> Iterator[CellRange]:
        return iter(self._ranges)

    def __len__(self) -> int:
        return len(self._ranges)

    def __bool__(self) -> bool:
        return bool(self._ranges)

    @property
    def ranges(self) -> tuple[CellRange, ...]:
        return self._ranges

    def _map(self, fn: Callable[[CellRange], CellRange | None]) -> SheetRangeSet:
        moved = [fn(r) for r in self._ranges]
        if all(new is old for new, old in zip(moved, self._ranges)):
            return self
        return SheetRangeSet(r for r in moved if r is not None)

    def shift_rows(self, at: int, count: int) -> SheetRangeSet:
        """Each area after inserting ``count`` rows before row ``at``; areas pushed off are dropped."""
        return self._map(lambda r: r.shift_rows(at, count))

    def shift_cols(self, at: int, count: int) -> SheetRangeSet:
        """Each area after inserting ``count`` columns before 1-based column ``at``."""
        return self._map(lambda r: r.shift_cols(at, count))

    def extend_to_row(self, row: int) -> SheetRangeSet:
        return self._map(lambda r: r.extend_to_row(row))

    def intersects(self, other: CellRange | SheetRangeSet | str) -> bool:
        if isinstance(other, str):
            other = SheetRangeSet.parse(other)
        others = (other,) if isinstance(other, CellRange) else tuple(other)
        return any(a.intersects(b) for a in self._ranges for b in others)

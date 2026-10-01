"""Merged cell range helpers."""

from __future__ import annotations

import re

from xlsxedit.exceptions import InvalidRangeError
from xlsxedit.oxml.address import col_to_index, index_to_col, join_address, split_address
from xlsxedit.range_set import CellRange

_RANGE_RE = re.compile(
    r"^\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$"
)


def parse_range(ref: str) -> tuple[str, int, str, int]:
    """Parse ``A1:C3`` (``$`` allowed) into top-left and bottom-right coordinates."""
    m = _RANGE_RE.match(ref)
    if not m:
        raise InvalidRangeError(f"invalid range: {ref!r}")
    c1, r1 = m.group(1).upper(), int(m.group(2))
    if m.group(3) is None:
        return c1, r1, c1, r1
    return c1, r1, m.group(3).upper(), int(m.group(4))


def address_in_range(address: str, ref: str) -> bool:
    col, row = split_address(address)
    c1, r1, c2, r2 = parse_range(ref)
    ci = col_to_index(col)
    c1i, c2i = col_to_index(c1), col_to_index(c2)
    return c1i <= ci <= c2i and r1 <= row <= r2


def merge_anchor(ref: str) -> str:
    c1, r1, _, _ = parse_range(ref)
    return join_address(c1, r1)


def normalize_range(ref: str) -> str:
    """Normalize ``A1`` to ``A1:A1``."""
    c1, r1, c2, r2 = parse_range(ref)
    if c1 == c2 and r1 == r2:
        return join_address(c1, r1)
    return f"{join_address(c1, r1)}:{join_address(c2, r2)}"


def ranges_overlap(a: str, b: str) -> bool:
    """Whether two ``ST_Ref`` areas share a cell (whole columns and rows included)."""
    return CellRange.parse(a).intersects(CellRange.parse(b))


class MergeMap:
    """Map cell addresses to merge anchor (top-left) cells.

    Reads every ``ST_Ref`` form a ``mergeCell`` may use, including whole
    columns or rows. Merges of up to ``_INDEXED_CELLS`` cells are indexed cell
    by cell; larger ones are checked by containment.
    """

    _INDEXED_CELLS = 4096

    def __init__(self, refs: list[str]):
        self._refs = refs
        self._anchor_for: dict[str, str] = {}
        self._large: list[tuple[CellRange, str]] = []
        for ref in refs:
            area = CellRange.parse(ref)
            anchor = join_address(index_to_col(area.min_col - 1), area.min_row)
            size = (area.max_col - area.min_col + 1) * (area.max_row - area.min_row + 1)
            if size > self._INDEXED_CELLS:
                self._large.append((area, anchor))
                continue
            for ci in range(area.min_col - 1, area.max_col):
                for ri in range(area.min_row, area.max_row + 1):
                    self._anchor_for[join_address(index_to_col(ci), ri)] = anchor

    @property
    def refs(self) -> list[str]:
        return list(self._refs)

    def anchor_for(self, address: str) -> str | None:
        if not address:
            return None
        anchor = self._anchor_for.get(address.upper())
        if anchor is not None or not self._large:
            return anchor
        try:
            col, row = split_address(address)
        except InvalidRangeError:
            return None
        col_idx = col_to_index(col) + 1
        return next((a for area, a in self._large if area.contains(col_idx, row)), None)

    def is_anchor(self, address: str) -> bool:
        a = self.anchor_for(address)
        return a is not None and a.upper() == address.upper()

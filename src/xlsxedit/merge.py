"""Merged cell range helpers."""

from __future__ import annotations

import re

from xlsxedit.exceptions import InvalidRangeError
from xlsxedit.oxml.address import col_to_index, index_to_col, join_address, split_address

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
    ac1, ar1, ac2, ar2 = parse_range(a)
    bc1, br1, bc2, br2 = parse_range(b)
    a_c1, a_c2 = col_to_index(ac1), col_to_index(ac2)
    b_c1, b_c2 = col_to_index(bc1), col_to_index(bc2)
    if a_c2 < b_c1 or b_c2 < a_c1:
        return False
    if ar2 < br1 or br2 < ar1:
        return False
    return True


class MergeMap:
    """Map cell addresses to merge anchor (top-left) cells."""

    def __init__(self, refs: list[str]):
        self._refs = refs
        self._anchor_for: dict[str, str] = {}
        for ref in refs:
            anchor = merge_anchor(ref)
            c1, r1, c2, r2 = parse_range(ref)
            for ci in range(col_to_index(c1), col_to_index(c2) + 1):
                for ri in range(r1, r2 + 1):
                    addr = join_address(index_to_col(ci), ri)
                    self._anchor_for[addr] = anchor

    @property
    def refs(self) -> list[str]:
        return list(self._refs)

    def anchor_for(self, address: str) -> str | None:
        return self._anchor_for.get(address.upper() if address else address)

    def is_anchor(self, address: str) -> bool:
        a = self.anchor_for(address)
        return a is not None and a.upper() == address.upper()

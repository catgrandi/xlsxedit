"""Excel A1-style cell address helpers."""

from __future__ import annotations

import re

from xlsxedit.exceptions import InvalidRangeError

_ADDRESS_RE = re.compile(r"^([A-Za-z]+)(\d+)$")


def col_to_index(col: str) -> int:
    """Convert column letters to 0-based index (``A`` → 0, ``B`` → 1)."""
    index = 0
    for ch in col.upper():
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return index - 1


def index_to_col(index: int) -> str:
    """Convert 0-based column index to letters."""
    if index < 0:
        raise InvalidRangeError(f"column index must be >= 0, got {index}")
    result = []
    n = index + 1
    while n:
        n, rem = divmod(n - 1, 26)
        result.append(chr(rem + ord("A")))
    return "".join(reversed(result))


def split_address(address: str) -> tuple[str, int]:
    """Split ``B2`` into ``('B', 2)``."""
    m = _ADDRESS_RE.match(address)
    if not m:
        raise InvalidRangeError(f"invalid cell address: {address!r}")
    return m.group(1).upper(), int(m.group(2))


def join_address(col: str, row: int) -> str:
    """Join column letters and row number into ``B2``."""
    if row < 1:
        raise InvalidRangeError(f"row must be >= 1, got {row}")
    return f"{col.upper()}{row}"

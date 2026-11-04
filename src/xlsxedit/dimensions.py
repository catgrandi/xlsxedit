"""Column and row dimension helpers."""

from __future__ import annotations

from copy import deepcopy
from typing import TYPE_CHECKING

from lxml import etree

from xlsxedit.opc.constants import SML_NS
from xlsxedit.oxml.address import col_to_index

if TYPE_CHECKING:
    from xlsxedit.worksheet import Worksheet

_COLS = f"{{{SML_NS}}}cols"
_COL = f"{{{SML_NS}}}col"
_SHEET_FORMAT_PR = f"{{{SML_NS}}}sheetFormatPr"
_ROW = f"{{{SML_NS}}}row"
_DEFAULT_COL_WIDTH = 8.43


class ColumnDimension:
    """Column width in Excel character units."""

    def __init__(self, worksheet: Worksheet, column: str):
        self._worksheet = worksheet
        self._column = column.upper()

    @property
    def column(self) -> str:
        return self._column

    @property
    def width(self) -> float:
        col_idx = col_to_index(self._column) + 1
        root = self._worksheet._part.element
        cols = root.find(_COLS)
        if cols is not None:
            for col_elm in cols.findall(_COL):
                min_c = int(col_elm.get("min", "0"))
                max_c = int(col_elm.get("max", "0"))
                if min_c <= col_idx <= max_c:
                    return float(col_elm.get("width", _DEFAULT_COL_WIDTH))
        fmt = root.find(_SHEET_FORMAT_PR)
        if fmt is not None and fmt.get("baseColWidth"):
            return float(fmt.get("baseColWidth"))
        if fmt is not None and fmt.get("defaultColWidth"):
            return float(fmt.get("defaultColWidth"))
        return _DEFAULT_COL_WIDTH

    @width.setter
    def width(self, value: float) -> None:
        """Set the width as Excel sets a column's ``ColumnWidth``.

        A ``<col>`` covering more columns than this one is split around it: the
        columns before and after keep every attribute, and so does this one,
        with the new ``width``, ``customWidth="1"`` and no ``bestFit``. Excel
        also unhides a hidden column; ``hidden`` stays here, as nothing in this
        API could hide the column again. A new ``<col>`` goes in column order,
        and a width the column's ``<col>`` already has changes nothing.
        """
        if value <= 0:
            raise ValueError("column width must be positive")
        col_idx = col_to_index(self._column) + 1
        root = self._worksheet._part.element
        cols = root.find(_COLS)
        if cols is None:
            sheet_data = root.find(f"{{{SML_NS}}}sheetData")
            cols = etree.Element(_COLS)
            if sheet_data is not None:
                root.insert(list(root).index(sheet_data), cols)
            else:
                root.append(cols)

        covering = [c for c in cols.findall(_COL) if _col_span(c)[0] <= col_idx <= _col_span(c)[1]]
        if covering:
            col_elm = covering[0]
            same = (col_elm.get("width"), col_elm.get("customWidth")) == (str(value), "1")
            if same and len(covering) == 1:
                return
            for other in covering[1:]:  # overlapping <col>s: only the first keeps this column
                _split_around(other, col_idx)
                cols.remove(other)
            _split_around(col_elm, col_idx)
            col_elm.set("width", str(value))
            col_elm.set("customWidth", "1")
            col_elm.attrib.pop("bestFit", None)
        else:
            col_elm = etree.Element(_COL)
            col_elm.set("min", str(col_idx))
            col_elm.set("max", str(col_idx))
            col_elm.set("width", str(value))
            col_elm.set("customWidth", "1")
            following = next((c for c in cols.findall(_COL) if _col_span(c)[0] > col_idx), None)
            if following is not None:
                following.addprevious(col_elm)
            else:
                cols.append(col_elm)
        self._worksheet._column_styles = None  # the styled <col> ranges first writes read
        self._worksheet._part.mark_dirty()


def _col_span(col_elm: etree._Element) -> tuple[int, int]:
    return int(col_elm.get("min", "0")), int(col_elm.get("max", "0"))


def _split_around(col_elm: etree._Element, col_idx: int) -> None:
    """Narrow ``col_elm`` to column ``col_idx``, adding copies with every attribute
    for the columns it covered before and after it."""
    low, high = _col_span(col_elm)
    if low < col_idx:
        before = deepcopy(col_elm)
        before.set("max", str(col_idx - 1))
        col_elm.addprevious(before)
    if col_idx < high:
        after = deepcopy(col_elm)
        after.set("min", str(col_idx + 1))
        col_elm.addnext(after)
    col_elm.set("min", str(col_idx))
    col_elm.set("max", str(col_idx))


class RowDimension:
    """Row height in points."""

    def __init__(self, worksheet: Worksheet, row: int):
        if row < 1:
            raise ValueError("row must be >= 1")
        self._worksheet = worksheet
        self._row = row

    @property
    def row(self) -> int:
        return self._row

    @property
    def height(self) -> float | None:
        row_elm = self._worksheet._find_row_element(self._row)
        if row_elm is None:
            return None
        ht = row_elm.get("ht")
        return float(ht) if ht is not None else None

    @height.setter
    def height(self, value: float) -> None:
        if value <= 0:
            raise ValueError("row height must be positive")
        existed = self._worksheet._find_row_element(self._row) is not None
        row_elm = self._worksheet._ensure_row(self._row)
        if not existed or (row_elm.get("ht"), row_elm.get("customHeight")) != (str(value), "1"):
            row_elm.set("ht", str(value))
            row_elm.set("customHeight", "1")
            self._worksheet._part.mark_dirty()


class _DimensionMap:
    def __init__(self, worksheet: Worksheet, dim_cls):
        self._worksheet = worksheet
        self._dim_cls = dim_cls
        self._cache: dict = {}

    def __getitem__(self, key: str | int):
        if key in self._cache:
            return self._cache[key]
        dim = self._dim_cls(self._worksheet, key)
        self._cache[key] = dim
        return dim


class ColumnDimensions(_DimensionMap):
    def __init__(self, worksheet: Worksheet):
        super().__init__(worksheet, ColumnDimension)

    def __getitem__(self, column: str) -> ColumnDimension:
        col = column.upper()
        if not col.isalpha():
            raise KeyError(f"invalid column: {column!r}")
        return super().__getitem__(col)


class RowDimensions(_DimensionMap):
    def __init__(self, worksheet: Worksheet):
        super().__init__(worksheet, RowDimension)

    def __getitem__(self, row: int) -> RowDimension:
        if isinstance(row, str):
            row = int(row)
        if row < 1:
            raise KeyError(f"invalid row: {row!r}")
        return super().__getitem__(row)

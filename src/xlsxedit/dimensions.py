"""Column and row dimension helpers."""

from __future__ import annotations

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

        for col_elm in list(cols.findall(_COL)):
            min_c = int(col_elm.get("min", "0"))
            max_c = int(col_elm.get("max", "0"))
            if min_c <= col_idx <= max_c:
                if min_c == max_c == col_idx:
                    col_elm.set("width", str(value))
                    col_elm.set("customWidth", "1")
                    return
                cols.remove(col_elm)

        col_elm = etree.SubElement(cols, _COL)
        col_elm.set("min", str(col_idx))
        col_elm.set("max", str(col_idx))
        col_elm.set("width", str(value))
        col_elm.set("customWidth", "1")


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
        row_elm = self._worksheet._ensure_row(self._row)
        row_elm.set("ht", str(value))
        row_elm.set("customHeight", "1")


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

"""Excel tables: :class:`Table` proxies over the live tree of each table part.

``ws.tables`` hands out one proxy per table part, and every property reads the
part's tree, so a held table follows edits such as ``insert_rows``. Setters
and :meth:`Table.resize` check everything before they change the tree; a
refused edit raises and leaves the part as it was, and an edit that changes
nothing leaves its bytes as they were.
"""

from __future__ import annotations

import re
import uuid
from dataclasses import dataclass
from typing import TYPE_CHECKING, Callable, Iterator

from lxml import etree

from xlsxedit._ooxml_order import ATTRIBUTE_DEFAULTS, ORDER
from xlsxedit.exceptions import InvalidRangeError, TableError
from xlsxedit.merge import parse_range
from xlsxedit.opc.constants import CT, OFFICE_REL_NS, SML_NS
from xlsxedit.opc.part import XmlPart
from xlsxedit.oxml.address import (
    MAX_COL,
    MAX_ROW,
    col_to_index,
    index_to_col,
    join_address,
    split_address,
)
from xlsxedit.parts import TablePart
from xlsxedit.styles import _XML_ILLEGAL
from xlsxedit.worksheet_order import insert_ordered

if TYPE_CHECKING:
    from xlsxedit.workbook import Workbook
    from xlsxedit.worksheet import Worksheet

_XM_NS = "http://schemas.microsoft.com/office/excel/2006/main"

_AUTO_FILTER = f"{{{SML_NS}}}autoFilter"
_FILTER_COLUMN = f"{{{SML_NS}}}filterColumn"
_SORT_STATE = f"{{{SML_NS}}}sortState"
_SORT_CONDITION = f"{{{SML_NS}}}sortCondition"
_TABLE_COLUMNS = f"{{{SML_NS}}}tableColumns"
_TABLE_COLUMN = f"{{{SML_NS}}}tableColumn"
_TABLE_STYLE_INFO = f"{{{SML_NS}}}tableStyleInfo"
_CALCULATED_COLUMN_FORMULA = f"{{{SML_NS}}}calculatedColumnFormula"
_TOTALS_ROW_FORMULA = f"{{{SML_NS}}}totalsRowFormula"
_TABLE_PARTS = f"{{{SML_NS}}}tableParts"
_TABLE_PART = f"{{{SML_NS}}}tablePart"
_DEFINED_NAMES = f"{{{SML_NS}}}definedNames"
_DEFINED_NAME = f"{{{SML_NS}}}definedName"
_DXFS = f"{{{SML_NS}}}dxfs"
_DXF = f"{{{SML_NS}}}dxf"
_TABLE_STYLES = f"{{{SML_NS}}}tableStyles"
_TABLE_STYLE = f"{{{SML_NS}}}tableStyle"
_CELL = f"{{{SML_NS}}}c"
_FORMULA_TAGS = (
    *(f"{{{SML_NS}}}{local}" for local in ("f", "formula", "formula1", "formula2")),
    f"{{{_XM_NS}}}f",
)
_TABLE_FORMULA_TAGS = (_CALCULATED_COLUMN_FORMULA, _TOTALS_ROW_FORMULA)

_TABLE_DEFAULTS = ATTRIBUTE_DEFAULTS["CT_Table"]
_COLUMN_DEFAULTS = ATTRIBUTE_DEFAULTS["CT_TableColumn"]
_BOOLEANS = {"true": True, "1": True, "false": False, "0": False}
_DIGITS = re.compile(r"[0-9]+")
_BUILT_IN_STYLE = re.compile(
    r"TableStyle(?:Light(?:[1-9]|1[0-9]|2[01])|Medium(?:[1-9]|1[0-9]|2[0-8])|Dark(?:[1-9]|1[01]))"
)

# xr, xr2 and xr3: Excel's revision uids, which must not repeat across parts.
_REVISION_NAMESPACES = frozenset(
    {
        "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
        "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
        "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    }
)

# Table names follow the defined-name rules: a letter, underscore or
# backslash, then letters, digits, underscores and periods, and never
# something Excel would read as an A1 or R1C1 reference inside the grid.
_TABLE_NAME_RE = re.compile(r"(?:[^\W\d]|\\)[\w.]*")
_A1_NAME_RE = re.compile(r"([A-Za-z]{1,3})(\d+)")
_R1C1_NAME_RE = re.compile(r"[Rr](\d*)(?:[Cc](\d*))?|[Cc](\d*)")

Box = tuple[int, int, int, int]


@dataclass(frozen=True)
class TableStyle:
    """A table's ``tableStyleInfo``: the style and the parts of the table it marks.

    ``name`` is a built-in style such as ``"TableStyleMedium2"``, a custom
    table style of ``styles.xml``, or ``None`` for Excel's "None" style. The
    defaults are those of a new table in Excel.
    """

    name: str | None = "TableStyleMedium2"
    show_first_column: bool = False
    show_last_column: bool = False
    show_row_stripes: bool = True
    show_column_stripes: bool = False


_STYLE_FLAGS = (
    ("show_first_column", "showFirstColumn"),
    ("show_last_column", "showLastColumn"),
    ("show_row_stripes", "showRowStripes"),
    ("show_column_stripes", "showColumnStripes"),
)


class TableColumn:
    """One column of a table, read from its ``tableColumn`` element (read-only)."""

    def __init__(self, element: etree._Element, table: Table):
        self._element = element
        self._table = table

    def __repr__(self) -> str:
        return f"<TableColumn id={self._element.get('id')!r} name={self.name!r}>"

    @property
    def _owner(self) -> str:
        return f"column {self.name!r} of table {self._table.display_name!r}"

    @property
    def name(self) -> str | None:
        """The column name, which the header cell holds when the table has a header row."""
        return self._element.get("name")

    @property
    def id(self) -> int | None:
        return _uint(self._element, "id", None, self._owner)

    @property
    def unique_name(self) -> str | None:
        return self._element.get("uniqueName")

    @property
    def calculated_column_formula(self) -> str | None:
        return _child_text(self._element, _CALCULATED_COLUMN_FORMULA)

    @property
    def totals_row_function(self) -> str:
        """``"none"`` (the schema default when absent), ``"sum"``, …, or ``"custom"``."""
        return self._element.get("totalsRowFunction", _COLUMN_DEFAULTS["totalsRowFunction"])

    @property
    def totals_row_formula(self) -> str | None:
        return _child_text(self._element, _TOTALS_ROW_FORMULA)

    @property
    def totals_row_label(self) -> str | None:
        return self._element.get("totalsRowLabel")

    @property
    def data_dxf_id(self) -> int | None:
        return _uint(self._element, "dataDxfId", None, self._owner)

    @property
    def query_table_field_id(self) -> int | None:
        return _uint(self._element, "queryTableFieldId", None, self._owner)


def _dxf_property(attr: str) -> property:
    def getter(self: Table) -> int | None:
        return _uint(self._element, attr, None, self._owner)

    def setter(self: Table, value: int | None) -> None:
        self._set_dxf_id(attr, value)

    return property(
        getter,
        setter,
        doc=f"Index of the ``dxfs`` entry in ``styles.xml`` behind ``{attr}``, or ``None``.",
    )


class Table:
    """An Excel table on a worksheet, from ``ws.tables`` or ``ws.add_table(...)``.

    Attributes that the file leaves out read as their schema defaults
    (``header_row_count`` 1, ``totals_row_count`` 0, ``totals_row_shown``
    ``True``). ``insertRow``, ``insertRowShift`` and everything the API does
    not name are kept as they are.
    """

    def __init__(self, part: TablePart, worksheet: Worksheet):
        self._part = part
        self._worksheet = worksheet

    def __repr__(self) -> str:
        return f"<Table name={self.name!r} ref={self.ref!r}>"

    @property
    def _element(self) -> etree._Element:
        return self._part.element

    @property
    def _owner(self) -> str:
        return f"table {self.display_name!r}"

    # --- identity -----------------------------------------------------------

    @property
    def name(self) -> str | None:
        return self._element.get("name")

    @property
    def id(self) -> int | None:
        """The table's workbook-wide id."""
        return _uint(self._element, "id", None, self._owner)

    @property
    def display_name(self) -> str | None:
        """The name formulas use for the table.

        Setting it checks the name as ``add_table`` does and refuses, with
        :class:`TableError`, a rename while a formula, a defined name or
        another part of the workbook (a pivot cache, a connection, …) refers
        to the table by its current name: xlsxedit does not rewrite formulas.
        """
        return self._element.get("displayName")

    @display_name.setter
    def display_name(self, value: str) -> None:
        current = self.display_name
        if value == current:
            return
        problem = _table_name_problem(value)
        if problem is None:
            used, _ = _table_names_and_ids(self._worksheet._workbook, exclude=self._part)
            if value.casefold() in used:
                problem = "is already used by a table or defined name in this workbook"
        if problem is not None:
            raise TableError(f"table display_name {value!r} {problem}")
        if current and current.casefold() != value.casefold():
            self._check_name_unreferenced(current, value)
        self._element.set("displayName", value)
        self._part.mark_dirty()

    @property
    def ref(self) -> str | None:
        return self._element.get("ref")

    # --- columns ------------------------------------------------------------

    @property
    def columns(self) -> list[str]:
        """The column names, left to right (columns without a name are left out)."""
        return [c.get("name") for c in self._column_elements() if c.get("name")]

    @property
    def column_objects(self) -> list[TableColumn]:
        """A :class:`TableColumn` per ``tableColumn``, left to right."""
        return [TableColumn(c, self) for c in self._column_elements()]

    def _column_elements(self) -> list[etree._Element]:
        cols = self._element.find(_TABLE_COLUMNS)
        return [] if cols is None else cols.findall(_TABLE_COLUMN)

    # --- layout -------------------------------------------------------------

    @property
    def header_row_count(self) -> int:
        """1 when the table has a header row, else 0.

        Setting it keeps ``ref``: 0 turns the header row into the first data
        row and removes the ``autoFilter`` (Excel writes none without a header
        row); 1 makes the first row the header row and requires its cells to
        hold the column names already.
        """
        return _uint(
            self._element, "headerRowCount", _TABLE_DEFAULTS["headerRowCount"], self._owner
        )

    @header_row_count.setter
    def header_row_count(self, value: int) -> None:
        count = _check_row_count(value, "header_row_count")
        if count == self.header_row_count:
            return
        box = self._box()
        action = f"setting header_row_count of table {self.display_name!r} to {count}"
        self._check_data_rows(box, count, self.totals_row_count, action)
        if count:
            self._check_columns_fit(box)
            problems = header_mismatches(self._worksheet, box[1], box[0], self._column_names())
            if problems:
                raise TableError(
                    f"the first row of table {self.display_name!r} ({self.ref}) would become its "
                    f"header row but does not hold its column names: {'; '.join(problems)}; "
                    "write them there first"
                )
        elm = self._element
        _set_attribute(elm, "headerRowCount", str(count), _TABLE_DEFAULTS["headerRowCount"])
        if not count:
            af = elm.find(_AUTO_FILTER)
            if af is not None:
                elm.remove(af)
        self._part.mark_dirty()

    @property
    def totals_row_count(self) -> int:
        """1 when the table has a totals row, its last row, else 0.

        Setting it keeps ``ref`` and the cells: 1 turns the last data row into
        the totals row (and sets ``totals_row_shown``), 0 turns the totals row
        into a data row. The ``autoFilter`` and sort ranges follow.
        """
        return _uint(
            self._element, "totalsRowCount", _TABLE_DEFAULTS["totalsRowCount"], self._owner
        )

    @totals_row_count.setter
    def totals_row_count(self, value: int) -> None:
        count = _check_row_count(value, "totals_row_count")
        current = self.totals_row_count
        if count == current:
            return
        box = self._box()
        header = self.header_row_count
        action = f"setting totals_row_count of table {self.display_name!r} to {count}"
        self._check_data_rows(box, header, count, action)
        elm = self._element
        _set_attribute(elm, "totalsRowCount", str(count), _TABLE_DEFAULTS["totalsRowCount"])
        if count:
            _set_attribute(elm, "totalsRowShown", "1", "1")
        af = elm.find(_AUTO_FILTER)
        if af is not None:
            af.set("ref", _box_ref(_filter_box(box, count)))
        self._refit_sort_states(_data_box(box, header, current), _data_box(box, header, count))
        self._part.mark_dirty()

    @property
    def totals_row_shown(self) -> bool:
        """Whether the totals row has been shown, which Excel uses to restore it."""
        return _bool(
            self._element, "totalsRowShown", _TABLE_DEFAULTS["totalsRowShown"], self._owner
        )

    @totals_row_shown.setter
    def totals_row_shown(self, value: bool) -> None:
        if not isinstance(value, bool):
            raise TypeError(f"totals_row_shown must be a bool, got {type(value)!r}")
        if value == self.totals_row_shown:
            return
        if not value and self.totals_row_count:
            raise TableError(
                f"table {self.display_name!r} shows its totals row; set totals_row_count to 0 first"
            )
        _set_attribute(self._element, "totalsRowShown", "1" if value else "0", "1")
        self._part.mark_dirty()

    # --- formatting ---------------------------------------------------------

    header_row_dxf_id = _dxf_property("headerRowDxfId")
    data_dxf_id = _dxf_property("dataDxfId")
    totals_row_dxf_id = _dxf_property("totalsRowDxfId")
    header_row_border_dxf_id = _dxf_property("headerRowBorderDxfId")
    table_border_dxf_id = _dxf_property("tableBorderDxfId")
    totals_row_border_dxf_id = _dxf_property("totalsRowBorderDxfId")

    @property
    def header_row_cell_style(self) -> str | None:
        return self._element.get("headerRowCellStyle")

    @property
    def data_cell_style(self) -> str | None:
        return self._element.get("dataCellStyle")

    @property
    def totals_row_cell_style(self) -> str | None:
        return self._element.get("totalsRowCellStyle")

    @property
    def style(self) -> TableStyle | None:
        """The ``tableStyleInfo`` as a :class:`TableStyle`, or ``None`` without one.

        Setting ``None`` removes the element. A style name must be one of
        Excel's built-in table styles or a custom table style in ``styles.xml``.
        """
        info = self._element.find(_TABLE_STYLE_INFO)
        if info is None:
            return None
        owner = f"tableStyleInfo of table {self.display_name!r}"
        flags = {field: _bool(info, attr, "0", owner) for field, attr in _STYLE_FLAGS}
        return TableStyle(name=info.get("name"), **flags)

    @style.setter
    def style(self, value: TableStyle | None) -> None:
        if value is not None and not isinstance(value, TableStyle):
            raise TypeError(f"style must be a TableStyle or None, got {type(value)!r}")
        if value is not None:
            self._check_style(value)
        if value == self.style:
            return
        elm = self._element
        info = elm.find(_TABLE_STYLE_INFO)
        if value is None:
            elm.remove(info)
            self._part.mark_dirty()
            return
        new = info is None
        if new:
            info = etree.Element(_TABLE_STYLE_INFO)
            insert_ordered(elm, info, ORDER["CT_Table"])
        _set_attribute(info, "name", value.name, None)
        owner = f"tableStyleInfo of table {self.display_name!r}"
        for field, attr in _STYLE_FLAGS:
            wanted = getattr(value, field)
            if new or _bool(info, attr, "0", owner) != wanted:
                info.set(attr, "1" if wanted else "0")
        self._part.mark_dirty()

    # --- ranges -------------------------------------------------------------

    @property
    def header_range(self) -> str | None:
        """The header row as a range, or ``None`` without one."""
        box = self._box()
        header = min(self.header_row_count, box[3] - box[1] + 1)
        if not header:
            return None
        return _box_ref((box[0], box[1], box[2], box[1] + header - 1))

    @property
    def data_range(self) -> str | None:
        """The data rows, between the header and totals rows, as a range."""
        box = self._box()
        data = _data_box(box, self.header_row_count, self.totals_row_count)
        return _box_ref(data) if data[1] <= data[3] else None

    @property
    def totals_range(self) -> str | None:
        """The totals row as a range, or ``None`` without one."""
        box = self._box()
        totals = min(self.totals_row_count, box[3] - box[1] + 1)
        if not totals:
            return None
        return _box_ref((box[0], box[3] - totals + 1, box[2], box[3]))

    @property
    def first_data_row(self) -> int | None:
        box = self._box()
        data = _data_box(box, self.header_row_count, self.totals_row_count)
        return data[1] if data[1] <= data[3] else None

    @property
    def last_data_row(self) -> int | None:
        box = self._box()
        data = _data_box(box, self.header_row_count, self.totals_row_count)
        return data[3] if data[1] <= data[3] else None

    # --- resize -------------------------------------------------------------

    def resize(self, ref: str, *, columns: list[str] | None = None) -> None:
        """Make ``ref`` the table's range, as Excel's Resize Table does.

        Cells are neither moved nor written, except the header cells of new
        columns (see below): the range takes in whatever the cells hold. With
        a totals row, the last row of ``ref`` is the totals row. The
        ``autoFilter`` covers ``ref`` without the totals row (and is removed
        when the table has no header row), and the sort state is moved and
        clamped into the new data rows. A ``ref`` equal to the current range
        changes nothing.

        ``columns`` names every column of ``ref``, left to right; it is
        required when ``ref`` adds or drops columns. Columns the table keeps
        keep their ``tableColumn`` (id, formulas, totals function) and must
        keep their names. A new column's header cell must hold its name or be
        empty, in which case the name is written there.

        Raises :class:`TableError`, changing nothing, when the header row
        would move, ``ref`` does not overlap the current range, would leave no
        data row, or overlaps another table on the sheet, when columns change
        without ``columns``, or when a dropped column is named in a structured
        reference; :class:`~xlsxedit.exceptions.InvalidRangeError` for a
        ``ref`` that is not a range on the grid.
        """
        apply = self._plan_resize(ref, columns)
        if apply is not None:
            apply()

    def _plan_resize(
        self, ref: str, columns: list[str] | None = None, *, move_header: bool = False
    ) -> Callable[[], None] | None:
        """Check a resize to ``ref`` and return what applies it, or ``None`` when
        it would change nothing. Raises before anything changes.

        ``move_header`` lets the header row move: ``write_dataframe`` checks
        that the new header row holds the column names.
        """
        if not isinstance(ref, str):
            raise TypeError(f"ref must be a range string such as 'A1:C10', got {type(ref)!r}")
        new = _range_box(ref)
        new_ref = _box_ref(new)
        old = self._box()
        if columns is None and new == old:
            return None
        name = self.display_name
        action = f"resizing table {name!r} to {new_ref}"
        header, totals = self.header_row_count, self.totals_row_count
        elements = self._check_columns_fit(old)
        current = [c.get("name") for c in elements]
        if columns is None:
            if (new[0], new[2]) != (old[0], old[2]):
                raise TableError(
                    f"{action} would change its columns ({self.ref}); pass columns= with "
                    "the name of every column of the new range"
                )
            names = current
        else:
            names = _check_column_names(columns, width=new[2] - new[0] + 1)
        if new == old and names == current:
            return None
        if header and new[1] != old[1] and not move_header:
            raise TableError(
                f"{action} would move its header row from row {old[1]} to row {new[1]}; "
                "the header row stays where it is"
            )
        if not _boxes_overlap(new, old):
            raise TableError(f"{action} would not overlap its current range {self.ref}")
        self._check_data_rows(new, header, totals, action)
        for other in table_parts_for_worksheet(self._worksheet):
            other_box = _try_box(other.ref)
            if other._part is not self._part and other_box and _boxes_overlap(other_box, new):
                raise TableError(f"{action} would overlap table {other.name!r} ({other.ref})")

        slots: list[etree._Element | str] = []
        header_writes: dict[str, str] = {}
        for offset, col_name in enumerate(names):
            col = new[0] + offset
            if old[0] <= col <= old[2]:
                element = elements[col - old[0]]
                if element.get("name") != col_name:
                    raise TableError(
                        f"{action}: column {index_to_col(col)} keeps its tableColumn, which is "
                        f"named {element.get('name')!r}, not {col_name!r}; resize does not "
                        "rename columns"
                    )
                slots.append(element)
                continue
            slots.append(col_name)
            if header:
                address = join_address(index_to_col(col), new[1])
                held = _header_cell(self._worksheet, address)
                if held is None:
                    header_writes[address] = col_name
                elif held != col_name:
                    raise TableError(
                        f"{action}: the header cell {address} of new column {col_name!r} "
                        f"{_describe(held)}; write the name there or clear it"
                    )
        kept = {id(s) for s in slots if not isinstance(s, str)}
        dropped = [e for e in elements if id(e) not in kept]
        self._check_columns_unreferenced(dropped, action)
        if header_writes and self._worksheet.shared_strings is None:
            raise ValueError("workbook has no shared strings part")

        def apply() -> None:
            self._apply_resize(old, new, slots, dropped, header_writes)

        return apply

    def _apply_resize(
        self,
        old: Box,
        new: Box,
        slots: list[etree._Element | str],
        dropped: list[etree._Element],
        header_writes: dict[str, str],
    ) -> None:
        from xlsxedit.cell import Cell

        elm = self._element
        header, totals = self.header_row_count, self.totals_row_count
        elm.set("ref", _box_ref(new))
        af = elm.find(_AUTO_FILTER)
        if af is not None and not header:
            elm.remove(af)
        elif af is not None:
            af_box = _try_box(af.get("ref")) or old
            af.set("ref", _box_ref(_filter_box(new, totals)))
            positions = {id(s): i for i, s in enumerate(slots) if not isinstance(s, str)}
            columns = self._column_elements()
            for fc in af.findall(_FILTER_COLUMN):
                col_id = fc.get("colId", "")
                if not _DIGITS.fullmatch(col_id):
                    continue
                index = af_box[0] + int(col_id) - old[0]
                element = columns[index] if 0 <= index < len(columns) else None
                position = positions.get(id(element))
                if position is None:
                    af.remove(fc)
                elif position != int(col_id):
                    fc.set("colId", str(position))
        self._refit_sort_states(_data_box(old, header, totals), _data_box(new, header, totals))

        cols_elm = elm.find(_TABLE_COLUMNS)
        for element in dropped:
            cols_elm.remove(element)
        ids = [
            int(c.get("id")) for c in self._column_elements() if _DIGITS.fullmatch(c.get("id", ""))
        ]
        ids += [int(c.get("id")) for c in dropped if _DIGITS.fullmatch(c.get("id", ""))]
        next_id = max(ids, default=0) + 1
        for index, slot in enumerate(slots):
            if not isinstance(slot, str):
                continue
            column = etree.Element(_TABLE_COLUMN)
            column.set("id", str(next_id))
            column.set("name", slot)
            next_id += 1
            following = next((s for s in slots[index + 1 :] if not isinstance(s, str)), None)
            if following is not None:
                following.addprevious(column)
            else:
                cols_elm.append(column)
        cols_elm.set("count", str(len(slots)))

        ws = self._worksheet
        for address, value in header_writes.items():
            row = ws._ensure_row(split_address(address)[1])
            Cell(ws._bulk_get_or_create_cell(address, row), ws).value = value
        self._part.mark_dirty()

    # --- row and column inserts --------------------------------------------

    def _move(self, move: Callable[[str], str]) -> None:
        """Apply an insert's ``move`` (a range before the insert -> after it) to
        the table, its ``autoFilter`` and its sort state.

        An ``autoFilter`` that covered the table without its totals row still
        does; a sort range that ended at the last data row still does, so an
        insert at the totals row grows them like the data rows.
        """
        elm = self._element
        ref = elm.get("ref")
        if not ref:
            return
        old = _range_box(ref)
        new_ref = move(ref)
        new = _range_box(new_ref)
        totals = self.totals_row_count
        changed = _put_ref(elm, new_ref)
        af = elm.find(_AUTO_FILTER)
        if af is not None and af.get("ref"):
            if af.get("ref") == _box_ref(_filter_box(old, totals)):
                changed |= _put_ref(af, _box_ref(_filter_box(new, totals)))
            else:
                changed |= _put_ref(af, move(af.get("ref")))
        old_last, new_last = old[3] - totals, new[3] - totals
        for sort_state in self._sort_states():
            for item in (sort_state, *sort_state.iterchildren(_SORT_CONDITION)):
                item_ref = item.get("ref")
                if not item_ref:
                    continue
                moved = move(item_ref)
                box, moved_box = _try_box(item_ref), _try_box(moved)
                if box and moved_box and box[3] == old_last and moved_box[3] != new_last:
                    moved = _box_ref((*moved_box[:3], new_last))
                changed |= _put_ref(item, moved)
        if changed:
            self._part.mark_dirty()

    def _check_move(self, check: Callable[[str, str], None]) -> None:
        """Run ``check(ref, what)`` on every range of the table an insert moves."""
        name = self.display_name or self.name or str(self._part.partname)
        what = f"table {name!r}"
        ref = self._element.get("ref")
        if ref:
            check(ref, what)
        af = self._element.find(_AUTO_FILTER)
        if af is not None and af.get("ref"):
            check(af.get("ref"), what)
        for sort_state in self._sort_states():
            for item in (sort_state, *sort_state.iterchildren(_SORT_CONDITION)):
                if item.get("ref"):
                    check(item.get("ref"), f"the sort state of {what}")

    # --- helpers ------------------------------------------------------------

    def _box(self) -> Box:
        ref = self.ref
        try:
            return _range_box(ref or "")
        except InvalidRangeError:
            raise TableError(f"table {self.display_name!r} has a malformed ref {ref!r}") from None

    def _column_names(self) -> list[str | None]:
        return [c.get("name") for c in self._column_elements()]

    def _check_columns_fit(self, box: Box) -> list[etree._Element]:
        """The ``tableColumn`` elements, or ``TableError`` when they do not fit ``box``."""
        elements = self._column_elements()
        cols = self._element.find(_TABLE_COLUMNS)
        count = None if cols is None else cols.get("count")
        width = box[2] - box[0] + 1
        if len(elements) != width or (count is not None and count != str(len(elements))):
            raise TableError(
                f"table {self.display_name!r} has {len(elements)} tableColumn elements "
                f"(count={count}) but its ref {self.ref} is {width} columns wide"
            )
        return elements

    def _check_data_rows(self, box: Box, header: int, totals: int, action: str) -> None:
        if box[3] - box[1] + 1 - header - totals < 1:
            raise TableError(f"{action} would leave it without a data row")

    def _sort_states(self) -> list[etree._Element]:
        elm = self._element
        af = elm.find(_AUTO_FILTER)
        found = [elm.find(_SORT_STATE), None if af is None else af.find(_SORT_STATE)]
        return [s for s in found if s is not None]

    def _refit_sort_states(self, old: Box, new: Box) -> None:
        """Move the sort ranges from data rows ``old`` to ``new``: an edge on the
        old data edge follows it, any other is clamped; a sort condition whose
        column leaves the table is dropped, and a sort state left empty too."""
        for sort_state in self._sort_states():
            parent = sort_state.getparent()
            box = _try_box(sort_state.get("ref"))
            if box is None:
                continue
            fitted = _refit(box, old, new, follow_columns=True)
            if fitted is None:
                parent.remove(sort_state)
                continue
            _put_ref(sort_state, _box_ref(fitted))
            conditions = list(sort_state.iterchildren(_SORT_CONDITION))
            for condition in conditions:
                condition_box = _try_box(condition.get("ref"))
                if condition_box is None:
                    continue
                fitted = _refit(condition_box, old, new, follow_columns=False)
                if fitted is None:
                    sort_state.remove(condition)
                else:
                    _put_ref(condition, _box_ref(fitted))
            if conditions and sort_state.find(_SORT_CONDITION) is None:
                parent.remove(sort_state)

    def _set_dxf_id(self, attr: str, value: int | None) -> None:
        if value is not None:
            if isinstance(value, bool) or not isinstance(value, int):
                raise TypeError(f"a dxf id must be an int or None, got {type(value)!r}")
            count = _dxf_count(self._worksheet._workbook)
            if not 0 <= value < count:
                raise TableError(
                    f"dxf id {value} is not in styles.xml, which has {count} dxf entries"
                )
        if value == getattr(self, _DXF_FIELDS[attr]):
            return
        _set_attribute(self._element, attr, None if value is None else str(value), None)
        self._part.mark_dirty()

    def _check_style(self, style: TableStyle) -> None:
        if style.name is not None:
            if not isinstance(style.name, str):
                raise TypeError(f"TableStyle.name must be a str or None, got {type(style.name)!r}")
            if not (
                _BUILT_IN_STYLE.fullmatch(style.name)
                or style.name in _custom_table_styles(self._worksheet._workbook)
            ):
                raise TableError(
                    f"table style {style.name!r} is neither a built-in table style nor a "
                    "table style in styles.xml"
                )
        for field, _attr in _STYLE_FLAGS:
            if not isinstance(getattr(style, field), bool):
                raise TypeError(f"TableStyle.{field} must be a bool")

    def _check_name_unreferenced(self, current: str, new: str) -> None:
        workbook = self._worksheet._workbook
        for where, formula in _formulas_mentioning(workbook, current):
            raise TableError(
                f"renaming table {current!r} to {new!r} would break {where}, whose formula "
                f"{formula!r} refers to it by name; xlsxedit does not rewrite formulas"
            )
        quoted = re.compile(b'"' + re.escape(current.encode()) + b'"', re.IGNORECASE)
        for part in workbook._package.iter_parts():
            if isinstance(part, XmlPart) or not part.content_type.endswith("+xml"):
                continue
            if quoted.search(part.blob):
                raise TableError(
                    f"renaming table {current!r} to {new!r} would break {part.partname}, "
                    "which names it"
                )

    def _check_columns_unreferenced(self, dropped: list[etree._Element], action: str) -> None:
        """Refuse dropping a column that a structured reference names."""
        names = [c.get("name") for c in dropped if c.get("name")]
        if not names:
            return
        patterns = {
            name: re.compile(r"\[@?" + re.escape(_structured_name(name)) + r"\]", re.IGNORECASE)
            for name in names
        }
        formulas = [
            (f"table {self.display_name!r}", elm.text)
            for column in self._column_elements()
            if column not in dropped
            for elm in column.iterchildren(*_TABLE_FORMULA_TAGS)
            if elm.text
        ]
        if self.display_name:
            formulas = [
                *formulas,
                *_formulas_mentioning(
                    self._worksheet._workbook, self.display_name, skip=self._part
                ),
            ]
        for where, formula in formulas:
            for name, pattern in patterns.items():
                if pattern.search(formula):
                    raise TableError(
                        f"{action} would drop column {name!r}, which {where} names in "
                        f"{formula!r}; xlsxedit does not rewrite formulas"
                    )


_DXF_FIELDS = {
    "headerRowDxfId": "header_row_dxf_id",
    "dataDxfId": "data_dxf_id",
    "totalsRowDxfId": "totals_row_dxf_id",
    "headerRowBorderDxfId": "header_row_border_dxf_id",
    "tableBorderDxfId": "table_border_dxf_id",
    "totalsRowBorderDxfId": "totals_row_border_dxf_id",
}


# --- worksheet tables ---------------------------------------------------------


def table_parts_for_worksheet(worksheet: Worksheet) -> list[Table]:
    """The tables of ``worksheet`` in ``tableParts`` order, one proxy per part."""
    tables: list[Table] = []
    block = worksheet._part.element.find(_TABLE_PARTS)
    if block is None:
        return tables
    proxies = worksheet._table_proxies
    for rel_elm in block.iterchildren(_TABLE_PART):
        r_id = rel_elm.get(f"{{{OFFICE_REL_NS}}}id")
        if r_id is None:
            continue
        try:
            part = worksheet._part.rels[r_id].target_part
        except KeyError:
            continue
        if not isinstance(part, TablePart):
            continue
        table = proxies.get(part)
        if table is None:
            table = proxies[part] = Table(part, worksheet)
        tables.append(table)
    return tables


def header_mismatches(ws: Worksheet, row: int, first_col: int, names: list) -> list[str]:
    """Describe each header cell from ``first_col`` on that does not hold its name as text."""
    problems = []
    for offset, name in enumerate(names):
        address = join_address(index_to_col(first_col + offset), row)
        held = _header_cell(ws, address)
        if held != name:
            problems.append(f"{address} {_describe(held)}, expected {name!r}")
    return problems


def _header_cell(ws: Worksheet, address: str):
    """What the cell at ``address`` holds as a header: its value, ``None`` when it
    is empty, or ``_FORMULA`` when it holds a formula."""
    from xlsxedit.cell import Cell

    c_elm = ws._find_cell_element(address)
    if c_elm is None:
        return None
    cell = Cell(c_elm, ws)
    return _FORMULA if cell.has_formula else cell.value


class _Formula:
    def __repr__(self) -> str:
        return "a formula"


_FORMULA = _Formula()


def _describe(value) -> str:
    if value is None:
        return "is empty"
    if value is _FORMULA:
        return "holds a formula"
    return f"holds {value!r}"


# --- names, ids and revision uids ----------------------------------------------


def _check_column_names(columns, *, width: int | None = None) -> list[str]:
    """``columns`` as a list of valid, distinct column names, else ``TableError``."""
    if isinstance(columns, str):
        raise TypeError("columns must be a list of column names, not a str")
    names = list(columns)
    if not names:
        raise TableError("columns must not be empty")
    if width is not None and len(names) != width:
        raise TableError(f"{len(names)} column names were given for a range {width} columns wide")
    seen: set[str] = set()
    for name in names:
        if not isinstance(name, str) or not name:
            raise TableError(f"column names must be non-empty strings, got {name!r}")
        if _XML_ILLEGAL.search(name):
            raise TableError(f"column name {name!r} holds characters XML cannot store")
        if name.casefold() in seen:
            raise TableError(f"duplicate column name {name!r} (names ignore case)")
        seen.add(name.casefold())
    return names


def _new_revision_uid() -> str:
    """A fresh revision uid in Excel's ``{GUID}`` form."""
    return "{" + str(uuid.uuid4()).upper() + "}"


def _refresh_revision_uids(root: etree._Element) -> None:
    """Give every revision ``uid`` attribute in ``root`` a fresh GUID.

    Elements that shared a uid still share one: Excel writes the same uid on
    a ``<table>`` and its ``<autoFilter>``.
    """
    fresh: dict[str, str] = {}
    for elm in root.iter(etree.Element):
        for attr, value in list(elm.attrib.items()):
            qname = etree.QName(attr)
            if qname.localname == "uid" and qname.namespace in _REVISION_NAMESPACES:
                elm.set(attr, fresh.setdefault(value.upper(), _new_revision_uid()))


def _table_name_problem(name) -> str | None:
    """Why ``name`` cannot name a table, or ``None`` when it can."""
    if not isinstance(name, str) or not name:
        return "must be a non-empty string"
    if len(name) > 255:
        return "is longer than 255 characters"
    if not _TABLE_NAME_RE.fullmatch(name):
        return (
            "must start with a letter, underscore or backslash and continue with "
            "letters, digits, underscores or periods"
        )
    if _is_cell_reference(name):
        return "looks like a cell reference"
    return None


def _is_cell_reference(name: str) -> bool:
    """Whether Excel reads ``name`` as an A1 or R1C1 reference inside the grid."""
    a1 = _A1_NAME_RE.fullmatch(name)
    if a1:
        return col_to_index(a1[1]) < MAX_COL and 1 <= int(a1[2]) <= MAX_ROW
    r1c1 = _R1C1_NAME_RE.fullmatch(name)
    if r1c1 is None:
        return False
    row, col, lone_col = r1c1.groups()
    if row is None:
        col = lone_col
    return (not row or 1 <= int(row) <= MAX_ROW) and (not col or 1 <= int(col) <= MAX_COL)


def _table_names_and_ids(
    workbook: Workbook, *, exclude: object | None = None
) -> tuple[set[str], set[int]]:
    """Names in use (casefolded) and table ids across the workbook.

    Names are every table's ``name`` and ``displayName`` plus the defined
    names, which share their namespace. Every table part in the package is
    read, so the tables of a sheet still being copied count too.
    """
    names: set[str] = set()
    ids: set[int] = set()
    for part in workbook._package.get_parts_of_type(CT.TABLE):
        if part is exclude or not isinstance(part, XmlPart):
            continue
        elm = part.element
        names.update(v.casefold() for v in (elm.get("name"), elm.get("displayName")) if v)
        table_id = elm.get("id", "")
        if _DIGITS.fullmatch(table_id):
            ids.add(int(table_id))
    defined = workbook._workbook_part.element.find(_DEFINED_NAMES)
    if defined is not None:
        names.update(
            elm.get("name").casefold()
            for elm in defined.iterchildren(_DEFINED_NAME)
            if elm.get("name")
        )
    return names, ids


def _next_table_name(used: set[str], start: int) -> str:
    """The first ``TableN`` with ``N >= start`` not in ``used`` (casefolded names)."""
    n = max(start, 1)
    while f"table{n}" in used:
        n += 1
    return f"Table{n}"


# --- ranges ---------------------------------------------------------------------


def _range_box(ref: str) -> Box:
    """``(first column index, first row, last column index, last row)`` of an A1 range.

    Raises ``InvalidRangeError`` unless the range is ordered and inside the grid.
    """
    c1, r1, c2, r2 = parse_range(ref)
    box = (col_to_index(c1), r1, col_to_index(c2), r2)
    if not (0 <= box[0] <= box[2] < MAX_COL and 1 <= box[1] <= box[3] <= MAX_ROW):
        raise InvalidRangeError(f"range {ref!r} is reversed or outside A1:XFD1048576")
    return box


def _try_box(ref: str | None) -> Box | None:
    """The box of ``ref``, or ``None`` when it is missing or malformed."""
    try:
        return _range_box(ref or "")
    except InvalidRangeError:
        return None


def _box_ref(box: Box) -> str:
    c1, r1, c2, r2 = box
    return f"{join_address(index_to_col(c1), r1)}:{join_address(index_to_col(c2), r2)}"


def _boxes_overlap(a: Box, b: Box) -> bool:
    return a[0] <= b[2] and b[0] <= a[2] and a[1] <= b[3] and b[1] <= a[3]


def _filter_box(box: Box, totals: int) -> Box:
    """Where a table's ``autoFilter`` goes: the table without its totals rows."""
    return (box[0], box[1], box[2], box[3] - totals)


def _data_box(box: Box, header: int, totals: int) -> Box:
    """A table's data rows; empty (first row past last) when it has none."""
    return (box[0], box[1] + header, box[2], box[3] - totals)


def _refit(box: Box, old: Box, new: Box, *, follow_columns: bool) -> Box | None:
    """``box`` carried from area ``old`` to area ``new``: a row edge (and, with
    ``follow_columns``, a column edge) on the edge of ``old`` follows it, any
    other edge is clamped into ``new``. ``None`` when nothing is left."""
    r1 = new[1] if box[1] == old[1] else max(box[1], new[1])
    r2 = new[3] if box[3] == old[3] else min(box[3], new[3])
    if follow_columns:
        c1 = new[0] if box[0] == old[0] else max(box[0], new[0])
        c2 = new[2] if box[2] == old[2] else min(box[2], new[2])
    else:
        c1, c2 = max(box[0], new[0]), min(box[2], new[2])
    return (c1, r1, c2, r2) if c1 <= c2 and r1 <= r2 else None


def _put_ref(elm: etree._Element, ref: str) -> bool:
    """Set ``elm``'s ``ref``; return whether it changed."""
    if elm.get("ref") == ref:
        return False
    elm.set("ref", ref)
    return True


# --- attributes -------------------------------------------------------------------


def _uint(elm: etree._Element, attr: str, default: str | None, owner: str) -> int | None:
    text = elm.get(attr, default)
    if text is None:
        return None
    if not _DIGITS.fullmatch(text):
        raise TableError(f"{owner} has a malformed {attr} {text!r}")
    return int(text)


def _bool(elm: etree._Element, attr: str, default: str, owner: str) -> bool:
    text = elm.get(attr, default)
    if text not in _BOOLEANS:
        raise TableError(f"{owner} has a malformed {attr} {text!r}")
    return _BOOLEANS[text]


def _set_attribute(elm: etree._Element, attr: str, value: str | None, default: str | None) -> None:
    """Set ``attr``, or remove it when ``value`` is its schema default, as Excel writes it.

    Defaults are compared as booleans when both are ``"0"``/``"1"``/``"true"``/``"false"``.
    """
    if (
        value is None
        or value == default
        or (value in _BOOLEANS and default in _BOOLEANS and _BOOLEANS[value] == _BOOLEANS[default])
    ):
        elm.attrib.pop(attr, None)
    else:
        elm.set(attr, value)


def _check_row_count(value, field: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise TypeError(f"{field} must be 0 or 1, got {type(value)!r}")
    if value not in (0, 1):
        raise TableError(f"{field} must be 0 or 1, got {value}")
    return value


def _child_text(elm: etree._Element, tag: str) -> str | None:
    child = elm.find(tag)
    return None if child is None else child.text or ""


def _dxf_count(workbook: Workbook) -> int:
    root = workbook.styles._element
    dxfs = None if root is None else root.find(_DXFS)
    return 0 if dxfs is None else len(dxfs.findall(_DXF))


def _custom_table_styles(workbook: Workbook) -> set[str]:
    root = workbook.styles._element
    styles = None if root is None else root.find(_TABLE_STYLES)
    if styles is None:
        return set()
    return {s.get("name") for s in styles.iterchildren(_TABLE_STYLE) if s.get("name")}


# --- formulas ----------------------------------------------------------------------


def _structured_name(column: str) -> str:
    """``column`` as a structured reference writes it between brackets."""
    return re.sub(r"(['\[\]#])", r"'\1", column)


def _formulas_mentioning(
    workbook: Workbook, name: str, *, skip: TablePart | None = None
) -> Iterator[tuple[str, str]]:
    """``(where, formula)`` for every formula that names ``name`` as a whole word:
    in cells, conditional formats and validations of every sheet, defined
    names, and the column formulas of every table part except ``skip``."""
    word = re.compile(rf"(?<![\w.\\]){re.escape(name)}(?![\w.])", re.IGNORECASE)
    raw = re.compile(re.escape(name.encode()), re.IGNORECASE) if name.isascii() else None
    for ws in workbook.worksheets:
        blob = ws._part.unparsed_blob
        if blob is not None and raw is not None and raw.search(blob) is None:
            continue
        for elm in ws._part.element.iter(*_FORMULA_TAGS):
            if elm.text and word.search(elm.text):
                parent = elm.getparent()
                cell = parent.get("r") if parent is not None and parent.tag == _CELL else None
                where = f"cell {cell} of sheet {ws.name!r}" if cell else f"sheet {ws.name!r}"
                yield where, elm.text
    defined = workbook._workbook_part.element.find(_DEFINED_NAMES)
    if defined is not None:
        for elm in defined.iterchildren(_DEFINED_NAME):
            if elm.text and word.search(elm.text):
                yield f"defined name {elm.get('name')!r}", elm.text
    for part in workbook._package.get_parts_of_type(CT.TABLE):
        if part is skip or not isinstance(part, XmlPart):
            continue
        for elm in part.element.iter(*_TABLE_FORMULA_TAGS):
            if elm.text and word.search(elm.text):
                yield f"table {part.element.get('displayName')!r}", elm.text

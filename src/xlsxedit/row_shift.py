"""Shift a worksheet's cell references when rows or columns are inserted.

``plan_insert_rows`` and ``plan_insert_columns`` compute every change one
insertion makes before applying any of it; ``docs/features.md`` ("Inserting
rows and columns") lists what moves and when an insertion is refused. Ranges
move by the rules of :class:`xlsxedit.range_set.CellRange`, and references
in defined names by ``rewrite_formula_refs``. Planning raises for anything it cannot do, so a
refused insertion changes nothing; ``apply`` only assigns computed values.

The ``check_*`` functions validate the rest of an insert up front. The
remaining ``shift_*`` functions are the older per-structure entry points.
"""

from __future__ import annotations

import math
import operator
import re
from collections.abc import Callable
from dataclasses import dataclass
from datetime import date
from functools import partial

from lxml import etree

from xlsxedit.bulk_styles import normalize_style_spec
from xlsxedit.exceptions import GridOverflowError, InvalidColorError, InvalidRangeError
from xlsxedit.hyperlinks import _release_rel
from xlsxedit.opc.constants import OFFICE_REL_NS, SML_NS
from xlsxedit.oxml.address import MAX_COL, MAX_ROW, col_to_index, index_to_col
from xlsxedit.oxml.parser import parse_xml, serialize_xml
from xlsxedit.range_set import CellRange, SheetRangeSet
from xlsxedit.styles import datetime_to_serial, normalize_rgb

_ROW = f"{{{SML_NS}}}row"
_C = f"{{{SML_NS}}}c"
_COLS = f"{{{SML_NS}}}cols"
_COL = f"{{{SML_NS}}}col"
_MERGE_CELLS = f"{{{SML_NS}}}mergeCells"
_MERGE_CELL = f"{{{SML_NS}}}mergeCell"
_CF = f"{{{SML_NS}}}conditionalFormatting"
_TABLE_PARTS = f"{{{SML_NS}}}tableParts"
_TABLE_PART = f"{{{SML_NS}}}tablePart"
_AUTO_FILTER = f"{{{SML_NS}}}autoFilter"
_HYPERLINKS = f"{{{SML_NS}}}hyperlinks"
_HYPERLINK = f"{{{SML_NS}}}hyperlink"
_DATA_VALIDATIONS = f"{{{SML_NS}}}dataValidations"
_DATA_VALIDATION = f"{{{SML_NS}}}dataValidation"
_DEFINED_NAMES = f"{{{SML_NS}}}definedNames"
_DEFINED_NAME = f"{{{SML_NS}}}definedName"
_F = f"{{{SML_NS}}}f"
_SHEET_DATA = f"{{{SML_NS}}}sheetData"
_SHEETS = f"{{{SML_NS}}}sheets"
_SHEET = f"{{{SML_NS}}}sheet"
_EXT_LST = f"{{{SML_NS}}}extLst"

_X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
_X14_SPARKLINES = f"{{{_X14_NS}}}sparklines"
_EXT = f"{{{SML_NS}}}ext"
# containers whose schema needs at least one child
_LIST_CONTAINERS = frozenset(
    f"{{{SML_NS}}}{name}"
    for name in (
        "mergeCells",
        "hyperlinks",
        "dataValidations",
        "cols",
        "protectedRanges",
        "ignoredErrors",
        "cellWatches",
        "smartTags",
        "cellSmartTags",
        "rowBreaks",
        "colBreaks",
    )
)
_CONTENT = frozenset({f"{{{SML_NS}}}v", _F, f"{{{SML_NS}}}is"})
_RANGED_FORMULAS = frozenset({"shared", "array", "dataTable"})
_WRITABLE = (str, int, float, date)  # bool is an int and datetime a date
_XML_ILLEGAL = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")

_CELL_ADDRESS_RE = re.compile(r"([A-Za-z]{1,3})([0-9]+)")

# A reference in formula text: a cell, a range, whole columns or whole rows.
_REF = (
    r"(?:\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?"
    r"|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}"
    r"|\$?[0-9]+:\$?[0-9]+)"
)
# A sheet prefix: quoted, or an unquoted name, or an unquoted 3D span.
_SHEET_NAME = r"(?:'(?:[^']|'')+'|[^\W\d][\w.]*(?::[^\W\d][\w.]*)?)"
# A reference ends where no name (which may hold ? and \), call, structured
# reference or sheet prefix continues.
_REF_END = r"(?![\w.(\[!$?\\])"
_FORMULA_TOKEN_RE = re.compile(
    rf"""
    (?P<string>"(?:[^"]|"")*")
  | (?P<bracket>\[(?:'.|[^\]'])*\])  # ' escapes a bracket inside a structured reference
  | (?P<sheet>{_SHEET_NAME})!(?P<qref>{_REF}{_REF_END}|\#REF!)
  | '(?:[^']|'')*'  # a quoted prefix to a name, never read inside
  | (?P<ref>{_REF}){_REF_END}
  | [\w.$?\\]+
  | .
    """,
    re.VERBOSE | re.DOTALL,
)
def shift_row_number(row: int, at_row: int, delta: int) -> int:
    if row >= at_row:
        return row + delta
    return row


def shift_col_index(col_idx: int, at_col: int, delta: int) -> int:
    """Shift a 0-based column index when inserting at ``at_col``."""
    if col_idx >= at_col:
        return col_idx + delta
    return col_idx


@dataclass(frozen=True)
class _Insert:
    """``count`` rows inserted before row ``at``, or columns before 1-based column ``at``."""

    rows: bool
    at: int
    count: int

    @property
    def limit(self) -> int:
        return MAX_ROW if self.rows else MAX_COL

    def move(self, index: int) -> int:
        """A 1-based row or column number on the shifted axis; may pass the grid."""
        return index + self.count if index >= self.at else index

    def area(self, area: CellRange) -> CellRange | None:
        if self.rows:
            return area.shift_rows(self.at, self.count)
        return area.shift_cols(self.at, self.count)

    def ranges(self, ranges: SheetRangeSet) -> SheetRangeSet:
        if self.rows:
            return ranges.shift_rows(self.at, self.count)
        return ranges.shift_cols(self.at, self.count)

    def clamps(self, area: CellRange) -> bool:
        """Whether moving ``area`` pushes any of it past the grid."""
        if self.rows:
            full, hi = area.spans_all_rows, area.max_row
        else:
            full, hi = area.spans_all_cols, area.max_col
        return not full and hi >= self.at and hi + self.count > self.limit

    def overflow(self, what: str) -> GridOverflowError:
        if self.rows:
            where, end = f"row {self.at}", f"row {MAX_ROW}"
        else:
            where, end = f"column {index_to_col(self.at - 1)}", "column XFD"
        noun = "row" if self.rows else "column"
        return GridOverflowError(
            f"inserting {self.count} {noun}(s) at {where} would push {what} past {end}"
        )


def _area(text: str, where: str) -> CellRange:
    try:
        return CellRange.parse(text)
    except InvalidRangeError as exc:
        raise InvalidRangeError(f"{where}: {exc}") from None


def _range_set(text: str, where: str) -> SheetRangeSet:
    try:
        return SheetRangeSet.parse(text)
    except InvalidRangeError as exc:
        raise InvalidRangeError(f"{where}: {exc}") from None


def _number(text: str | None, where: str) -> int:
    if text is None or not text.isdigit():
        raise InvalidRangeError(f"{where}: invalid number {text!r}")
    return int(text)


def _cell_column(address: str, where: str) -> int:
    """1-based column of a ``<c r>`` address."""
    m = _CELL_ADDRESS_RE.fullmatch(address)
    index = col_to_index(m.group(1)) + 1 if m else 0
    if not 1 <= index <= MAX_COL:
        raise InvalidRangeError(f"{where}: invalid cell address {address!r}")
    return index


def _holds_content(c: etree._Element) -> bool:
    return any(child.tag in _CONTENT for child in c)


def _requires_children(container: etree._Element) -> bool:
    """Whether ``container`` is invalid, or meaningless, once its last child goes."""
    if container.tag in _LIST_CONTAINERS:
        return True
    return etree.QName(container).namespace == _X14_NS or container.tag in (_EXT, _EXT_LST)


def _detach(elm: etree._Element) -> None:
    """Remove ``elm``, and the list container it leaves empty; refresh the container's ``count``."""
    parent = elm.getparent()
    if parent is None:
        return
    if parent.tag == _X14_SPARKLINES and len(parent) == 1 and parent.getparent() is not None:
        _detach(parent.getparent())  # a sparkline group needs at least one sparkline
        return
    parent.remove(elm)
    empty = not any(isinstance(c.tag, str) for c in parent)
    if empty and parent.getparent() is not None and _requires_children(parent):
        _detach(parent)
    elif parent.get("count") is not None:
        parent.set("count", str(sum(1 for c in parent if c.tag == elm.tag)))


def _walk_sheet_data(sheet_data: etree._Element | None, ins: _Insert, *, apply: bool) -> None:
    """Check (``apply=False``) or carry out (``apply=True``) the move of every row and cell.

    Checking raises for an unparseable address and for content pushed past the
    grid, so carrying out the same walk afterwards cannot fail. Rows and cells
    without ``r`` get the number they implied.
    """
    if sheet_data is None:
        return
    for f in sheet_data.iter(_F):
        if f.get("t") in _RANGED_FORMULAS:
            _move_formula_ranges(f, ins, apply=apply, cell=f.getparent().get("r") or "")
    number = 0
    for row in list(sheet_data.iterchildren(_ROW)):
        r = row.get("r")
        number = _number(r, "row/@r") if r is not None else number + 1
        if not ins.rows:
            _move_cells_right(row, number, ins, apply=apply)
        elif number >= ins.at or r is None:  # rows above the insertion stay
            _move_row_down(sheet_data, row, number, ins, apply=apply)


def _move_row_down(
    sheet_data: etree._Element, row: etree._Element, number: int, ins: _Insert, *, apply: bool
) -> None:
    new_number = ins.move(number)
    if new_number > MAX_ROW:
        blocker = next((c for c in row.iterchildren(_C) if _holds_content(c)), None)
        if blocker is not None:
            raise ins.overflow(f"cell {blocker.get('r') or f'in row {number}'}")
        if apply:
            sheet_data.remove(row)
        return
    if apply:
        row.set("r", str(new_number))
    suffix = str(new_number)
    letters = ""
    for c in row.iterchildren(_C):
        address = c.get("r")
        if address is None:  # implied: the column after the previous cell's
            letters = index_to_col(col_to_index(letters) + 1) if letters else "A"
        else:
            m = _CELL_ADDRESS_RE.fullmatch(address)
            letters = m.group(1).upper() if m else ""
            if not letters or (len(letters) == 3 and letters > "XFD"):
                raise InvalidRangeError(f"c/@r: invalid cell address {address!r}")
        if apply:
            c.set("r", letters + suffix)


def _move_cells_right(row: etree._Element, number: int, ins: _Insert, *, apply: bool) -> None:
    col = 0
    for c in list(row.iterchildren(_C)):
        address = c.get("r")
        col = _cell_column(address, "c/@r") if address is not None else col + 1
        new_col = ins.move(col)
        if new_col > MAX_COL:
            if _holds_content(c):
                raise ins.overflow(f"cell {address or index_to_col(col - 1) + str(number)}")
            if apply:
                row.remove(c)
            continue
        if apply and (address is None or new_col != col):
            c.set("r", f"{index_to_col(new_col - 1)}{number}")


def _move_formula_ranges(f: etree._Element, ins: _Insert, *, apply: bool, cell: str) -> None:
    kind = f.get("t")
    for attr in ("ref", "r1", "r2") if kind == "dataTable" else ("ref",):
        value = f.get(attr)
        if not value:
            continue
        area = _area(value, f"{cell} <f t={kind!r}> @{attr}")
        moved = ins.area(area)
        if moved is None:
            raise ins.overflow(f"the {kind} formula range {attr}={value!r} of {cell}")
        if apply and moved is not area:
            f.set(attr, str(moved))


def _same_sheet(prefix: str, sheet_name: str | None) -> bool:
    """Whether a reference prefix names exactly ``sheet_name`` (not a 3D span or another book)."""
    if sheet_name is None:
        return False
    if prefix.startswith("'"):
        name = prefix[1:-1].replace("''", "'")
        if ":" in name or "[" in name:
            return False
    elif ":" in prefix:
        return False
    else:
        name = prefix
    return name.casefold() == sheet_name.casefold()


def _move_ref_text(ref: str, move: Callable[[CellRange], CellRange | None]) -> str:
    try:
        area = CellRange.parse(ref)
    except InvalidRangeError:
        return ref  # past the grid, so a name rather than a cell
    moved = move(area)
    return "#REF!" if moved is None else str(moved)


def rewrite_formula_refs(
    text: str,
    sheet_name: str,
    move: Callable[[CellRange], CellRange | None],
    *,
    unqualified: bool,
) -> str:
    """Move the references to ``sheet_name`` in formula ``text``.

    Rewrites references qualified with that sheet, and, when ``unqualified``,
    references without a sheet prefix. A reference is a whole token: never
    part of a sheet name (``Data2024!``), a function name (``LOG10(``,
    ``DAYS360(``), a defined name or a structured reference. String literals,
    other sheets, 3D spans and external workbooks are left alone. A reference
    pushed off the grid becomes ``#REF!``.
    """
    out: list[str] = []
    after_bracket = False
    for m in _FORMULA_TOKEN_RE.finditer(text):
        token = m.group(0)
        qref, ref = m.group("qref"), m.group("ref")
        if qref is not None and qref != "#REF!" and not after_bracket:
            if _same_sheet(m.group("sheet"), sheet_name):
                token = f"{m.group('sheet')}!{_move_ref_text(qref, move)}"
        elif ref is not None and unqualified and text[m.start() - 1 : m.start()] not in ("!", "]"):
            token = _move_ref_text(ref, move)
        out.append(token)
        after_bracket = m.group("bracket") is not None
    return "".join(out)


def _local_sheet_index(wb_elm: etree._Element, sheet_name: str) -> int | None:
    """Position of ``sheet_name`` among ``<sheets>``, the index ``localSheetId`` uses."""
    for index, sheet in enumerate(wb_elm.iterfind(f"{_SHEETS}/{_SHEET}")):
        if sheet.get("name") == sheet_name:
            return index
    return None


class _Planner:
    """Collects the edits for one insertion; ``apply`` performs them in order.

    Attribute and text assignments run first, then removals, then the
    ``after`` steps that act on other parts.
    """

    def __init__(self, ins: _Insert, *, sheet_name: str | None = None, part=None):
        self.ins = ins
        self.sheet_name = sheet_name
        self.part = part
        self.merges: list[CellRange] = []
        self._sets: list[Callable[[], None]] = []
        self._removals: list[Callable[[], None]] = []
        self._after: list[Callable[[], None]] = []

    def apply(self) -> None:
        for step in (*self._sets, *self._removals, *self._after):
            step()

    def set(self, elm: etree._Element, attr: str, value: str) -> None:
        if elm.get(attr) != value:
            self._sets.append(partial(elm.set, attr, value))

    def set_text(self, elm: etree._Element, value: str) -> None:
        if elm.text != value:
            self._sets.append(partial(setattr, elm, "text", value))

    def remove(self, elm: etree._Element) -> None:
        self._removals.append(partial(_detach, elm))

    def after(self, step: Callable[[], None]) -> None:
        self._after.append(step)

    # -- generic reference carriers -------------------------------------

    def area_attr(self, elm: etree._Element, attr: str, where: str, *, lost: str) -> None:
        """Move an ``ST_Ref`` attribute. When it is pushed off the grid, ``lost`` says
        what happens: ``"remove"`` the element or ``"raise"``."""
        text = elm.get(attr)
        if not text:
            return
        area = _area(text, where)
        moved = self.ins.area(area)
        if moved is None:
            if lost == "raise":
                raise self.ins.overflow(f"{where} {text!r}")
            self.remove(elm)
        elif moved is not area:
            self.set(elm, attr, str(moved))

    def sqref_attr(self, elm: etree._Element, attr: str, where: str) -> None:
        """Move an ``ST_Sqref`` attribute; remove ``elm`` when every area is pushed off."""
        text = elm.get(attr)
        if text is None:
            return
        ranges = _range_set(text, where)
        moved = self.ins.ranges(ranges)
        if moved is ranges:
            return
        if moved:
            self.set(elm, attr, str(moved))
        else:
            self.remove(elm)

    # -- worksheet structures -------------------------------------------

    def worksheet(self, root: etree._Element) -> None:
        """Every reference carrier inside the worksheet part other than ``sheetData``."""
        self.merge_cells(root)
        for cf in root.iterchildren(_CF):
            self.sqref_attr(cf, "sqref", "conditionalFormatting/@sqref")
        for dv in root.iterfind(f"{_DATA_VALIDATIONS}/{_DATA_VALIDATION}"):
            self.sqref_attr(dv, "sqref", "dataValidation/@sqref")
        self.hyperlinks(root)
        if not self.ins.rows:
            self.columns(root)

    def merge_cells(self, root: etree._Element) -> None:
        block = root.find(_MERGE_CELLS)
        if block is None:
            return
        for merge in block.iterchildren(_MERGE_CELL):
            ref = merge.get("ref")
            if not ref:
                continue
            area = _area(ref, "mergeCell/@ref")
            moved = self.ins.area(area)
            if moved is None or moved.is_cell:
                self.remove(merge)  # pushed off, or clamped to one cell
                continue
            self.merges.append(moved)
            if moved is not area:
                self.set(merge, "ref", str(moved))

    def hyperlinks(self, root: etree._Element) -> None:
        for link in root.iterfind(f"{_HYPERLINKS}/{_HYPERLINK}"):
            ref = link.get("ref")
            if not ref:
                continue
            area = _area(ref, "hyperlink/@ref")
            moved = self.ins.area(area)
            if moved is None:
                r_id = link.get(f"{{{OFFICE_REL_NS}}}id")
                if r_id is not None and self.part is not None:
                    self._removals.append(partial(_release_rel, self.part, r_id, keep=link))
                self.remove(link)
            elif moved is not area:
                self.set(link, "ref", str(moved))

    def columns(self, root: etree._Element) -> None:
        """Move ``<col min max>`` spans, clamped to column XFD."""
        for col in root.iterfind(f"{_COLS}/{_COL}"):
            lo = _number(col.get("min"), "col/@min")
            hi = _number(col.get("max"), "col/@max")
            if hi < self.ins.at:
                continue
            new_lo = self.ins.move(lo)
            if new_lo > MAX_COL:
                self.remove(col)
                continue
            self.set(col, "min", str(new_lo))
            self.set(col, "max", str(min(hi + self.ins.count, MAX_COL)))

    def defined_names(self, workbook) -> None:
        wb_elm = workbook._workbook_part.element
        block = wb_elm.find(_DEFINED_NAMES)
        if block is None:
            return
        index = _local_sheet_index(wb_elm, self.sheet_name)
        for elm in block.iterchildren(_DEFINED_NAME):
            if not elm.text:
                continue
            local_id = elm.get("localSheetId")
            local = index is not None and local_id is not None and local_id == str(index)
            new = rewrite_formula_refs(elm.text, self.sheet_name, self.ins.area, unqualified=local)
            self.set_text(elm, new)


class InsertPlan:
    """Every change that inserting rows or columns into one worksheet makes.

    Building the plan reads the worksheet, its table parts and the
    workbook's defined names, and raises
    ``InvalidRangeError`` or ``GridOverflowError`` before anything changes.
    ``apply`` then only assigns precomputed values. ``merges`` lists the
    sheet's merge ranges as they will be after ``apply``.
    """

    def __init__(self, worksheet, ins: _Insert):
        self._worksheet = worksheet
        self._insert = ins
        root = worksheet._part.element
        _walk_sheet_data(root.find(_SHEET_DATA), ins, apply=False)
        planner = _Planner(ins, sheet_name=worksheet.name, part=worksheet._part)
        planner.worksheet(root)
        workbook = worksheet._workbook
        self._check_tables()
        if ins.rows:
            planner.after(partial(shift_table_parts, worksheet, ins.at, ins.count))
        else:
            planner.after(partial(shift_table_parts_cols, worksheet, ins.at - 1, ins.count))
        planner.defined_names(workbook)
        planner.after(workbook._invalidate_calc_chain)
        planner.after(workbook.set_full_calc_on_load)
        self._planner = planner
        self.merges = planner.merges

    def _check_tables(self) -> None:
        """Refuse an insertion that would push any part of a table off the grid:
        its columns and rows cannot shrink with it."""
        part = self._worksheet._part
        block = part.element.find(_TABLE_PARTS)
        if block is None:
            return
        for table_part in block.iterchildren(_TABLE_PART):
            try:
                target = part.rels[table_part.get(f"{{{OFFICE_REL_NS}}}id")].target_part
            except KeyError:
                continue
            table = parse_xml(target.blob)
            name = table.get("displayName") or table.get("name") or str(target.partname)
            for elm in (table, table.find(_AUTO_FILTER)):
                ref = None if elm is None else elm.get("ref")
                if ref and self._insert.clamps(_area(ref, f"table {name!r}")):
                    raise self._insert.overflow(f"table {name!r} ({ref})")

    def apply(self) -> None:
        _walk_sheet_data(self._worksheet._part.element.find(_SHEET_DATA), self._insert, apply=True)
        self._planner.apply()


def plan_insert_rows(worksheet, at_row: int, count: int) -> InsertPlan:
    """Plan inserting ``count`` rows before row ``at_row`` (1-based)."""
    return InsertPlan(worksheet, _Insert(True, at_row, count))


def plan_insert_columns(worksheet, at_col: int, count: int) -> InsertPlan:
    """Plan inserting ``count`` columns before 0-based column ``at_col``."""
    return InsertPlan(worksheet, _Insert(False, at_col + 1, count))


def check_block_in_grid(
    method: str, first_row: int, first_col_idx: int, rows: int, cols: int
) -> None:
    """Refuse ``rows`` x ``cols`` new cells from ``first_row`` / 0-based ``first_col_idx``
    that start or end off the grid."""
    first_row, first_col_idx = operator.index(first_row), operator.index(first_col_idx)
    if first_row < 1 or first_col_idx < 0:
        raise InvalidRangeError(
            f"{method}: row {first_row}, column index {first_col_idx} is not on the grid"
        )
    if first_row + rows - 1 > MAX_ROW:
        raise GridOverflowError(f"{method} would write past row {MAX_ROW}")
    if first_col_idx + cols > MAX_COL:
        raise GridOverflowError(f"{method} would write past column XFD")


def check_values_writable(
    method: str,
    rows: list,
    *,
    start_col_idx: int,
    string_columns: set[int] | None,
    has_shared_strings: bool,
) -> None:
    """Refuse, before anything changes, a value the bulk writer would fail on.

    Mirrors ``bulk.write_cell_value``: the accepted types, the shared-strings
    part every string needs, and the characters XML cannot hold.
    """
    string_columns = string_columns or set()
    for row in rows:
        values = row if isinstance(row, (list, tuple)) else (row,)
        for offset, value in enumerate(values):
            if value is None:
                continue
            forced = start_col_idx + offset in string_columns
            if not forced and not isinstance(value, _WRITABLE):
                raise TypeError(
                    f"{method} accepts str, int, float, bool, datetime, date, or None; "
                    f"got {type(value)!r}"
                )
            if not forced and isinstance(value, date):
                try:
                    serial = datetime_to_serial(value)
                except (TypeError, ValueError, OverflowError) as exc:
                    raise TypeError(f"{method} cannot write {value!r}: {exc}") from None
                if not math.isfinite(serial):  # pandas.NaT
                    raise ValueError(f"{method} cannot write {value!r}; use None")
            if forced or isinstance(value, str):
                if not has_shared_strings:
                    raise ValueError("workbook has no shared strings part")
                if _XML_ILLEGAL.search(str(value)):
                    raise ValueError(f"{method}: {str(value)!r} holds characters XML cannot store")


def check_style_specs(*spec_lists: list[dict] | None) -> None:
    """Refuse, before anything changes, an inline style the bulk writer would fail on:
    a colour ``normalize_rgb`` rejects or a ``font_size`` that is not a number."""
    for specs in spec_lists:
        for spec in specs or ():
            normalized = normalize_style_spec(spec)
            for key in ("font_color", "bg_color"):
                color = normalized.get(key)
                if color is not None:
                    if not isinstance(color, str):
                        raise InvalidColorError(f"invalid color: {color!r}")
                    normalize_rgb(color)
            if "font_size" in normalized:
                int(normalized["font_size"])


def shift_range_ref(ref: str, at_row: int, delta: int) -> str:
    """Shift a range at or below ``at_row``; a straddling range grows. Clamped to the grid."""
    moved = CellRange.parse(ref).shift_rows(at_row, delta)
    if moved is None:
        raise _Insert(True, at_row, delta).overflow(repr(ref))
    return str(moved)


def shift_range_ref_cols(ref: str, at_col: int, delta: int) -> str:
    """Shift column letters in a range at or to the right of ``at_col`` (0-based)."""
    moved = CellRange.parse(ref).shift_cols(at_col + 1, delta)
    if moved is None:
        raise _Insert(False, at_col + 1, delta).overflow(repr(ref))
    return str(moved)


def shift_sqref(sqref: str, at_row: int, delta: int) -> str:
    return str(SheetRangeSet.parse(sqref).shift_rows(at_row, delta))


def shift_sqref_cols(sqref: str, at_col: int, delta: int) -> str:
    return str(SheetRangeSet.parse(sqref).shift_cols(at_col + 1, delta))


def _shift_worksheet_element(ws_element: etree._Element, ins: _Insert) -> None:
    planner = _Planner(ins)
    planner.worksheet(ws_element)
    planner.apply()


def shift_sheet_row_references(ws_element: etree._Element, at_row: int, delta: int) -> None:
    """Shift every reference carrier outside ``sheetData`` after row insertion."""
    _shift_worksheet_element(ws_element, _Insert(True, at_row, delta))


def shift_sheet_col_references(ws_element: etree._Element, at_col: int, delta: int) -> None:
    """Shift every reference carrier outside ``sheetData`` after column insertion (0-based)."""
    _shift_worksheet_element(ws_element, _Insert(False, at_col + 1, delta))


def shift_table_parts(worksheet, at_row: int, delta: int) -> None:
    """Shift table ``ref`` / ``autoFilter`` for tables linked from the worksheet."""
    tp = worksheet._part.element.find(_TABLE_PARTS)
    if tp is None:
        return
    for rel_elm in tp.findall(_TABLE_PART):
        r_id = rel_elm.get(f"{{{OFFICE_REL_NS}}}id")
        if r_id is None:
            continue
        try:
            part = worksheet._part.rels[r_id].target_part
        except KeyError:
            continue
        table_elm = parse_xml(part.blob)
        ref = table_elm.get("ref")
        if ref:
            table_elm.set("ref", shift_range_ref(ref, at_row, delta))
        af = table_elm.find(_AUTO_FILTER)
        if af is not None and af.get("ref"):
            af.set("ref", shift_range_ref(af.get("ref"), at_row, delta))
        part._blob = serialize_xml(table_elm)


def shift_table_parts_cols(worksheet, at_col: int, delta: int) -> None:
    """Shift table ``ref`` / ``autoFilter`` columns after column insertion."""
    tp = worksheet._part.element.find(_TABLE_PARTS)
    if tp is None:
        return
    for rel_elm in tp.findall(_TABLE_PART):
        r_id = rel_elm.get(f"{{{OFFICE_REL_NS}}}id")
        if r_id is None:
            continue
        try:
            part = worksheet._part.rels[r_id].target_part
        except KeyError:
            continue
        table_elm = parse_xml(part.blob)
        ref = table_elm.get("ref")
        if ref:
            table_elm.set("ref", shift_range_ref_cols(ref, at_col, delta))
        af = table_elm.find(_AUTO_FILTER)
        if af is not None and af.get("ref"):
            af.set("ref", shift_range_ref_cols(af.get("ref"), at_col, delta))
        part._blob = serialize_xml(table_elm)


def shift_sheet_data_rows(sheet_data: etree._Element, at_row: int, delta: int) -> None:
    """Bump ``r`` on rows and cells with row >= ``at_row``."""
    ins = _Insert(True, at_row, delta)
    _walk_sheet_data(sheet_data, ins, apply=False)
    _walk_sheet_data(sheet_data, ins, apply=True)


def shift_sheet_data_cols(sheet_data: etree._Element, at_col: int, delta: int) -> None:
    """Bump column letters on cells with col >= ``at_col`` (0-based)."""
    ins = _Insert(False, at_col + 1, delta)
    _walk_sheet_data(sheet_data, ins, apply=False)
    _walk_sheet_data(sheet_data, ins, apply=True)


def shift_col_dimensions(ws_element: etree._Element, at_col: int, delta: int) -> None:
    """Shift ``<col min/max>`` (1-based) for columns at or to the right of ``at_col`` (0-based)."""
    planner = _Planner(_Insert(False, at_col + 1, delta))
    planner.columns(ws_element)
    planner.apply()


def shift_dollar_ref(ref: str, at_row: int, delta: int) -> str:
    """Shift rows in an A1 ref, preserving ``$``; an unparseable ref is returned as is."""
    return _move_ref_text(ref, lambda area: area.shift_rows(at_row, delta))


def shift_dollar_ref_cols(ref: str, at_col: int, delta: int) -> str:
    """Shift columns in an A1 ref, preserving ``$`` (``at_col`` is 0-based)."""
    return _move_ref_text(ref, lambda area: area.shift_cols(at_col + 1, delta))


def shift_defined_names(workbook, sheet_name: str, at_row: int, delta: int) -> None:
    """Shift A1 rows in definedNames that refer to ``sheet_name``."""
    planner = _Planner(_Insert(True, at_row, delta), sheet_name=sheet_name)
    planner.defined_names(workbook)
    planner.apply()


def shift_defined_names_cols(workbook, sheet_name: str, at_col: int, delta: int) -> None:
    """Shift A1 columns in definedNames that refer to ``sheet_name``."""
    planner = _Planner(_Insert(False, at_col + 1, delta), sheet_name=sheet_name)
    planner.defined_names(workbook)
    planner.apply()

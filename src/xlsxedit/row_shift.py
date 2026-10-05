"""Shift a worksheet's cell references when rows or columns are inserted.

``plan_insert_rows`` and ``plan_insert_columns`` compute every change one
insertion makes before applying any of it; ``docs/features.md`` ("Inserting
rows and columns") lists what moves and when an insertion is refused. Ranges
move by the rules of :class:`xlsxedit.range_set.CellRange`; references in
text by ``rewrite_formula_refs`` (defined names) and ``rewrite_plain_ref``
(text that is one reference). Planning raises for anything it cannot do, so a
refused insertion changes nothing; ``apply`` only assigns computed values.

The ``check_*`` functions validate the rest of an insert up front. The
remaining ``shift_*`` functions are the older per-structure entry points.
"""

from __future__ import annotations

import math
import operator
import re
from collections.abc import Callable
from dataclasses import dataclass, replace
from datetime import date
from functools import partial

from lxml import etree

from xlsxedit.bulk_styles import normalize_style_spec
from xlsxedit.exceptions import (
    GridOverflowError,
    InvalidColorError,
    InvalidRangeError,
    TableError,
)
from xlsxedit.hyperlinks import _release_rel
from xlsxedit.opc.constants import OFFICE_REL_NS, SML_NS
from xlsxedit.oxml.address import MAX_COL, MAX_ROW, col_to_index, index_to_col
from xlsxedit.oxml.parser import parse_xml, serialize_xml
from xlsxedit.range_set import CellRange, SheetRangeSet
from xlsxedit.styles import _XML_ILLEGAL, check_style_text, datetime_to_serial, normalize_rgb
from xlsxedit.worksheet_order import insert_worksheet_child

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
_FILTER_COLUMN = f"{{{SML_NS}}}filterColumn"
_SORT_STATE = f"{{{SML_NS}}}sortState"
_SORT_CONDITION = f"{{{SML_NS}}}sortCondition"
_SHEETS = f"{{{SML_NS}}}sheets"
_SHEET = f"{{{SML_NS}}}sheet"
_SHEET_VIEWS = f"{{{SML_NS}}}sheetViews"
_SHEET_VIEW = f"{{{SML_NS}}}sheetView"
_CUSTOM_SHEET_VIEWS = f"{{{SML_NS}}}customSheetViews"
_CUSTOM_SHEET_VIEW = f"{{{SML_NS}}}customSheetView"
_PANE = f"{{{SML_NS}}}pane"
_SELECTION = f"{{{SML_NS}}}selection"
_ROW_BREAKS = f"{{{SML_NS}}}rowBreaks"
_COL_BREAKS = f"{{{SML_NS}}}colBreaks"
_BRK = f"{{{SML_NS}}}brk"
_PROTECTED_RANGES = f"{{{SML_NS}}}protectedRanges"
_PROTECTED_RANGE = f"{{{SML_NS}}}protectedRange"
_IGNORED_ERRORS = f"{{{SML_NS}}}ignoredErrors"
_IGNORED_ERROR = f"{{{SML_NS}}}ignoredError"
_CELL_WATCHES = f"{{{SML_NS}}}cellWatches"
_CELL_WATCH = f"{{{SML_NS}}}cellWatch"
_SCENARIOS = f"{{{SML_NS}}}scenarios"
_SCENARIO = f"{{{SML_NS}}}scenario"
_INPUT_CELLS = f"{{{SML_NS}}}inputCells"
_SMART_TAGS = f"{{{SML_NS}}}smartTags"
_CELL_SMART_TAGS = f"{{{SML_NS}}}cellSmartTags"
_EXT_LST = f"{{{SML_NS}}}extLst"

_X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
_XM_NS = "http://schemas.microsoft.com/office/excel/2006/main"
_XM_SQREF = f"{{{_XM_NS}}}sqref"
_XM_F = f"{{{_XM_NS}}}f"
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
# x14 elements whose xm:f children hold references to the sheet's cells
_X14_FORMULA_OWNERS = frozenset(
    f"{{{_X14_NS}}}{name}" for name in ("conditionalFormatting", "dataValidation", "sparklineGroup")
)
_CONTENT = frozenset({f"{{{SML_NS}}}v", _F, f"{{{SML_NS}}}is"})
_RANGED_FORMULAS = frozenset({"shared", "array", "dataTable"})
_CFVO = f"{{{SML_NS}}}cfvo"
_WRITABLE = (str, int, float, date)  # bool is an int and datetime a date

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
_PLAIN_REF_RE = re.compile(rf"\s*(?:(?P<sheet>{_SHEET_NAME})!)?(?P<ref>{_REF})\s*")


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

    def pinned(self, area: CellRange) -> CellRange:
        """``area`` moved, or pinned to the last row or column when pushed off (view state)."""
        moved = self.area(area)
        if moved is not None:
            return moved
        if self.rows:
            return replace(area, min_row=MAX_ROW, max_row=MAX_ROW, text=None)
        return replace(area, min_col=MAX_COL, max_col=MAX_COL, text=None)

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


def _refresh_manual_break_count(block: etree._Element) -> None:
    if block.get("manualBreakCount") is not None:
        manual = sum(1 for b in block.iterchildren(_BRK) if b.get("man") in ("1", "true"))
        block.set("manualBreakCount", str(manual))


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
            if apply and r is None:
                row.set("r", str(number))
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


def rewrite_plain_ref(
    text: str,
    sheet_name: str,
    move: Callable[[CellRange], CellRange | None],
    *,
    unqualified: bool,
) -> str:
    """Move ``text`` when the whole of it is one reference to ``sheet_name``; else leave it."""
    m = _PLAIN_REF_RE.fullmatch(text)
    if m is None:
        return text
    sheet = m.group("sheet")
    if (sheet is None and not unqualified) or (sheet is not None and not _same_sheet(sheet, sheet_name)):
        return text
    return text[: m.start("ref")] + _move_ref_text(m.group("ref"), move) + text[m.end("ref") :]


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

    def area_attr(
        self, elm: etree._Element, attr: str, where: str, *, lost: str, ins: _Insert | None = None
    ) -> None:
        """Move an ``ST_Ref`` attribute. When it is pushed off the grid, ``lost`` says
        what happens: ``"remove"`` the element, ``"pin"`` it to the edge, or ``"raise"``."""
        text = elm.get(attr)
        if not text:
            return
        if lost == "pin":
            try:
                area = CellRange.parse(text)
            except InvalidRangeError:
                return  # view state only; never block an insertion over it
            pinned = (ins or self.ins).pinned(area)
            if pinned is not area:
                self.set(elm, attr, str(pinned))
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

    def view_sqref_attr(self, elm: etree._Element, attr: str) -> None:
        text = elm.get(attr)
        if not text:
            return
        try:
            ranges = SheetRangeSet.parse(text)
        except InvalidRangeError:
            return
        moved = [self.ins.pinned(area) for area in ranges]
        if any(new is not old for new, old in zip(moved, ranges)):
            self.set(elm, attr, str(SheetRangeSet(dict.fromkeys(moved))))

    # -- worksheet structures -------------------------------------------

    def worksheet(self, root: etree._Element) -> None:
        """Every reference carrier inside the worksheet part other than ``sheetData``."""
        self.merge_cells(root)
        for cf in root.iterchildren(_CF):
            self.sqref_attr(cf, "sqref", "conditionalFormatting/@sqref")
        self.cfvo_values(root, own=True)
        for dv in root.iterfind(f"{_DATA_VALIDATIONS}/{_DATA_VALIDATION}"):
            self.sqref_attr(dv, "sqref", "dataValidation/@sqref")
        self.hyperlinks(root)
        self.auto_filter(root, "")
        self.sort_state(root, "")
        self.breaks(root)
        self.views(root)
        for elm in root.iterfind(f"{_PROTECTED_RANGES}/{_PROTECTED_RANGE}"):
            self.sqref_attr(elm, "sqref", "protectedRange/@sqref")
        for elm in root.iterfind(f"{_IGNORED_ERRORS}/{_IGNORED_ERROR}"):
            self.sqref_attr(elm, "sqref", "ignoredError/@sqref")
        for elm in root.iterfind(f"{_CELL_WATCHES}/{_CELL_WATCH}"):
            self.area_attr(elm, "r", "cellWatch/@r", lost="remove")
        for elm in root.iterfind(f"{_SMART_TAGS}/{_CELL_SMART_TAGS}"):
            self.area_attr(elm, "r", "cellSmartTags/@r", lost="remove")
        scenarios = root.find(_SCENARIOS)
        if scenarios is not None:
            self.view_sqref_attr(scenarios, "sqref")
            for elm in scenarios.iterfind(f"{_SCENARIO}/{_INPUT_CELLS}"):
                self.area_attr(elm, "r", "scenario inputCells/@r", lost="raise")
        if not self.ins.rows:
            self.columns(root)
            self.row_spans(root.find(_SHEET_DATA))
        self.extensions(root)

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

    def auto_filter(self, container: etree._Element, where: str) -> None:
        af = container.find(_AUTO_FILTER)
        if af is None:
            return
        ref = af.get("ref")
        if ref:
            area = _area(ref, f"{where}autoFilter/@ref")
            moved = self.ins.area(area)
            if moved is None:
                self.remove(af)
                return
            if moved is not area:
                self.set(af, "ref", str(moved))
            if not self.ins.rows and area.min_col < self.ins.at <= area.max_col:
                # filterColumn/@colId counts from the range's first column
                first_moved = self.ins.at - area.min_col
                for column in af.iterchildren(_FILTER_COLUMN):
                    col_id = column.get("colId")
                    if col_id is not None and col_id.isdigit() and int(col_id) >= first_moved:
                        self.set(column, "colId", str(int(col_id) + self.ins.count))
        self.sort_state(af, f"{where}autoFilter/")

    def sort_state(self, container: etree._Element, where: str) -> None:
        state = container.find(_SORT_STATE)
        if state is None:
            return
        ref = state.get("ref")
        if ref:
            area = _area(ref, f"{where}sortState/@ref")
            moved = self.ins.area(area)
            if moved is None:
                self.remove(state)
                return
            if moved is not area:
                self.set(state, "ref", str(moved))
        for condition in state.iterchildren(_SORT_CONDITION):
            self.area_attr(condition, "ref", f"{where}sortCondition/@ref", lost="remove")

    def breaks(self, container: etree._Element) -> None:
        """Move manual page breaks on the shifted axis.

        ``brk/@id`` is the last row (column) before the break; a break belongs
        to the row after it, so it moves when that row does.
        """
        block = container.find(_ROW_BREAKS if self.ins.rows else _COL_BREAKS)
        if block is None:
            return
        removed = False
        for brk in block.iterchildren(_BRK):
            brk_id = _number(brk.get("id", "0"), "brk/@id")
            if brk_id + 1 < self.ins.at:
                continue
            new_id = brk_id + self.ins.count
            if new_id >= self.ins.limit:
                self.remove(brk)
                removed = True
            else:
                self.set(brk, "id", str(new_id))
        if removed:
            self.after(partial(_refresh_manual_break_count, block))

    def views(self, root: etree._Element) -> None:
        """Selections follow their cells. The scroll positions (``topLeftCell``)
        are places on screen: they move only for an insertion strictly above
        (left of) them, so new rows inserted at the top of a view stay in sight."""
        custom = list(root.iterfind(f"{_CUSTOM_SHEET_VIEWS}/{_CUSTOM_SHEET_VIEW}"))
        screen = _Insert(self.ins.rows, self.ins.at + 1, self.ins.count)
        for view in [*root.iterfind(f"{_SHEET_VIEWS}/{_SHEET_VIEW}"), *custom]:
            top = self._view_start(view.get("topLeftCell"))
            self.area_attr(view, "topLeftCell", "topLeftCell", lost="pin", ins=screen)
            pane = view.find(_PANE)
            if pane is not None:
                self.frozen_split(pane, top)
                self.area_attr(pane, "topLeftCell", "pane/@topLeftCell", lost="pin", ins=screen)
            for selection in view.iterchildren(_SELECTION):
                self.area_attr(selection, "activeCell", "activeCell", lost="pin")
                self.view_sqref_attr(selection, "sqref")
        for view in custom:
            self.breaks(view)
            self.auto_filter(view, "customSheetView/")

    def _view_start(self, top_left: str | None) -> int:
        """Row (column) of a view's top-left cell on the shifted axis; 1 if unreadable."""
        try:
            area = CellRange.parse(top_left or "A1")
        except InvalidRangeError:
            return 1
        return area.min_row if self.ins.rows else area.min_col

    def frozen_split(self, pane: etree._Element, top: int) -> None:
        """Grow a frozen split when rows (columns) are inserted inside the frozen band.

        The split counts rows (columns) from the view's top-left cell, so the
        band is ``top .. top + split - 1``.
        """
        if pane.get("state") not in ("frozen", "frozenSplit"):
            return  # a plain split is measured in twips, not cells
        attr = "ySplit" if self.ins.rows else "xSplit"
        try:
            split = float(pane.get(attr, "0"))
        except ValueError:
            return
        if split > 0 and top <= self.ins.at < top + split:
            new = split + self.ins.count
            self.set(pane, attr, str(int(new)) if new.is_integer() else repr(new))

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

    def row_spans(self, sheet_data: etree._Element | None) -> None:
        """Move each row's ``spans`` hint (columns of its block's cells)."""
        if sheet_data is None:
            return
        for row in sheet_data.iterchildren(_ROW):
            spans = row.get("spans")
            if not spans:
                continue
            intervals = _parse_spans(spans)
            if intervals is None:
                continue  # an optimisation hint only; leave what we cannot read
            moved = []
            for lo, hi in intervals:
                if hi >= self.ins.at:
                    lo, hi = self.ins.move(lo), min(hi + self.ins.count, MAX_COL)
                if lo <= MAX_COL:
                    moved.append(f"{lo}:{hi}")
            if moved:
                self.set(row, "spans", " ".join(moved))

    def extensions(self, root: etree._Element) -> None:
        """``xm:sqref`` anywhere in the worksheet's extensions, and x14 ``xm:f`` references."""
        ext_lst = root.find(_EXT_LST)
        if ext_lst is None:
            return
        for sqref in ext_lst.iter(_XM_SQREF):
            ranges = _range_set(sqref.text or "", "xm:sqref")
            moved = self.ins.ranges(ranges)
            if moved is ranges:
                continue
            if moved:
                self.set_text(sqref, str(moved))
            else:
                self.remove(sqref.getparent())
        self.extension_formulas(ext_lst, own=True)

    def extension_formulas(self, ext_lst: etree._Element, *, own: bool) -> None:
        """Plain-reference ``xm:f`` under x14 conditional formats, validations and sparklines.

        ``own`` is whether the extensions belong to the shifted sheet, where an
        unqualified reference points at it.
        """
        move = self.ins.area
        for f in ext_lst.iter(_XM_F):
            if not f.text or not any(a.tag in _X14_FORMULA_OWNERS for a in f.iterancestors()):
                continue
            new = rewrite_plain_ref(f.text, self.sheet_name, move, unqualified=own)
            self.set_text(f, new)

    def cfvo_values(self, root: etree._Element, *, own: bool) -> None:
        """Plain-reference ``cfvo/@val`` of conditional formats, which an x14
        ``xm:f`` twin repeats; the two must stay equal."""
        for cf in root.iterchildren(_CF):
            for cfvo in cf.iter(_CFVO):
                val = cfvo.get("val")
                if val:
                    new = rewrite_plain_ref(val, self.sheet_name, self.ins.area, unqualified=own)
                    self.set(cfvo, "val", new)

    def hyperlink_locations(self, root: etree._Element, *, own: bool) -> None:
        """Internal hyperlinks only: with ``r:id``, ``location`` is inside another document."""
        for link in root.iterfind(f"{_HYPERLINKS}/{_HYPERLINK}"):
            location = link.get("location")
            if location and link.get(f"{{{OFFICE_REL_NS}}}id") is None:
                new = rewrite_plain_ref(location, self.sheet_name, self.ins.area, unqualified=own)
                self.set(link, "location", new)

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


def _parse_spans(spans: str) -> list[tuple[int, int]] | None:
    intervals = []
    for item in spans.split():
        lo, sep, hi = item.partition(":")
        if not (sep and lo.isdigit() and hi.isdigit()):
            return None
        intervals.append((int(lo), int(hi)))
    return intervals


_DEFERRED_MARKERS = (_XM_NS.encode(), b"location=")


def _root_if_it_may_refer(worksheet) -> etree._Element | None:
    """The worksheet's element, unless it is still unparsed and cannot hold a
    reference this module rewrites (keeps ``Workbook.open(large=True)`` lazy)."""
    blob = worksheet._part.unparsed_blob
    if blob is not None and not any(marker in blob for marker in _DEFERRED_MARKERS):
        return None
    return worksheet._part.element


class InsertPlan:
    """Every change that inserting rows or columns into one worksheet makes.

    Building the plan reads the worksheet, its table parts, the workbook's
    defined names and the other sheets' references to it, and raises
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
        planner.hyperlink_locations(root, own=True)
        workbook = worksheet._workbook
        for other in workbook.worksheets:
            if other is worksheet:
                continue
            other_root = _root_if_it_may_refer(other)
            if other_root is None:
                continue
            ext_lst = other_root.find(_EXT_LST)
            if ext_lst is not None:
                planner.extension_formulas(ext_lst, own=False)
                planner.cfvo_values(other_root, own=False)
            planner.hyperlink_locations(other_root, own=False)
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
        """Refuse an insertion that would push any part of a table off the grid
        (its columns and rows cannot shrink with it), or that would add columns
        inside a table bound to a query."""
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
            if not self._insert.rows:
                _check_table_accepts_columns(table, self._insert.at - 1)

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
    a colour ``normalize_rgb`` rejects, a ``font_size`` that is not a number, or
    text and alignments ``check_style_text`` rejects."""
    for specs in spec_lists:
        for spec in specs or ():
            normalized = normalize_style_spec(spec)
            check_style_text(normalized)
            for key in ("font_color", "bg_color"):
                color = normalized.get(key)
                if color is not None:
                    if not isinstance(color, str):
                        raise InvalidColorError(f"invalid color: {color!r}")
                    normalize_rgb(color)
            if "font_size" in normalized:
                int(normalized["font_size"])


def template_row_merges(
    merges: list[str], template_rows: list[int], at_row: int, count: int
) -> list[CellRange]:
    """The one-row merges of each template row, copied onto the ``count`` rows from ``at_row``.

    ``template_rows`` cycle over the new rows and name rows as they are before
    the insert; merges of any ``ST_Ref`` form are read.
    """
    if not template_rows:
        return []
    bands: dict[int, list[CellRange]] = {row: [] for row in template_rows}
    for ref in merges:
        area = CellRange.parse(ref)
        if area.min_row == area.max_row and area.min_col != area.max_col and area.min_row in bands:
            bands[area.min_row].append(area)
    copies = []
    for offset in range(count):
        row = at_row + offset
        for area in bands[template_rows[offset % len(template_rows)]]:
            copies.append(CellRange.from_bounds(area.min_col, row, area.max_col, row))
    return copies


def append_merges(ws_element: etree._Element, areas: list[CellRange]) -> None:
    """Add ``areas`` as ``<mergeCell>`` elements in one pass (checked beforehand)."""
    if not areas:
        return
    block = ws_element.find(_MERGE_CELLS)
    if block is None:
        block = etree.Element(_MERGE_CELLS)
        insert_worksheet_child(ws_element, block)
    for area in areas:
        etree.SubElement(block, _MERGE_CELL).set("ref", str(area))
    block.set("count", str(sum(1 for _ in block.iterchildren(_MERGE_CELL))))


def check_merges_disjoint(existing: list[CellRange], new: list[CellRange]) -> None:
    """Raise ``ValueError``, as ``Worksheet.merge_cells`` does, when a merge in
    ``new`` overlaps one in ``existing``. ``new`` holds single-row merges copied
    from one template row each, so they cannot overlap one another."""
    for area in new:
        for other in existing:
            if area.intersects(other):
                raise ValueError(f"merge range {str(area)!r} overlaps existing {str(other)!r}")


def _column_key(c: etree._Element) -> int:
    address = c.get("r")
    m = _CELL_ADDRESS_RE.fullmatch(address or "")
    return col_to_index(m.group(1)) if m else -1


def sort_row_cells(sheet_data: etree._Element, first_row: int, last_row: int) -> None:
    """Put the ``<c>`` children of rows ``first_row..last_row`` back in column order.

    The bulk writer appends cells; after ``insert_columns`` new cells land
    after the moved ones, and Excel repairs a row whose cells are out of order.
    """
    for row in sheet_data.iterchildren(_ROW):
        r = row.get("r")
        if r is None or not r.isdigit() or not first_row <= int(r) <= last_row:
            continue
        children = [child for child in row if isinstance(child.tag, str)]
        cells = [child for child in children if child.tag == _C]
        keys = [_column_key(c) for c in cells]
        if children[: len(cells)] == cells and all(a < b for a, b in zip(keys, keys[1:])):
            continue
        # CT_Row is c* then extLst: cells first, in column order
        for index, (_, c) in enumerate(sorted(zip(keys, cells), key=lambda kc: kc[0])):
            row.insert(index, c)


def place_cell_in_order(c: etree._Element) -> None:
    """Move ``<c>`` before every sibling it should follow: a cell with a higher
    column, or the row's ``extLst``."""
    key = _column_key(c)
    first_later = None
    for sibling in c.itersiblings(preceding=True):
        if sibling.tag == _C and _column_key(sibling) < key:
            break
        first_later = sibling
    if first_later is not None:
        first_later.addprevious(c)


def widen_row_spans(sheet_data: etree._Element, first_row: int, last_row: int) -> None:
    """Grow the ``spans`` hint of rows ``first_row..last_row`` to cover their cells."""
    for row in sheet_data.iterchildren(_ROW):
        spans, r = row.get("spans"), row.get("r")
        if not spans or r is None or not r.isdigit() or not first_row <= int(r) <= last_row:
            continue
        intervals = _parse_spans(spans)
        cols = [
            _cell_column(c.get("r"), "c/@r") for c in row.iterchildren(_C) if c.get("r")
        ]
        if intervals is None or not cols:
            continue
        if all(any(lo <= col <= hi for lo, hi in intervals) for col in cols):
            continue
        lo = min(min(cols), *(lo for lo, _ in intervals))
        hi = max(max(cols), *(hi for _, hi in intervals))
        row.set("spans", f"{lo}:{hi}")


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


_TABLE_COLUMNS = f"{{{SML_NS}}}tableColumns"
_TABLE_COLUMN = f"{{{SML_NS}}}tableColumn"


def _inner_column_offset(ref: str | None, at_col: int) -> int | None:
    """Where 0-based ``at_col`` falls in table ``ref`` when inserting there widens the table.

    That is the case strictly inside the table: an insertion at its first
    column moves the whole table right, and one past its last column misses it.
    """
    if not ref:
        return None
    area = _area(ref, "table")
    first, last = area.min_col - 1, area.max_col - 1
    return at_col - first if first < at_col <= last else None


def _check_table_accepts_columns(table_elm: etree._Element, at_col: int) -> None:
    """Raise ``TableError`` when columns inserted at 0-based ``at_col`` would join
    a table bound to a query: its query table part would not know them."""
    if _inner_column_offset(table_elm.get("ref"), at_col) is None:
        return
    if any(c.get("queryTableFieldId") is not None for c in table_elm.iter(_TABLE_COLUMN)):
        raise TableError(
            f"cannot insert columns inside table {table_elm.get('displayName')!r} "
            f"({table_elm.get('ref')}): its columns are bound to a query "
            "(queryTableFieldId), and the query table part is not updated"
        )


def _insert_table_columns(table_elm: etree._Element, offset: int, delta: int) -> None:
    """Add ``delta`` ``tableColumn``s before the one at ``offset``, like Excel names them.

    Each gets the next free ``id`` and the first unused ``ColumnN`` name;
    ``autoFilter`` ``filterColumn`` ids at or after ``offset`` move with
    their columns.
    """
    cols_elm = table_elm.find(_TABLE_COLUMNS)
    if cols_elm is None:
        return
    existing = cols_elm.findall(_TABLE_COLUMN)
    next_id = max((int(c.get("id")) for c in existing if c.get("id", "").isdigit()), default=0)
    used = {(c.get("name") or "").casefold() for c in existing}
    anchor = existing[offset] if offset < len(existing) else None
    n = 0
    for _ in range(delta):
        n += 1
        while f"column{n}" in used:
            n += 1
        next_id += 1
        column = etree.Element(_TABLE_COLUMN)
        column.set("id", str(next_id))
        column.set("name", f"Column{n}")
        if anchor is None:
            cols_elm.append(column)
        else:
            anchor.addprevious(column)
    cols_elm.set("count", str(len(existing) + delta))
    af = table_elm.find(_AUTO_FILTER)
    if af is not None:
        for fc in af.iterchildren(_FILTER_COLUMN):
            col_id = fc.get("colId", "")
            if col_id.isdigit() and int(col_id) >= offset:
                fc.set("colId", str(int(col_id) + delta))


def shift_table_parts_cols(worksheet, at_col: int, delta: int) -> None:
    """Shift table ``ref`` / ``autoFilter`` columns after column insertion.

    Columns inserted strictly inside a table join it as new ``tableColumn``s
    (see ``_insert_table_columns``). Raises ``TableError``, before changing
    any table part, when such a table is bound to a query: its query table
    part would not know the new columns. ``plan_insert_columns`` makes the
    same check before anything on the sheet moves.
    """
    tp = worksheet._part.element.find(_TABLE_PARTS)
    if tp is None:
        return
    tables = []
    for rel_elm in tp.findall(_TABLE_PART):
        r_id = rel_elm.get(f"{{{OFFICE_REL_NS}}}id")
        if r_id is None:
            continue
        try:
            part = worksheet._part.rels[r_id].target_part
        except KeyError:
            continue
        table_elm = parse_xml(part.blob)
        tables.append((part, table_elm, _inner_column_offset(table_elm.get("ref"), at_col)))
    for _part, table_elm, _offset in tables:
        _check_table_accepts_columns(table_elm, at_col)
    for part, table_elm, offset in tables:
        if offset is not None:
            _insert_table_columns(table_elm, offset, delta)
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

"""Shift a worksheet's cell references when rows or columns are inserted.

``plan_insert_rows`` and ``plan_insert_columns`` compute every change one
insertion makes before applying any of it; ``docs/features.md`` ("Inserting
rows and columns") lists what moves and when an insertion is refused. Ranges
move by the rules of :class:`xlsxedit.range_set.CellRange`; references in
text by ``rewrite_formula_refs`` (defined names, sparklines, other sheets'
formats and validations), ``rewrite_anchored_refs`` (formulas of the sheet's
own formats and validations, relative to the top-left cell of their areas)
and ``rewrite_plain_ref`` (hyperlink locations, which are one reference).
Planning raises for anything it cannot do, so a refused insertion changes
nothing; ``apply`` only assigns computed values.

A :class:`_Shift` also describes a delete and a shift restricted to a span of
columns (rows), and the :class:`_Planner` moves every structure it reaches as
Excel does for those. ``InsertPlan``, which also moves the cells, carries out
whole-row and whole-column inserts; ``CellShiftPlan`` (``shift_cells``)
carries out inserts and deletes restricted to a span, Excel's "insert cells"
and "delete cells", with Excel's refusals. No public method runs it yet.

The ``check_*`` functions validate the rest of an insert up front. The
``shift_*`` functions other than ``shift_cells`` are thin wrappers kept for
the table code and tests.
"""

from __future__ import annotations

import bisect
import math
import operator
import re
from collections.abc import Callable
from dataclasses import dataclass, replace
from datetime import date
from functools import partial

from lxml import etree

from xlsxedit.bulk_styles import normalize_style_spec
from xlsxedit.cell import _refuse_range_formula_edit
from xlsxedit.drawing import (
    _XDR_FROM,
    _XDR_TO,
    _iter_anchors,
    _marker,
    chart_parts,
    chart_references,
    drawing_parts_for_worksheet,
)
from xlsxedit.exceptions import (
    FormulaGroupError,
    GridOverflowError,
    InvalidRangeError,
    TableError,
)
from xlsxedit.hyperlinks import _release_rel
from xlsxedit.opc.constants import OFFICE_REL_NS, SML_NS
from xlsxedit.oxml.address import MAX_COL, MAX_ROW, col_to_index, index_to_col, join_address
from xlsxedit.parts import WorksheetPart
from xlsxedit.range_set import CellRange, SheetRangeSet, shift_bounds
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
_X14_SPARKLINE_GROUP = f"{{{_X14_NS}}}sparklineGroup"
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
_FORMULA = f"{{{SML_NS}}}formula"
_FORMULA1 = f"{{{SML_NS}}}formula1"
_FORMULA2 = f"{{{SML_NS}}}formula2"
_WRITABLE = (str, int, float, date)  # bool is an int and datetime a date
_MAX_INT_BITS = 2_000  # at most 603 digits: under the lowest digit limit str() allows (640)

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


@dataclass(frozen=True)
class _Shift:
    """Rows (``rows``) or columns moved by an insert or a delete.

    ``delta > 0`` inserts that many rows before row ``at`` (columns before
    1-based column ``at``); ``delta < 0`` deletes ``-delta`` of them from
    ``at``. ``span`` restricts the shift to columns (rows) ``lo..hi``, as
    Excel's "insert cells" and "delete cells" do: only areas within it move,
    and whole rows (columns) stay. An area the shift :meth:`splits` cannot
    follow it, so a caller must refuse such a shift. Areas move by
    :meth:`CellRange.shift`.
    """

    rows: bool
    at: int
    delta: int
    span: tuple[int, int] | None = None

    def __post_init__(self) -> None:
        if self.delta == 0 or self.at < 1:
            raise ValueError(f"not a shift: at={self.at}, delta={self.delta}")
        if self.at > self.limit + (1 if self.inserting else 0):
            raise InvalidRangeError(f"line {self.at} is outside the grid")
        if not self.inserting and self.end > self.limit:
            raise InvalidRangeError(f"{self.describe()} passes the edge of the grid")
        if self.span is not None:
            lo, hi = self.span
            cross = MAX_COL if self.rows else MAX_ROW
            if not 1 <= lo <= hi <= cross:
                raise InvalidRangeError(f"span {lo}..{hi} is outside the grid")
            if (lo, hi) == (1, cross):
                object.__setattr__(self, "span", None)

    @property
    def limit(self) -> int:
        return MAX_ROW if self.rows else MAX_COL

    @property
    def inserting(self) -> bool:
        return self.delta > 0

    @property
    def end(self) -> int:
        """The last line of the band: the last new line, or the last deleted one."""
        return self.at + abs(self.delta) - 1

    def move(self, index: int) -> int | None:
        """Where 1-based line ``index`` goes, ignoring ``span``; ``None`` when it is
        deleted. An insert may push it past the grid."""
        if index < self.at:
            return index
        if self.inserting:
            return index + self.delta
        return None if index <= self.end else index + self.delta

    def source(self, index: int) -> int | None:
        """The line :meth:`move` takes to ``index``; ``None`` for an inserted line
        or one a delete brings in from past the grid."""
        if index < self.at:
            return index
        if self.inserting and index <= self.end:
            return None
        index -= self.delta
        return index if index <= self.limit else None

    def cell(self, col: int, row: int, *, back: bool = False) -> tuple[int, int] | None:
        """Where cell ``(col, row)`` goes, or with ``back`` where it comes from;
        ``None`` when it is deleted, new, or off the grid."""
        index, cross = (row, col) if self.rows else (col, row)
        if self.span is None or self.span[0] <= cross <= self.span[1]:
            index = self.source(index) if back else self.move(index)
            if index is None or index > self.limit:
                return None
        return (cross, index) if self.rows else (index, cross)

    def area(self, area: CellRange) -> CellRange | None:
        return area.shift(self.rows, self.at, self.delta, self.span)

    def ranges(self, ranges: SheetRangeSet) -> SheetRangeSet:
        return ranges.shift(self.rows, self.at, self.delta, self.span)

    def interval(self, lo: int, hi: int) -> tuple[int, int] | None:
        """Lines ``lo..hi`` of a whole-line structure (``<col>``, ``spans``), ignoring ``span``."""
        return shift_bounds(lo, hi, self.at, self.delta, self.limit)

    def _extents(self, area: CellRange) -> tuple[int, int, int, int]:
        """``area``'s first and last line on the shifted axis, then across it."""
        if self.rows:
            return area.min_row, area.max_row, area.min_col, area.max_col
        return area.min_col, area.max_col, area.min_row, area.max_row

    def clamps(self, area: CellRange) -> bool:
        """Whether an insert pushes any of ``area`` past the grid."""
        _, hi, lo_x, hi_x = self._extents(area)
        full = area.spans_all_rows if self.rows else area.spans_all_cols
        within = self.span is None or self.span[0] <= lo_x and hi_x <= self.span[1]
        return self.inserting and within and not full and hi >= self.at and hi + self.delta > self.limit

    def splits(self, area: CellRange) -> bool:
        """Whether the shift would move some of ``area``'s cells and not others:
        ``area`` crosses an edge of ``span`` and reaches the shifted region."""
        if self.span is None:
            return False
        _, hi, lo_x, hi_x = self._extents(area)
        lo_s, hi_s = self.span
        crosses = lo_x < lo_s <= hi_x or lo_x <= hi_s < hi_x
        return crosses and hi >= self.at

    def tears(self, area: CellRange) -> bool:
        """Whether the shift would move or delete some of ``area``'s cells and
        keep others where they are: it :meth:`splits` the area, the area
        straddles the inserted lines, or a delete takes only some of its lines."""
        if self.splits(area):
            return True
        lo, hi, lo_x, hi_x = self._extents(area)
        if self.span is not None and not self.span[0] <= lo_x <= hi_x <= self.span[1]:
            return False
        if self.inserting:
            return lo < self.at <= hi
        return lo <= self.end and hi >= self.at and not self.at <= lo <= hi <= self.end

    def carries(self, cells: CellRange) -> bool:
        """Whether the shift moves a drawing object over ``cells`` whole, without
        resizing it: every column (row) of it lies in ``span``, and it starts
        past the shifted lines, at or after an insert's first new line or after
        a delete's last deleted one.

        Excel's "insert cells" moves only such objects (recorded 2026-10-29); an
        object across an edge of the span, or one the inserted lines cut
        through, stays where it is. Its placement (``editAs``) still decides
        how the anchor follows: an absolute object keeps its place on the page.
        """
        lo, _, lo_x, hi_x = self._extents(cells)
        if self.span is not None and not self.span[0] <= lo_x <= hi_x <= self.span[1]:
            return False
        return lo >= self.at if self.inserting else lo > self.end

    def pinned(self, area: CellRange) -> CellRange:
        """``area`` moved or, for view state, pinned where it went: the last
        line when an insert pushes it off, the band's first line when deleted."""
        moved = self.area(area)
        if moved is not None:
            return moved
        line = self.limit if self.inserting else self.at
        if self.rows:
            return replace(area, min_row=line, max_row=line, text=None)
        return replace(area, min_col=line, max_col=line, text=None)

    @property
    def screen(self) -> _Shift | None:
        """The shift of a view's scroll position, which an insert at the view's top
        leaves in place; ``None`` for a span, which moves no whole rows (columns)."""
        if self.span is not None:
            return None
        if not self.inserting or self.at > self.limit:
            return self
        return replace(self, at=self.at + 1)

    def describe(self) -> str:
        noun = "row" if self.rows else "column"
        if self.inserting:
            where = f"row {self.at}" if self.rows else f"column {index_to_col(self.at - 1)}"
            text = f"inserting {self.delta} {noun}(s) at {where}"
        elif self.rows:
            text = f"deleting rows {self.at}:{self.end}"
        else:
            text = f"deleting columns {index_to_col(self.at - 1)}:{index_to_col(self.end - 1)}"
        if self.span is None:
            return text
        lo, hi = self.span
        if self.rows:
            return f"{text} in columns {index_to_col(lo - 1)}:{index_to_col(hi - 1)}"
        return f"{text} in rows {lo}:{hi}"

    def overflow(self, what: str, *, excel: str | None = None) -> GridOverflowError:
        end = f"row {MAX_ROW}" if self.rows else "column XFD"
        return GridOverflowError(_excel_says(f"{self.describe()} would push {what} past {end}", excel))

    def refusal(self, what: str) -> InvalidRangeError:
        """The error for a shift that would push ``what`` off the grid, or delete it."""
        if self.inserting:
            return self.overflow(what)
        return InvalidRangeError(f"{self.describe()} would delete {what}")

    def split_error(self, what: str, *, excel: str | None = None) -> InvalidRangeError:
        return InvalidRangeError(_excel_says(f"{self.describe()} would move only part of {what}", excel))


# What Excel says when it refuses, or asks before, an "insert cells" or "delete
# cells" (Excel 16.0.20527 for Microsoft 365, recorded 2026-10-29).
_EXCEL_TABLE = "This won't work because it would move cells in a table on your worksheet."
_EXCEL_MERGE = "This operation will cause some merged cells to unmerge.  Do you wish to continue?"
_EXCEL_ARRAY = "You can't change part of an array."
_EXCEL_DATA_TABLE = "Can't change part of a data table."
_EXCEL_GRID = (
    "Microsoft Excel can't insert new cells because it would push non-empty cells off the end"
    " of the worksheet."
)
_EXCEL_OBJECTS = "Can't push objects off the sheet."


def _excel_says(message: str, excel: str | None) -> str:
    return message if excel is None else f"{message} (Excel: {excel})"


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
    """1-based column of a ``<c r>`` address, which must name a row from 1."""
    m = _CELL_ADDRESS_RE.fullmatch(address)
    index = col_to_index(m.group(1)) + 1 if m else 0
    if not 1 <= index <= MAX_COL or int(m.group(2)) < 1:
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


def _walk_sheet_data(sheet_data: etree._Element | None, ins: _Shift, *, apply: bool) -> None:
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
    sheet_data: etree._Element, row: etree._Element, number: int, ins: _Shift, *, apply: bool
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


def _move_cells_right(row: etree._Element, number: int, ins: _Shift, *, apply: bool) -> None:
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


def _move_formula_ranges(f: etree._Element, ins: _Shift, *, apply: bool, cell: str) -> None:
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
    sheet_name: str | None,
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
    pushed off the grid or deleted becomes ``#REF!``.
    """

    def visit(sheet: str | None, ref: str, external: bool) -> str:
        if sheet is None:
            return _move_ref_text(ref, move) if unqualified else ref
        if external or not _same_sheet(sheet, sheet_name):
            return ref
        return _move_ref_text(ref, move)

    return _map_refs(text, visit)


def translate_formula_refs(text: str, cols: int, rows: int) -> str:
    """Formula ``text`` as if copied ``cols`` columns right and ``rows`` rows down.

    The relative part of every reference moves, whatever sheet or workbook it
    names; one that would leave the grid becomes ``#REF!``.
    """
    move = partial(CellRange.translate, cols=cols, rows=rows)
    return _map_refs(text, lambda sheet, ref, external: _move_ref_text(ref, move))


def _map_refs(text: str, visit: Callable[[str | None, str, bool], str]) -> str:
    """``text`` with each reference token replaced by ``visit(sheet, ref, external)``:
    its sheet prefix (``None`` without one), the reference, and whether a
    workbook reference (``[1]``) precedes the prefix."""
    out: list[str] = []
    after_bracket = False
    for m in _FORMULA_TOKEN_RE.finditer(text):
        token = m.group(0)
        qref, ref = m.group("qref"), m.group("ref")
        if qref is not None and qref != "#REF!":
            token = f"{m.group('sheet')}!{visit(m.group('sheet'), qref, after_bracket)}"
        elif ref is not None and text[m.start() - 1 : m.start()] not in ("!", "]"):
            token = visit(None, ref, False)
        out.append(token)
        after_bracket = m.group("bracket") is not None
    return "".join(out)


def rewrite_anchored_refs(
    text: str,
    sheet_name: str | None,
    shift: _Shift,
    before: SheetRangeSet,
    after: SheetRangeSet,
) -> str:
    """Move the references in ``text``, a formula Excel evaluates relative to the
    top-left cell of ``before`` (a conditional format's or a validation's
    areas), which ``shift`` turns into ``after``.

    The result is what Excel writes for the cell that becomes the top-left of
    ``after``: the formula is first re-anchored to the cell it comes from,
    which differs when the old top-left cell is deleted, and its references
    then move as seen from there. Excel adjusts every cell and splits the
    format where cells would need different formulas; here one formula keeps
    serving every area.

    When the new top-left cell is itself inserted, because the format spans
    every row (column) and stays where it is, or because a shift restricted
    to a span split the format's areas, its relative references stay and only
    the absolute ones move. Excel does the same for most formulas over every
    row, but moves some ranges that mix relative and absolute bounds as a
    whole.
    """
    old, new = before.top_left, after.top_left
    if old is None or new is None:
        return text
    origin = shift.cell(*new, back=True)
    if origin is None:
        return rewrite_formula_refs(text, sheet_name, partial(_move_absolute, shift), unqualified=True)
    if origin != old:
        text = translate_formula_refs(text, origin[0] - old[0], origin[1] - old[1])
    return rewrite_formula_refs(text, sheet_name, shift.area, unqualified=True)


def _move_absolute(shift: _Shift, area: CellRange) -> CellRange | None:
    """``area`` with only its absolute (``$``) bounds on the shifted axis moved;
    bounds that cross are swapped with their ``$``, as Excel orders a range."""
    moved = shift.area(area)
    if moved is area:
        return area
    i, j = (1, 3) if shift.rows else (0, 2)
    lo_abs, hi_abs = area.absolute[i], area.absolute[j]
    if not (lo_abs or hi_abs):
        return area
    if moved is None or (lo_abs and hi_abs):
        return moved
    lo_attr, hi_attr = ("min_row", "max_row") if shift.rows else ("min_col", "max_col")
    lo = getattr(moved if lo_abs else area, lo_attr)
    hi = getattr(moved if hi_abs else area, hi_attr)
    if lo > hi:
        lo, hi, lo_abs, hi_abs = hi, lo, hi_abs, lo_abs
    flags = list(area.absolute)
    flags[i], flags[j] = lo_abs, hi_abs
    return replace(area, **{lo_attr: lo, hi_attr: hi}, absolute=tuple(flags), text=None)


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
    """Collects the edits for one shift; ``apply`` performs them in order.

    Attribute and text assignments run first, then removals, then the
    ``after`` steps that act on other parts. A structure that belongs to
    whole rows (columns), such as a page break or a ``<col>``, ignores a
    shift restricted to a span.
    """

    def __init__(self, ins: _Shift, *, sheet_name: str | None = None, part=None):
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

    def drop(self, elm: etree._Element, attr: str) -> None:
        if elm.get(attr) is not None:
            self._sets.append(partial(elm.attrib.pop, attr))

    def remove(self, elm: etree._Element) -> None:
        self._removals.append(partial(_detach, elm))

    def after(self, step: Callable[[], None]) -> None:
        self._after.append(step)

    # -- generic reference carriers -------------------------------------

    def area_attr(
        self, elm: etree._Element, attr: str, where: str, *, lost: str, ins: _Shift | None = None
    ) -> None:
        """Move an ``ST_Ref`` attribute. When it is pushed off the grid or deleted,
        ``lost`` says what happens: ``"remove"`` the element, ``"pin"`` it where it
        went (:meth:`_Shift.pinned`), or ``"raise"``."""
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
                raise self.ins.refusal(f"{where} {text!r}")
            self.remove(elm)
        elif moved is not area:
            self.set(elm, attr, str(moved))

    def sqref_attr(
        self, elm: etree._Element, attr: str, where: str
    ) -> tuple[SheetRangeSet, SheetRangeSet] | None:
        """Move an ``ST_Sqref`` attribute, removing ``elm`` when every area is gone;
        return the areas before and after the shift."""
        text = elm.get(attr)
        if text is None:
            return None
        ranges = _range_set(text, where)
        moved = self.ins.ranges(ranges)
        if moved is not ranges:
            if moved:
                self.set(elm, attr, str(moved))
            else:
                self.remove(elm)
        return ranges, moved

    def view_sqref_attr(self, elm: etree._Element, attr: str) -> list[int] | None:
        """Move or pin view-state areas, merging any that land on one another;
        return each old area's index among the new ones, ``None`` if none moved."""
        text = elm.get(attr)
        if not text:
            return None
        try:
            ranges = SheetRangeSet.parse(text)
        except InvalidRangeError:
            return None
        moved = [self.ins.pinned(area) for area in ranges]
        if all(new is old for new, old in zip(moved, ranges)):
            return None
        kept = list(dict.fromkeys(moved))
        self.set(elm, attr, str(SheetRangeSet(kept)))
        return [kept.index(area) for area in moved]

    def selection(self, selection: etree._Element) -> None:
        self.area_attr(selection, "activeCell", "activeCell", lost="pin")
        index = self.view_sqref_attr(selection, "sqref")
        active = selection.get("activeCellId")
        if index and active is not None and active.isdigit() and int(active) < len(index):
            self.set(selection, "activeCellId", str(index[int(active)]))

    # -- worksheet structures -------------------------------------------

    def worksheet(self, root: etree._Element) -> None:
        """Every reference carrier inside the worksheet part other than ``sheetData``."""
        self.merge_cells(root)
        for cf in root.iterchildren(_CF):
            sqrefs = self.sqref_attr(cf, "sqref", "conditionalFormatting/@sqref")
            self.format_formulas(cf, partial(self.anchored, sqrefs=sqrefs))
        for dv in root.iterfind(f"{_DATA_VALIDATIONS}/{_DATA_VALIDATION}"):
            sqrefs = self.sqref_attr(dv, "sqref", "dataValidation/@sqref")
            self.validation_formulas(dv, partial(self.anchored, sqrefs=sqrefs))
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
            if self.ins.splits(area):
                raise self.ins.split_error(f"merged cells {ref}", excel=_EXCEL_MERGE)
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
            self.filter_columns(af, area, moved)
        self.sort_state(af, f"{where}autoFilter/")

    def filter_columns(self, af: etree._Element, area: CellRange, moved: CellRange) -> None:
        """``filterColumn/@colId`` counts from the range's first column: follow
        each column, and drop the filter of one deleted or pushed off the grid."""
        span = self.ins.span
        if self.ins.rows or span is not None and not span[0] <= area.min_row <= area.max_row <= span[1]:
            return  # the range's columns stay where they are
        for column in af.iterchildren(_FILTER_COLUMN):
            col_id = column.get("colId")
            if col_id is None or not col_id.isdigit():
                continue
            new = self.ins.move(area.min_col + int(col_id))
            if new is None or new > self.ins.limit:
                self.remove(column)
            else:
                self.set(column, "colId", str(new - moved.min_col))

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
        to the row after it, so it moves when that row does and goes when that
        row is deleted.
        """
        block = container.find(_ROW_BREAKS if self.ins.rows else _COL_BREAKS)
        if block is None or self.ins.span is not None:
            return
        removed = False
        for brk in block.iterchildren(_BRK):
            brk_id = _number(brk.get("id", "0"), "brk/@id")
            after = self.ins.move(brk_id + 1)
            if after == brk_id + 1:
                continue
            if after is None or not 1 < after <= self.ins.limit:
                self.remove(brk)
                removed = True
            else:
                self.set(brk, "id", str(after - 1))
        if removed:
            self.after(partial(_refresh_manual_break_count, block))

    def views(self, root: etree._Element) -> None:
        """Selections follow their cells. The scroll positions (``topLeftCell``)
        are places on screen: they move only for an insertion strictly above
        (left of) them, so new rows inserted at the top of a view stay in sight,
        and a span shift leaves them alone."""
        custom = list(root.iterfind(f"{_CUSTOM_SHEET_VIEWS}/{_CUSTOM_SHEET_VIEW}"))
        screen = self.ins.screen
        for view in [*root.iterfind(f"{_SHEET_VIEWS}/{_SHEET_VIEW}"), *custom]:
            top = self._view_start(view.get("topLeftCell"))
            pane = view.find(_PANE)
            if screen is not None:
                self.area_attr(view, "topLeftCell", "topLeftCell", lost="pin", ins=screen)
                if pane is not None:
                    self.pane(pane, top, screen)
            for selection in view.iterchildren(_SELECTION):
                self.selection(selection)
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

    def pane(self, pane: etree._Element, top: int, screen: _Shift) -> None:
        """The frozen split, then the pane's scroll position, which a delete
        leaves no higher (further left) than the first unfrozen row (column)."""
        split = self.frozen_split(pane, top)
        text = pane.get("topLeftCell")
        if not text:
            return
        try:
            area = CellRange.parse(text)
        except InvalidRangeError:
            return  # view state only; never block a shift over it
        pinned = screen.pinned(area)
        if split is not None and not self.ins.inserting:
            first = (self.ins.move(top) or self.ins.at) + int(split)
            if self.ins.rows and pinned.min_row < first:
                pinned = replace(pinned, min_row=first, max_row=max(first, pinned.max_row), text=None)
            elif not self.ins.rows and pinned.min_col < first:
                pinned = replace(pinned, min_col=first, max_col=max(first, pinned.max_col), text=None)
        if pinned is not area:
            self.set(pane, "topLeftCell", str(pinned))

    def frozen_split(self, pane: etree._Element, top: int) -> float | None:
        """Grow a frozen split when rows (columns) are inserted inside the frozen
        band, and shrink it, to one row (column) at least, by those deleted
        from it; return the split, or ``None`` for a pane that is not frozen.

        The split counts rows (columns) from the view's top-left cell, so the
        band is ``top .. top + split - 1``.
        """
        if pane.get("state") not in ("frozen", "frozenSplit"):
            return None  # a plain split is measured in twips, not cells
        attr = "ySplit" if self.ins.rows else "xSplit"
        try:
            split = float(pane.get(attr, "0"))
        except ValueError:
            return None
        if split <= 0:
            return None
        if self.ins.inserting:
            new = split + self.ins.delta if top <= self.ins.at < top + split else split
        else:
            gone = min(self.ins.end, top + split - 1) - max(self.ins.at, top) + 1
            new = max(1.0, split - max(gone, 0))
        if new != split:
            self.set(pane, attr, str(int(new)) if new.is_integer() else repr(new))
        return new

    def columns(self, root: etree._Element) -> None:
        """Move ``<col min max>`` spans like ranges of columns, clamped to column XFD."""
        if self.ins.span is not None:
            return
        for col in root.iterfind(f"{_COLS}/{_COL}"):
            lo = _number(col.get("min"), "col/@min")
            hi = _number(col.get("max"), "col/@max")
            moved = self.ins.interval(lo, hi)
            if moved is None:
                self.remove(col)
            elif moved != (lo, hi):
                self.set(col, "min", str(moved[0]))
                self.set(col, "max", str(moved[1]))

    def row_spans(self, sheet_data: etree._Element | None) -> None:
        """Move each row's ``spans`` hint (columns of its block's cells); drop a
        hint whose columns are all deleted."""
        if sheet_data is None:
            return
        number = 0
        for row in sheet_data.iterchildren(_ROW):
            r = row.get("r")
            number = int(r) if r is not None and r.isdigit() else number + 1
            spans = row.get("spans")
            if not spans or (self.ins.span and not self.ins.span[0] <= number <= self.ins.span[1]):
                continue
            intervals = _parse_spans(spans)
            if intervals is None:
                continue  # an optimisation hint only; leave what we cannot read
            moved = [m for m in (self.ins.interval(lo, hi) for lo, hi in intervals) if m]
            if moved:
                self.set(row, "spans", " ".join(f"{lo}:{hi}" for lo, hi in moved))
            else:
                self.drop(row, "spans")

    def extensions(self, root: etree._Element) -> None:
        """``xm:sqref`` anywhere in the worksheet's extensions, and the ``xm:f``
        formulas of x14 conditional formats, validations and sparklines."""
        ext_lst = root.find(_EXT_LST)
        if ext_lst is None:
            return
        sqrefs: dict[etree._Element, tuple[SheetRangeSet, SheetRangeSet]] = {}
        for sqref in ext_lst.iter(_XM_SQREF):
            ranges = _range_set(sqref.text or "", "xm:sqref")
            moved = self.ins.ranges(ranges)
            sqrefs[sqref.getparent()] = ranges, moved
            if moved is ranges:
                continue
            if moved:
                self.set_text(sqref, str(moved))
            else:
                self.remove(sqref.getparent())
        for f, owner in _extension_formulas(ext_lst):
            if owner.tag == _X14_SPARKLINE_GROUP:  # data ranges, not relative to a cell
                self.set_text(f, self.plain(f.text, own=True))
            else:
                self.set_text(f, self.anchored(f.text, sqrefs=sqrefs.get(owner)))

    def plain(self, text: str, *, own: bool) -> str:
        """``text`` with its references to the shifted sheet moved; ``own`` is
        whether the text belongs to that sheet, so unqualified references name it."""
        return rewrite_formula_refs(text, self.sheet_name, self.ins.area, unqualified=own)

    def anchored(
        self, text: str, *, sqrefs: tuple[SheetRangeSet, SheetRangeSet] | None
    ) -> str:
        """``text``, a formula of a format or validation on the shifted sheet that
        Excel evaluates relative to the top-left cell of its areas ``sqrefs``
        (before, after), moved by :func:`rewrite_anchored_refs`."""
        if sqrefs is None:
            return self.plain(text, own=True)
        return rewrite_anchored_refs(text, self.sheet_name, self.ins, *sqrefs)

    def format_formulas(self, cf: etree._Element, rewrite: Callable[[str], str]) -> None:
        """A conditional format's rule formulas and ``cfvo/@val`` values; an x14
        ``xm:f`` twin repeats a value, and the two must stay equal."""
        for formula in cf.iter(_FORMULA):
            if formula.text:
                self.set_text(formula, rewrite(formula.text))
        for cfvo in cf.iter(_CFVO):
            val = cfvo.get("val")
            if val:
                self.set(cfvo, "val", rewrite(val))

    def validation_formulas(self, dv: etree._Element, rewrite: Callable[[str], str]) -> None:
        for formula in (dv.find(_FORMULA1), dv.find(_FORMULA2)):
            if formula is not None and formula.text:
                self.set_text(formula, rewrite(formula.text))

    def foreign_sheet(self, root: etree._Element) -> None:
        """References to the shifted sheet from another sheet's conditional formats,
        validations (with their x14 twins and sparklines) and internal hyperlinks.
        That sheet's cells stay put, so its formulas need no re-anchoring."""
        rewrite = partial(self.plain, own=False)
        for cf in root.iterchildren(_CF):
            self.format_formulas(cf, rewrite)
        for dv in root.iterfind(f"{_DATA_VALIDATIONS}/{_DATA_VALIDATION}"):
            self.validation_formulas(dv, rewrite)
        ext_lst = root.find(_EXT_LST)
        if ext_lst is not None:
            for f, _owner in _extension_formulas(ext_lst):
                self.set_text(f, rewrite(f.text))
        self.hyperlink_locations(root, own=False)

    def hyperlink_locations(self, root: etree._Element, *, own: bool) -> None:
        """Internal hyperlinks only: with ``r:id``, ``location`` is inside another document."""
        for link in root.iterfind(f"{_HYPERLINKS}/{_HYPERLINK}"):
            location = link.get("location")
            if location and link.get(f"{{{OFFICE_REL_NS}}}id") is None:
                new = rewrite_plain_ref(location, self.sheet_name, self.ins.area, unqualified=own)
                self.set(link, "location", new)

    def charts(self, workbook) -> None:
        """References to the shifted sheet in every chart: series, titles and
        labels with their c15 copies. A chart's references name their sheet."""
        for part in chart_parts(workbook._package):
            for elm in chart_references(part.element):
                self.set_text(elm, self.plain(elm.text, own=False))

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


def _extension_formulas(
    ext_lst: etree._Element,
) -> list[tuple[etree._Element, etree._Element]]:
    """Each ``xm:f`` with text under an x14 conditional format, validation or
    sparkline group, with that owner."""
    found = []
    for f in ext_lst.iter(_XM_F):
        owner = next((a for a in f.iterancestors() if a.tag in _X14_FORMULA_OWNERS), None)
        if f.text and owner is not None:
            found.append((f, owner))
    return found


def _parse_spans(spans: str) -> list[tuple[int, int]] | None:
    intervals = []
    for item in spans.split():
        lo, sep, hi = item.partition(":")
        if not (sep and lo.isdigit() and hi.isdigit()):
            return None
        intervals.append((int(lo), int(hi)))
    return intervals


_DEFERRED_MARKERS = (_XM_NS.encode(), b"location=", b"formula", b"cfvo")


def _root_if_it_may_refer(worksheet) -> etree._Element | None:
    """The worksheet's element, unless it is still unparsed and cannot hold a
    reference this module rewrites (keeps ``Workbook.open(large=True)`` lazy),
    or is a chartsheet, which holds none."""
    if not isinstance(worksheet._part, WorksheetPart):
        return None
    blob = worksheet._part.unparsed_blob
    if blob is not None and not any(marker in blob for marker in _DEFERRED_MARKERS):
        return None
    return worksheet._part.element


class InsertPlan:
    """Every change that inserting rows or columns into one worksheet makes.

    Building the plan reads the worksheet, its table parts, the workbook's
    defined names and charts and the other sheets' references to it, and raises
    ``InvalidRangeError``, ``GridOverflowError``, ``TableError`` or ``ValueError``
    before anything changes.
    ``apply`` then only assigns precomputed values. ``merges`` lists the
    sheet's merge ranges as they will be after ``apply``; ``header_cells``
    the header cells of new table columns to write after it (see
    ``_check_tables``, which reads ``cells``).
    """

    def __init__(self, worksheet, ins: _Shift, cells: tuple[list, int] | None = None):
        if not ins.inserting or ins.span is not None:
            raise ValueError("InsertPlan carries out inserts of whole rows or columns only")
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
            if other_root is not None:
                planner.foreign_sheet(other_root)
        self._check_tables(cells, planner.merges)
        if ins.rows:
            planner.after(partial(shift_table_parts, worksheet, ins.at, ins.delta))
        else:
            names = self._column_names
            planner.after(partial(shift_table_parts_cols, worksheet, ins.at - 1, ins.delta, names))
        planner.defined_names(workbook)
        planner.charts(workbook)
        planner.after(workbook._invalidate_calc_chain)
        planner.after(workbook.set_full_calc_on_load)
        self._planner = planner
        self.merges = planner.merges

    def _check_tables(self, cells: tuple[list, int] | None, merges: list[CellRange]) -> None:
        """Refuse an insertion that would push any part of a table off the grid
        (its columns and rows cannot shrink with it), or that would add columns
        inside a table bound to a query; name the columns that join a table.

        ``cells`` is ``(rows, first_row)``, the values ``insert_columns``
        writes into the new columns from ``first_row`` down. The value it
        writes into a joining column's header cell names the column when it is
        a non-empty string no other column of the table uses (ignoring case);
        where it writes nothing, the column gets the first unused ``ColumnN``,
        which ``header_cells`` lists for the caller to write into that cell;
        anything else raises ``TableError``. So does a value for a header cell
        that one of ``merges`` (the sheet's merges after the insert) covers,
        since the bulk write would send it to the merge's top-left cell.
        """
        ins = self._insert
        self._column_names: dict[object, list[str]] = {}
        self.header_cells: dict[str, str] = {}

        def check(ref: str, what: str) -> None:
            if ins.clamps(_area(ref, what)):
                raise ins.overflow(f"{what} ({ref})")

        for table in self._worksheet.tables:
            table._check_move(check)
            if ins.rows or _inner_column_offset(table.ref, ins.at - 1) is None:
                continue
            _check_table_accepts_columns(table._element, ins.at - 1)
            header_row = table._box()[1] if table.header_row_count else None
            supplied: list = [None] * ins.delta
            if header_row is not None and cells is not None:
                rows, first_row = cells
                if 0 <= header_row - first_row < len(rows):
                    supplied = list(rows[header_row - first_row])
            used = {(name or "").casefold() for name in table._column_names()}
            for offset, value in enumerate(supplied):
                if value is None:
                    continue
                address = join_address(index_to_col(ins.at - 1 + offset), header_row)
                covering = next(
                    (m for m in merges if m.contains(ins.at + offset, header_row)), None
                )
                if covering is not None:
                    raise TableError(
                        f"the header cell {address} of a column joining table "
                        f"{table.display_name!r} lies inside the merge {covering}, so the "
                        f"value {value!r} would go to its top-left cell and cannot name "
                        "the column; write nothing there to get a generated ColumnN name"
                    )
                if isinstance(value, str) and value and value.casefold() not in used:
                    used.add(value.casefold())
                    continue
                raise TableError(
                    f"the header cell {address} of a column joining table "
                    f"{table.display_name!r} would hold {value!r}; a new table column's "
                    "header must be a non-empty string no other column of the table uses "
                    "(ignoring case), or nothing for a generated ColumnN name"
                )
            names = _joining_column_names(supplied, used)
            self._column_names[table._part] = names
            if header_row is not None:
                for offset, (value, name) in enumerate(zip(supplied, names)):
                    if value is None:
                        address = join_address(index_to_col(ins.at - 1 + offset), header_row)
                        self.header_cells[address] = name
        if self.header_cells and self._worksheet.shared_strings is None:
            raise ValueError("workbook has no shared strings part")

    def apply(self) -> None:
        _walk_sheet_data(self._worksheet._part.element.find(_SHEET_DATA), self._insert, apply=True)
        self._planner.apply()


def plan_insert_rows(worksheet, at_row: int, count: int) -> InsertPlan:
    """Plan inserting ``count`` rows before row ``at_row`` (1-based)."""
    return InsertPlan(worksheet, _Shift(True, at_row, count))


def plan_insert_columns(
    worksheet, at_col: int, count: int, cells: tuple[list, int] | None = None
) -> InsertPlan:
    """Plan inserting ``count`` columns before 0-based column ``at_col``.

    ``cells`` is ``(rows, first_row)``: the values written into the new
    columns, which name the columns that join a table (see ``InsertPlan``).
    """
    return InsertPlan(worksheet, _Shift(False, at_col + 1, count), cells)


class CellShiftPlan:
    """Every change that shifting the cells of a span makes, as Excel's "insert
    cells" and "delete cells" do: ``shift`` has a ``span``.

    Only the ``<c>`` elements in the span's columns (rows) from the shifted
    lines on move: down (right) for an insert, up (left) for a delete, which
    drops the cells it deletes. A cell moving down goes into the row it lands
    on, created when missing; rows keep their height and style where they are,
    and a row left with no cells and no attributes of its own goes, as in
    Excel. The sheet's other structures move as :class:`_Planner` moves them
    for a span, and its tables as Excel moves them (``_check_span_tables``).
    Cell formulas, comments and drawing anchors stay, as for an
    :class:`InsertPlan`.

    Building the plan reads what an ``InsertPlan`` reads and the sheet's
    drawings, and raises before anything changes: ``GridOverflowError`` for
    content or a drawing object pushed past the grid; ``TableError`` for a
    table the shift would move only part of, and the table changes
    ``_check_span_tables`` refuses; ``InvalidRangeError`` for a merge it
    would move only part of (Excel unmerges it after asking), a scenario or
    data-table input cell it would delete, or a reference on the sheet it
    cannot read; ``FormulaGroupError`` for a shared, array or data-table
    formula range it would change only part of; and ``ValueError`` for a
    shift without a span, as one whose span covers every column (row) is.
    ``apply`` carries the plan out and refreshes the worksheet's caches and
    ``<dimension>``.
    """

    def __init__(self, worksheet, shift: _Shift):
        if shift.span is None:
            raise ValueError("CellShiftPlan shifts the cells of a span; whole lines need an InsertPlan")
        self._worksheet = worksheet
        root = worksheet._part.element
        sheet_data = root.find(_SHEET_DATA)
        self._cells = _CellMoves(sheet_data, shift)
        planner = _Planner(shift, sheet_name=worksheet.name, part=worksheet._part)
        _plan_formula_ranges(sheet_data, planner)
        planner.worksheet(root)
        planner.hyperlink_locations(root, own=True)
        workbook = worksheet._workbook
        for other in workbook.worksheets:
            if other is worksheet:
                continue
            other_root = _root_if_it_may_refer(other)
            if other_root is not None:
                planner.foreign_sheet(other_root)
        tables = _check_span_tables(worksheet, shift)
        planner.after(partial(_move_span_tables, tables, shift))
        _check_span_drawings(worksheet, shift)
        planner.defined_names(workbook)
        planner.charts(workbook)
        planner.after(workbook._invalidate_calc_chain)
        planner.after(workbook.set_full_calc_on_load)
        self._planner = planner
        self.merges = planner.merges

    def apply(self) -> None:
        worksheet = self._worksheet
        if self._cells.apply():
            worksheet._part.mark_dirty()
        self._planner.apply()
        worksheet._invalidate_merge_map()
        worksheet._invalidate_bulk_indexes()
        worksheet.update_dimension()


def plan_shift_cells(worksheet, shift: _Shift) -> CellShiftPlan:
    """Plan shifting the cells of ``shift.span`` (see :class:`CellShiftPlan`)."""
    return CellShiftPlan(worksheet, shift)


def shift_cells(worksheet, shift: _Shift) -> None:
    """Shift the cells of ``shift.span`` as Excel's "insert cells" or "delete
    cells" does (see :class:`CellShiftPlan`); nothing changes when it raises."""
    CellShiftPlan(worksheet, shift).apply()


_X14AC_DY_DESCENT = "{http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac}dyDescent"
# attributes that say nothing about a row without cells, which Excel then leaves out
_PLAIN_ROW_ATTRIBUTES = frozenset({"r", "spans", _X14AC_DY_DESCENT})


class _CellMoves:
    """The ``<c>`` elements a span shift moves, checked when built.

    Building it raises ``InvalidRangeError`` for an unreadable row or cell
    address anywhere on the sheet, so that refreshing ``<dimension>`` after
    the move cannot fail, and ``GridOverflowError`` for content an insert
    would push past the grid; a cell pushed off that holds no content goes,
    as in Excel.
    """

    def __init__(self, sheet_data: etree._Element | None, shift: _Shift):
        self._sheet_data = sheet_data
        self._shift = shift
        self._implied: list[tuple[etree._Element, str]] = []  # r values only implied
        self._moves: list[tuple[etree._Element, int, int]] = []  # cell, new row, new column
        self._drops: list[etree._Element] = []
        if sheet_data is None:
            return
        lo_s, hi_s = shift.span
        number = 0
        for row in sheet_data.iterchildren(_ROW):
            r = row.get("r")
            number = _number(r, "row/@r") if r is not None else number + 1
            if r is None:  # removing a row before it would renumber it
                self._implied.append((row, str(number)))
            moving = number >= shift.at if shift.rows else lo_s <= number <= hi_s
            col = 0
            for c in row.iterchildren(_C):
                address = c.get("r")
                col = _cell_column(address, "c/@r") if address is not None else col + 1
                if not moving:
                    continue
                if address is None:
                    self._implied.append((c, f"{index_to_col(col - 1)}{number}"))
                if not (lo_s <= col <= hi_s if shift.rows else col >= shift.at):
                    continue
                new = shift.move(number if shift.rows else col)
                if new is not None and new > shift.limit:
                    if _holds_content(c):
                        name = address or f"{index_to_col(col - 1)}{number}"
                        raise shift.overflow(f"cell {name}", excel=_EXCEL_GRID)
                    new = None
                if new is None:
                    self._drops.append(c)
                else:
                    self._moves.append((c, new, col) if shift.rows else (c, number, new))

    def apply(self) -> bool:
        """Move the cells; return whether any moved or went."""
        if not (self._moves or self._drops):
            return False
        for elm, value in self._implied:
            elm.set("r", value)
        left = {c.getparent() for c in self._drops}
        for c in self._drops:
            c.getparent().remove(c)
        if self._shift.rows:
            left.update(c.getparent() for c, _, _ in self._moves)
            self._move_between_rows()
        else:
            for c, row, col in self._moves:
                c.set("r", f"{index_to_col(col - 1)}{row}")
        for row in left:
            empty = not any(isinstance(child.tag, str) for child in row)
            if empty and _PLAIN_ROW_ATTRIBUTES.issuperset(row.attrib.keys()):
                row.getparent().remove(row)
        return True

    def _move_between_rows(self) -> None:
        """Put each moving cell into its new row, after the row's cells left of the span."""
        for c, _, _ in self._moves:
            c.getparent().remove(c)
        rows: dict[int, etree._Element] = {}
        for row in self._sheet_data.iterchildren(_ROW):
            rows.setdefault(int(row.get("r")), row)
        numbers = sorted(rows)
        incoming: dict[int, list[tuple[int, etree._Element]]] = {}
        for c, number, col in self._moves:
            incoming.setdefault(number, []).append((col, c))
        last_col = self._shift.span[1]
        # new rows come in ascending order, so each goes before the first existing row after it
        for number in sorted(incoming):
            row = rows.get(number)
            if row is None:
                row = etree.Element(_ROW)
                row.set("r", str(number))
                index = bisect.bisect_right(numbers, number)
                if index < len(numbers):
                    rows[numbers[index]].addprevious(row)
                else:
                    self._sheet_data.append(row)
                rows[number] = row
            following = next(
                (child for child in row if child.tag != _C or _column_key(child) >= last_col), None
            )
            for col, c in sorted(incoming[number], key=operator.itemgetter(0)):
                c.set("r", f"{index_to_col(col - 1)}{number}")
                if following is None:
                    row.append(c)
                else:
                    following.addprevious(c)
            _widen_spans(row)


def _plan_formula_ranges(sheet_data: etree._Element | None, planner: _Planner) -> None:
    """Move the ranges of shared, array and data-table formulas for a span shift;
    raise ``FormulaGroupError`` for a range the shift would change only part of."""
    if sheet_data is None:
        return
    shift = planner.ins
    for f in sheet_data.iter(_F):
        kind = f.get("t")
        if kind not in _RANGED_FORMULAS:
            continue
        cell = f.getparent().get("r") or ""
        for attr in ("ref", "r1", "r2") if kind == "dataTable" else ("ref",):
            value = f.get(attr)
            if not value:
                continue
            area = _area(value, f"{cell} <f t={kind!r}> @{attr}")
            if attr == "ref" and shift.tears(_formula_group(kind, area)):
                raise _torn_formula_error(shift, kind, value, cell)
            moved = shift.area(area)
            if moved is None and attr == "ref" and not shift.inserting:
                break  # deleted with its cells, input cells and all
            if moved is None or shift.clamps(area):
                if attr == "ref":
                    raise shift.refusal(f"the {kind} formula range {value} of {cell}")
                raise shift.refusal(f"the data table input cell {attr}={value!r} of {cell}")
            if moved is not area:
                planner.set(f, attr, str(moved))


def _formula_group(kind: str, ref: CellRange) -> CellRange:
    """The cells that belong to a ranged formula: its ``ref``, and for a data
    table also the row of formulas above it and the column of inputs left of
    it, which Excel keeps together with the results."""
    if kind != "dataTable":
        return ref
    return CellRange.from_bounds(max(ref.min_col - 1, 1), max(ref.min_row - 1, 1), ref.max_col, ref.max_row)


def _torn_formula_error(shift: _Shift, kind: str, ref: str, cell: str) -> FormulaGroupError:
    message = f"{shift.describe()} would change only part of the {kind} formula range {ref} of {cell}"
    if kind == "shared":
        return FormulaGroupError(
            f"{message}; Excel turns the cells it moves into separate formulas, which xlsxedit does not do"
        )
    return FormulaGroupError(_excel_says(message, _EXCEL_ARRAY if kind == "array" else _EXCEL_DATA_TABLE))


def _check_span_tables(worksheet, shift: _Shift) -> list:
    """Refuse a span shift that would move only part of a table, as Excel does;
    return the tables it moves.

    A table the shift reaches, past the shifted lines and across the span,
    must lie in the span. One wholly past the shifted lines moves; one they
    cut through must fill the span exactly, and then grows or shrinks by
    rows, as in Excel. xlsxedit refuses, where Excel would not, to add or
    delete table columns that way, or to delete a table's header row, totals
    row or every data row.
    """
    reached = []
    for table in worksheet.tables:
        box, header, totals = table._layout()
        name = table.display_name or table.name or str(table._part.partname)
        what = f"table {name!r} ({table.ref})"
        area = CellRange.from_bounds(box[0] + 1, box[1], box[2] + 1, box[3])
        lo, hi, lo_x, hi_x = shift._extents(area)
        if hi < shift.at or hi_x < shift.span[0] or shift.span[1] < lo_x:
            continue  # before the shifted lines, or beside the span
        cut = lo < shift.at if shift.inserting else lo <= shift.end
        if not shift.span[0] <= lo_x <= hi_x <= shift.span[1] or cut and (lo_x, hi_x) != shift.span:
            raise TableError(_excel_says(f"{shift.describe()} would move only part of {what}", _EXCEL_TABLE))
        if cut and not shift.rows:
            change = "add columns to" if shift.inserting else "delete columns of"
            raise TableError(f"{shift.describe()} would {change} {what}, which xlsxedit does not do")
        if cut and not shift.inserting:
            first, last = box[1] + header, box[3] - totals
            lost = None
            if header and shift.at < first:
                lost = "the header row"
            elif totals and shift.end > last:
                lost = "the totals row"
            elif shift.at <= first and shift.end >= last:
                lost = "every data row"
            if lost is not None:
                raise TableError(
                    f"{shift.describe()} would delete {lost} of {what}, which xlsxedit does not do"
                )

        def check(ref: str, label: str) -> None:
            area = _area(ref, label)
            if shift.area(area) is None or shift.clamps(area):
                raise shift.refusal(f"{label} ({ref})")

        table._check_move(check)
        reached.append(table)
    return reached


def _move_span_tables(tables: list, shift: _Shift) -> None:
    """Move each of ``tables``, with its ``autoFilter`` and sort state (see
    ``Table._move``), as ``_check_span_tables`` allowed."""
    for table in tables:
        table._move(lambda ref: str(shift.area(CellRange.parse(ref))))


def _anchor_extent(anchor: etree._Element) -> tuple[CellRange, CellRange] | None:
    """The cells a drawing anchor's ``xdr:from`` and ``xdr:to`` markers name, and
    the cells it covers, which leave out a last column (row) that ``xdr:to``
    reaches only at offset 0. A one-cell anchor gives its ``xdr:from`` cell,
    leaving its size (``xdr:ext``) out, since that needs the sheet's column
    widths and row heights; ``None`` for an absolute anchor or markers that
    cannot be read."""
    first = anchor.find(_XDR_FROM)
    if first is None:
        return None
    last = anchor.find(_XDR_TO)
    try:
        col, _, row, _ = _marker(first)
        to_col, col_off, to_row, row_off = (col, 1, row, 1) if last is None else _marker(last)
        markers = CellRange.from_bounds(col + 1, row + 1, to_col + 1, to_row + 1)
    except (ValueError, InvalidRangeError):
        return None
    cover_col = to_col if col_off or to_col == col else to_col - 1
    cover_row = to_row if row_off or to_row == row else to_row - 1
    return markers, CellRange.from_bounds(col + 1, row + 1, cover_col + 1, cover_row + 1)


def _check_span_drawings(worksheet, shift: _Shift) -> None:
    """Refuse an insert that would push a drawing object it :meth:`_Shift.carries`
    past the grid, whatever the object's placement, as Excel does."""
    if not shift.inserting:
        return
    for part in drawing_parts_for_worksheet(worksheet):
        for anchor, _alternate in _iter_anchors(part.element, None):
            extent = _anchor_extent(anchor)
            if extent is None:
                continue
            markers, cells = extent
            end = markers.max_row if shift.rows else markers.max_col
            if shift.carries(cells) and end + shift.delta > shift.limit:
                raise shift.overflow(f"the drawing object at {markers}", excel=_EXCEL_OBJECTS)


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
            if isinstance(value, int) and value.bit_length() > _MAX_INT_BITS:
                try:  # str() refuses more digits than sys.get_int_max_str_digits()
                    str(value)
                except ValueError as exc:
                    raise ValueError(f"{method} cannot write an integer that long: {exc}") from None
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


def check_merge_anchors_writable(
    worksheet,
    merges: list[CellRange],
    rows: list,
    *,
    start_row: int,
    start_col_idx: int,
    new_rows: bool,
) -> None:
    """Refuse, before anything changes, a value the bulk writer would send to a
    merge's top-left cell that anchors a multi-cell array or data-table formula.

    ``rows`` are the values for the cells of the new rows (``new_rows``) or
    columns, from ``start_row`` / 0-based ``start_col_idx``; ``merges`` are the
    sheet's merges as they will be once those cells exist. A merge whose
    top-left cell is not new reaches the new cells by straddling the insertion
    point, so that cell lies before the point and holds now what it will hold
    then.
    """
    last_row = start_row + len(rows) - 1
    for area in merges:
        if area.min_row >= start_row if new_rows else area.min_col > start_col_idx:
            continue  # the top-left cell is new as well: this call writes it, or it is empty
        first, last = max(area.min_row, start_row), min(area.max_row, last_row)
        lo, hi = max(area.min_col - 1 - start_col_idx, 0), area.max_col - start_col_idx
        if first > last or lo >= hi:
            continue
        covered = (rows[row - start_row] for row in range(first, last + 1))
        if not any(
            value is not None
            for values in covered
            for value in (values if isinstance(values, (list, tuple)) else (values,))[lo:hi]
        ):
            continue
        anchor = f"{index_to_col(area.min_col - 1)}{area.min_row}"
        c = worksheet._find_cell_element(anchor)
        f = None if c is None else c.find(_F)
        if f is not None:
            _refuse_range_formula_edit(anchor, f)


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
        r = row.get("r")
        if r is not None and r.isdigit() and first_row <= int(r) <= last_row:
            _widen_spans(row)


def _widen_spans(row: etree._Element) -> None:
    """Grow the row's ``spans`` hint, if it has one, to cover its cells."""
    spans = row.get("spans")
    if not spans:
        return
    intervals = _parse_spans(spans)
    cols = [_cell_column(c.get("r"), "c/@r") for c in row.iterchildren(_C) if c.get("r")]
    if intervals is None or not cols:
        return
    if all(any(lo <= col <= hi for lo, hi in intervals) for col in cols):
        return
    lo = min(min(cols), *(lo for lo, _ in intervals))
    hi = max(max(cols), *(hi for _, hi in intervals))
    row.set("spans", f"{lo}:{hi}")


def shift_range_ref(ref: str, at_row: int, delta: int) -> str:
    """Shift a range at or below ``at_row``; a straddling range grows. Clamped to the grid."""
    moved = CellRange.parse(ref).shift_rows(at_row, delta)
    if moved is None:
        raise _Shift(True, at_row, delta).overflow(repr(ref))
    return str(moved)


def shift_range_ref_cols(ref: str, at_col: int, delta: int) -> str:
    """Shift column letters in a range at or to the right of ``at_col`` (0-based)."""
    moved = CellRange.parse(ref).shift_cols(at_col + 1, delta)
    if moved is None:
        raise _Shift(False, at_col + 1, delta).overflow(repr(ref))
    return str(moved)


def shift_table_parts(worksheet, at_row: int, delta: int) -> None:
    """Shift each table of the worksheet, with its ``autoFilter`` and sort state,
    after row insertion (see ``Table._move``)."""
    for table in worksheet.tables:
        table._move(lambda ref: shift_range_ref(ref, at_row, delta))


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


def _joining_column_names(supplied: list, used: set[str]) -> list[str]:
    """Names for columns joining a table: each supplied name, and the first
    unused ``ColumnN`` (``used`` holds casefolded names) where ``None``."""
    names = []
    n = 0
    for value in supplied:
        if value is not None:
            names.append(value)
            continue
        n += 1
        while f"column{n}" in used:
            n += 1
        names.append(f"Column{n}")
    return names


def _insert_table_columns(table_elm: etree._Element, offset: int, names: list[str]) -> None:
    """Add a ``tableColumn`` per name in ``names`` before the one at ``offset``.

    Each gets the next free ``id``; ``autoFilter`` ``filterColumn`` ids at or
    after ``offset`` move with their columns.
    """
    cols_elm = table_elm.find(_TABLE_COLUMNS)
    if cols_elm is None:
        return
    delta = len(names)
    existing = cols_elm.findall(_TABLE_COLUMN)
    next_id = max((int(c.get("id")) for c in existing if c.get("id", "").isdigit()), default=0)
    anchor = existing[offset] if offset < len(existing) else None
    for name in names:
        next_id += 1
        column = etree.Element(_TABLE_COLUMN)
        column.set("id", str(next_id))
        column.set("name", name)
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


def shift_table_parts_cols(
    worksheet, at_col: int, delta: int, names: dict[object, list[str]]
) -> None:
    """Shift each table of the worksheet, with its ``autoFilter`` and sort state,
    after column insertion (see ``Table._move``).

    Columns inserted strictly inside a table join it as new ``tableColumn``s
    (see ``_insert_table_columns``), named from ``names``, which the plan
    chose for each such table part. Raises ``TableError``, before changing any table
    part, when such a table is bound to a query: its query table part would
    not know the new columns. ``plan_insert_columns`` makes the same check
    before anything on the sheet moves.
    """
    tables = [(t, _inner_column_offset(t.ref, at_col)) for t in worksheet.tables]
    for table, _offset in tables:
        _check_table_accepts_columns(table._element, at_col)
    for table, offset in tables:
        if offset is not None:
            _insert_table_columns(table._element, offset, names[table._part])
            table._part.mark_dirty()
        table._move(lambda ref: shift_range_ref_cols(ref, at_col, delta))


def shift_dollar_ref(ref: str, at_row: int, delta: int) -> str:
    """Shift rows in an A1 ref, preserving ``$``; an unparseable ref is returned as is."""
    return _move_ref_text(ref, lambda area: area.shift_rows(at_row, delta))

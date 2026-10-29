"""Move a worksheet's drawing anchors with a row or column shift, as Excel does.

Excel writes each placement of a drawing object as ``editAs`` on an
``xdr:twoCellAnchor``: ``twoCell`` (the default) moves and sizes with cells,
``oneCell`` moves but keeps its size, ``absolute`` neither moves nor sizes.
``xdr:oneCellAnchor`` (a corner and an extent) moves; ``xdr:absoluteAnchor``
never changes. ``docs/features.md`` ("Drawings on inserts") states the rules,
which follow Excel's recordings in ``fixtures/excel-reference/``.

A shift is a :class:`xlsxedit.row_shift._Shift`. A ``oneCell`` or
``absolute`` anchor keeps its size or place in EMU, measured over the sheet's
row heights and column widths before the shift and placed over those after
it, which the shift implies: lines keep their size as they move, and inserted
lines take what the moved ``<col>`` spans (columns) or the default height
(rows) give them. :class:`DrawingShift` computes every change, raising before
anything changes, and ``apply`` only assigns the result.
"""

from __future__ import annotations

from functools import cached_property
from typing import TYPE_CHECKING

from lxml import etree

from xlsxedit.drawing import (
    _XDR_ABSOLUTE,
    _XDR_COL,
    _XDR_COLOFF,
    _XDR_FROM,
    _XDR_ONE_CELL,
    _XDR_ROW,
    _XDR_ROWOFF,
    _XDR_TO,
    _Axis,
    _iter_anchors,
    _iter_cnvpr,
    _marker,
    drawing_parts_for_worksheet,
    sheet_axis,
)
from xlsxedit.exceptions import InvalidRangeError
from xlsxedit.opc.constants import OFFICE_REL_NS

if TYPE_CHECKING:
    from xlsxedit.row_shift import _Shift

_Marker = tuple[int, int]  # (line, offset in EMU) on the shifted axis


def anchor_in_span(shift: _Shift, col: int, row: int) -> bool:
    """Whether ``shift`` moves an anchor whose top-left cell is 0-based
    ``(col, row)``: a shift restricted to a span of columns (rows) moves only
    the anchors whose column (row) lies inside it. A whole-row or whole-column
    shift moves every anchor."""
    if shift.span is None:
        return True
    lo, hi = shift.span
    return lo <= (col if shift.rows else row) + 1 <= hi


def _shifted_runs(axis: _Axis, shift: _Shift) -> list[tuple[int, int, int]]:
    """The runs of ``axis`` after a whole-line ``shift``, as the sheet will have
    them: rows move with their heights and inserted rows have none; ``<col>``
    spans move as ``_Planner.columns`` moves them, so a span that straddles an
    insert also covers the inserted columns."""
    at, end = shift.at - 1, shift.end - 1
    runs = []
    for lo, hi, size in axis.runs:
        if not shift.rows:
            moved = shift.interval(lo + 1, hi + 1)
            if moved is not None:
                runs.append((moved[0] - 1, moved[1] - 1, size))
            continue
        after = at if shift.inserting else end + 1
        pieces = ((lo, min(hi, at - 1)), (max(lo, after) + shift.delta, hi + shift.delta))
        runs += [(a, min(b, axis.count - 1), size) for a, b in pieces if a <= min(b, axis.count - 1)]
    return runs


class DrawingShift:
    """The changes one shift makes to the anchors of a worksheet's drawings.

    Building it reads every anchor of the sheet's drawing parts, both branches
    of ``mc:AlternateContent`` included, and raises ``GridOverflowError`` when
    an anchor would leave the grid and ``InvalidRangeError`` for an anchor it
    cannot read. ``apply`` then sets the planned markers, removes the anchors a
    delete takes away with the relationships only they used, and marks the
    changed drawing parts.
    """

    def __init__(self, worksheet, shift: _Shift):
        self._worksheet = worksheet
        self._shift = shift
        self._at = shift.at - 1  # 0-based first line of the band
        self._end = shift.end - 1  # 0-based last line of the band
        self._sets: list[tuple[etree._Element, str]] = []
        self._removals: list[etree._Element] = []
        self._released: list[tuple[object, str]] = []
        self._changed: list = []
        for part in drawing_parts_for_worksheet(worksheet):
            self._plan_part(part)

    def apply(self) -> None:
        for elm, text in self._sets:
            elm.text = text
        for elm in self._removals:
            elm.getparent().remove(elm)
        for part, r_id in self._released:
            part.rels._rels.pop(r_id, None)
        for part in self._changed:
            part.mark_dirty()

    # -- geometry -----------------------------------------------------------

    @cached_property
    def _before(self) -> _Axis:
        return sheet_axis(self._worksheet, rows=self._shift.rows)

    @cached_property
    def _after(self) -> _Axis:
        """The shifted axis; a shift restricted to a span moves no whole line."""
        if self._shift.span is not None:
            return self._before
        before = self._before
        return _Axis(_shifted_runs(before, self._shift), before.default, before.count)

    def _place(self, position: int, what: str) -> _Marker:
        found = self._after.locate(position)
        if found is None:
            raise self._shift.overflow(what)
        return found

    def _in_band(self, line: int) -> bool:
        return not self._shift.inserting and self._at <= line <= self._end

    @staticmethod
    def _edge(marker: _Marker, top: int | None) -> int:
        """The line a corner lies in. The bottom (right) corner of an object whose
        top (left) corner is in line ``top`` lies, at offset 0, on the gridline
        above its line: lines inserted on that gridline leave it, and deleting
        the line above it reaches it, as Excel does. ``top`` is ``None`` for a
        top corner."""
        line, offset = marker
        return line - 1 if top is not None and offset == 0 and line > top else line

    def _move(self, marker: _Marker, *, top: int | None, what: str) -> _Marker:
        """Where a cell-relative corner goes (``top`` as for :meth:`_edge`); one in
        deleted lines goes to the top of the line after them."""
        edge = self._edge(marker, top)
        if edge < self._at:
            return marker
        if self._in_band(edge):
            return self._at, 0
        return self._shifted(marker, what)

    def _shifted(self, marker: _Marker, what: str) -> _Marker:
        line = marker[0] + self._shift.delta
        if line >= self._shift.limit:
            raise self._shift.overflow(what)
        return line, marker[1]

    # -- anchors ------------------------------------------------------------

    def _plan_part(self, part) -> None:
        removed: list[etree._Element] = []
        changed = False
        for top in part.element:
            anchors = [anchor for anchor, _ in _iter_anchors([top], None)]
            outcomes = [self._plan_anchor(anchor) for anchor in anchors]
            if "delete" in outcomes:
                removed.append(top)
            changed |= any(outcomes)
        if removed:
            self._removals += removed
            self._released += [(part, r_id) for r_id in _orphaned_rels(part.element, removed)]
        if changed:
            self._changed.append(part)

    def _plan_anchor(self, anchor: etree._Element) -> str | bool:
        """Plan the move of one anchor: ``"delete"`` when a delete takes it away,
        else whether it changes."""
        from_elm = anchor.find(_XDR_FROM)
        if anchor.tag == _XDR_ABSOLUTE or from_elm is None:
            return False
        what = f"drawing object {_anchor_name(anchor)!r}"
        start = _read_marker(from_elm, what)
        if not anchor_in_span(self._shift, start[0], start[2]):
            return False
        top = self._axis_part(start)
        if anchor.tag == _XDR_ONE_CELL:
            return self._plan_corner(from_elm, top, self._move(top, top=None, what=what), what)
        to_elm = anchor.find(_XDR_TO)
        if to_elm is None:
            return False
        bottom = self._axis_part(_read_marker(to_elm, what))
        mode = anchor.get("editAs")
        if mode == "absolute":
            new_top, new_bottom = self._keep_place(top, what), self._keep_place(bottom, what)
        elif mode == "oneCell":
            new_top = self._move(top, top=None, what=what)
            new_bottom = self._keep_size(top, bottom, new_top, what)
        else:
            if self._deletes(top, bottom):
                return "delete"
            new_top = self._move(top, top=None, what=what)
            new_bottom = self._move(bottom, top=top[0], what=what)
        changed = self._plan_corner(from_elm, top, new_top, what)
        return self._plan_corner(to_elm, bottom, new_bottom, what) or changed

    def _deletes(self, top: _Marker, bottom: _Marker) -> bool:
        """Whether a delete takes away a two-cell anchor: both its corners lie in
        the deleted lines (see :meth:`_edge`)."""
        return self._in_band(top[0]) and self._in_band(self._edge(bottom, top[0]))

    def _keep_place(self, marker: _Marker, what: str) -> _Marker:
        """An ``absolute`` corner: the same position on the page, over the lines as
        they are after the shift. A corner above the band does not move."""
        if marker[0] < self._at or self._shift.span is not None:
            return marker
        return self._place(self._before.start(marker[0]) + marker[1], what)

    def _keep_size(self, top: _Marker, bottom: _Marker, new_top: _Marker, what: str) -> _Marker:
        """A ``oneCell`` bottom corner: the same extent from the top corner's new place."""
        if self._edge(bottom, top[0]) < self._at:
            return bottom
        if self._shift.span is None and new_top == (top[0] + self._shift.delta, top[1]):
            return self._shifted(bottom, what)  # moved with its lines, sizes and all
        before = self._before
        extent = before.start(bottom[0]) + bottom[1] - before.start(top[0]) - top[1]
        return self._place(self._after.start(new_top[0]) + new_top[1] + extent, what)

    def _axis_part(self, marker: tuple[int, int, int, int]) -> _Marker:
        col, col_off, row, row_off = marker
        return (row, row_off) if self._shift.rows else (col, col_off)

    def _plan_corner(self, marker_elm: etree._Element, old: _Marker, new: _Marker, what: str) -> bool:
        if new == old:
            return False
        line_tag, offset_tag = (_XDR_ROW, _XDR_ROWOFF) if self._shift.rows else (_XDR_COL, _XDR_COLOFF)
        for tag, value, before in ((line_tag, new[0], old[0]), (offset_tag, new[1], old[1])):
            child = marker_elm.find(tag)
            if child is None:
                raise InvalidRangeError(
                    f"{what}: xdr:{etree.QName(marker_elm).localname} has no "
                    f"xdr:{etree.QName(tag).localname}"
                )
            if value != before:
                self._sets.append((child, str(value)))
        return True


def _read_marker(marker_elm: etree._Element, what: str) -> tuple[int, int, int, int]:
    try:
        marker = _marker(marker_elm)
    except ValueError:
        marker = None
    if marker is None or marker[0] < 0 or marker[2] < 0:
        raise InvalidRangeError(
            f"{what}: xdr:{etree.QName(marker_elm).localname} does not name a cell"
        )
    return marker


def _anchor_name(anchor: etree._Element) -> str | None:
    return next((cnv.get("name") for cnv in _iter_cnvpr(anchor)), None)


def _rel_ids(elements) -> set[str]:
    return {
        value
        for root in elements
        for elm in root.iter()
        for name, value in elm.attrib.items()
        if name.startswith(f"{{{OFFICE_REL_NS}}}")
    }


def _orphaned_rels(root: etree._Element, removed: list[etree._Element]) -> list[str]:
    """The relationship ids that only ``removed`` children of ``root`` use."""
    kept = [child for child in root if child not in removed]
    return sorted(_rel_ids(removed) - _rel_ids(kept))


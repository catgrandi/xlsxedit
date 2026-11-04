"""Hyperlink helpers for worksheet cells."""

from __future__ import annotations

import functools
from typing import TYPE_CHECKING, Any, Callable, TypeVar

from lxml import etree

from xlsxedit.opc.constants import OFFICE_REL_NS, RT, SML_NS
from xlsxedit.worksheet_order import insert_worksheet_child, reposition_worksheet_child

if TYPE_CHECKING:
    from xlsxedit.cell import Cell

_HYPERLINKS = f"{{{SML_NS}}}hyperlinks"
_HYPERLINK = f"{{{SML_NS}}}hyperlink"

_Setter = TypeVar("_Setter", bound=Callable[..., Any])


def _marks_sheet(setter: _Setter) -> _Setter:
    """Mark the worksheet's part dirty when ``setter`` changes the cell's hyperlink.

    A setter changes only the cell's ``<hyperlink>`` and the ``<hyperlinks>``
    block that holds it, so comparing them before and after tells; a call that
    leaves them as they were, such as setting the location a link already
    has, marks nothing.
    """

    @functools.wraps(setter)
    def wrapper(self: Hyperlink, value: Any) -> None:
        part = self._cell._worksheet._part
        if part.is_dirty:
            return setter(self, value)
        before = self._state()
        try:
            return setter(self, value)
        finally:
            if self._state() != before:
                part.mark_dirty()

    return wrapper  # type: ignore[return-value]


def _ensure_hyperlinks_block(ws_elm: etree._Element) -> etree._Element:
    block = ws_elm.find(_HYPERLINKS)
    if block is None:
        block = etree.Element(_HYPERLINKS)
        insert_worksheet_child(ws_elm, block)
    else:
        reposition_worksheet_child(ws_elm, block)
    return block


class Hyperlink:
    """Read/write hyperlink on a single cell."""

    def __init__(self, cell: Cell):
        self._cell = cell

    def __repr__(self) -> str:
        return f"<Hyperlink url={self.url!r} location={self.location!r}>"

    def _hyperlink_element(self) -> etree._Element | None:
        addr = self._cell.address
        ws_elm = self._cell._worksheet._part.element
        block = ws_elm.find(_HYPERLINKS)
        if block is None:
            return None
        for hl in block.findall(_HYPERLINK):
            if hl.get("ref") == addr:
                return hl
        return None

    def _links_to(self, url: str) -> bool:
        """Whether the cell's hyperlink already goes to ``url``, and nowhere inside it."""
        hl = self._hyperlink_element()
        r_id = None if hl is None else hl.get(f"{{{OFFICE_REL_NS}}}id")
        rel = None if r_id is None else self._cell._worksheet._part.rels.get(r_id)
        return (
            rel is not None
            and rel.is_external
            and rel.reltype == RT.HYPERLINK
            and rel.target_ref == url
            and hl.get("location") is None
        )

    def _state(self) -> tuple[int | None, bytes | None]:
        """Where the ``<hyperlinks>`` block sits, and the cell's ``<hyperlink>``."""
        ws_elm = self._cell._worksheet._part.element
        block = ws_elm.find(_HYPERLINKS)
        if block is None:
            return None, None
        hl = self._hyperlink_element()
        return ws_elm.index(block), None if hl is None else etree.tostring(hl, with_tail=False)

    @property
    def url(self) -> str | None:
        hl = self._hyperlink_element()
        if hl is None:
            return None
        r_id = hl.get(f"{{{OFFICE_REL_NS}}}id")
        if r_id is None:
            return None
        try:
            rel = self._cell._worksheet._part.rels[r_id]
        except KeyError:
            return None
        if rel.is_external:
            return rel.target_ref
        return None

    @property
    def location(self) -> str | None:
        """Internal link target, e.g. ``Sheet2!B1`` (no external relationship)."""
        hl = self._hyperlink_element()
        return hl.get("location") if hl is not None else None

    @property
    def display(self) -> str | None:
        hl = self._hyperlink_element()
        return hl.get("display") if hl is not None else None

    @url.setter
    @_marks_sheet
    def url(self, value: str | None) -> None:
        if value is None:
            self._remove()
            return
        address = self._cell.address  # raises, changing nothing, for a cell that cannot have one
        etree.Element(_HYPERLINK).set("location", value)  # lxml refuses text XML cannot store
        if self._links_to(value):
            return
        ws_part = self._cell._worksheet._part
        ws_elm = ws_part.element
        block = _ensure_hyperlinks_block(ws_elm)

        hl = self._hyperlink_element()
        if hl is None:
            hl = etree.SubElement(block, _HYPERLINK)
            hl.set("ref", address)
        else:
            old_r_id = hl.get(f"{{{OFFICE_REL_NS}}}id")
            if old_r_id is not None:
                _release_rel(ws_part, old_r_id, keep=hl)
                hl.attrib.pop(f"{{{OFFICE_REL_NS}}}id", None)

        hl.attrib.pop("location", None)

        r_id = ws_part.rels.get_or_add_ext_rel(RT.HYPERLINK, value)
        hl.set(f"{{{OFFICE_REL_NS}}}id", r_id)

    @location.setter
    @_marks_sheet
    def location(self, value: str | None) -> None:
        if value is None:
            hl = self._hyperlink_element()
            if hl is None:
                return
            if hl.get(f"{{{OFFICE_REL_NS}}}id") is not None:
                hl.attrib.pop("location", None)
            else:
                self._remove()
            return

        address = self._cell.address  # raises, changing nothing, for a cell that cannot have one
        etree.Element(_HYPERLINK).set("location", value)  # lxml refuses text XML cannot store
        ws_part = self._cell._worksheet._part
        ws_elm = ws_part.element
        block = _ensure_hyperlinks_block(ws_elm)

        hl = self._hyperlink_element()
        if hl is None:
            hl = etree.SubElement(block, _HYPERLINK)
            hl.set("ref", address)
        else:
            old_r_id = hl.get(f"{{{OFFICE_REL_NS}}}id")
            if old_r_id is not None:
                _release_rel(ws_part, old_r_id, keep=hl)
                hl.attrib.pop(f"{{{OFFICE_REL_NS}}}id", None)

        hl.set("location", value)

    @display.setter
    @_marks_sheet
    def display(self, value: str | None) -> None:
        hl = self._hyperlink_element()
        if hl is None:
            raise ValueError("set url or location first")
        if value is None:
            hl.attrib.pop("display", None)
        else:
            hl.set("display", value)

    def _remove(self) -> None:
        ws_part = self._cell._worksheet._part
        ws_elm = ws_part.element
        block = ws_elm.find(_HYPERLINKS)
        if block is None:
            return
        hl = self._hyperlink_element()
        if hl is None:
            return
        r_id = hl.get(f"{{{OFFICE_REL_NS}}}id")
        if r_id is not None:
            _release_rel(ws_part, r_id, keep=hl)
        block.remove(hl)
        if len(block) == 0:
            ws_elm.remove(block)


def _release_rel(ws_part, r_id: str, *, keep) -> None:
    """Drop the relationship unless another hyperlink element still uses it."""
    block = ws_part.element.find(_HYPERLINKS)
    if block is not None:
        for other in block.findall(_HYPERLINK):
            if other is not keep and other.get(f"{{{OFFICE_REL_NS}}}id") == r_id:
                return
    ws_part.rels._rels.pop(r_id, None)

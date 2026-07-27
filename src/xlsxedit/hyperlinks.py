"""Hyperlink helpers for worksheet cells."""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from xlsxedit.opc.constants import OFFICE_REL_NS, RT, SML_NS
from xlsxedit.worksheet_order import insert_worksheet_child, reposition_worksheet_child

if TYPE_CHECKING:
    from xlsxedit.cell import Cell

_HYPERLINKS = f"{{{SML_NS}}}hyperlinks"
_HYPERLINK = f"{{{SML_NS}}}hyperlink"


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
        ws_elm = self._cell._worksheet._part.element
        block = ws_elm.find(_HYPERLINKS)
        if block is None:
            return None
        addr = self._cell.address
        for hl in block.findall(_HYPERLINK):
            if hl.get("ref") == addr:
                return hl
        return None

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
    def url(self, value: str | None) -> None:
        if value is None:
            self._remove()
            return
        ws_part = self._cell._worksheet._part
        ws_elm = ws_part.element
        block = _ensure_hyperlinks_block(ws_elm)

        hl = self._hyperlink_element()
        if hl is None:
            hl = etree.SubElement(block, _HYPERLINK)
            hl.set("ref", self._cell.address)
        else:
            old_r_id = hl.get(f"{{{OFFICE_REL_NS}}}id")
            if old_r_id is not None:
                _release_rel(ws_part, old_r_id, keep=hl)
                hl.attrib.pop(f"{{{OFFICE_REL_NS}}}id", None)

        hl.attrib.pop("location", None)

        r_id = ws_part.rels.get_or_add_ext_rel(RT.HYPERLINK, value)
        hl.set(f"{{{OFFICE_REL_NS}}}id", r_id)

    @location.setter
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

        ws_part = self._cell._worksheet._part
        ws_elm = ws_part.element
        block = _ensure_hyperlinks_block(ws_elm)

        hl = self._hyperlink_element()
        if hl is None:
            hl = etree.SubElement(block, _HYPERLINK)
            hl.set("ref", self._cell.address)
        else:
            old_r_id = hl.get(f"{{{OFFICE_REL_NS}}}id")
            if old_r_id is not None:
                _release_rel(ws_part, old_r_id, keep=hl)
                hl.attrib.pop(f"{{{OFFICE_REL_NS}}}id", None)

        hl.set("location", value)

    @display.setter
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

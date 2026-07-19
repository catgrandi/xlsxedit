"""Hyperlink helpers for worksheet cells."""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from xlsxedit.opc.constants import OFFICE_REL_NS, RT, SML_NS

if TYPE_CHECKING:
    from xlsxedit.cell import Cell

_HYPERLINKS = f"{{{SML_NS}}}hyperlinks"
_HYPERLINK = f"{{{SML_NS}}}hyperlink"
_SHEET_DATA = f"{{{SML_NS}}}sheetData"

# OOXML: these elements must appear before <hyperlinks> on a worksheet.
_BEFORE_HYPERLINKS = frozenset({
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "sheetCalcPr",
    "sheetProtection",
    "protectedRanges",
    "scenarios",
    "autoFilter",
    "sortState",
    "dataConsolidate",
    "customSheetViews",
    "mergeCells",
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
})


def _hyperlinks_insert_index(ws_elm: etree._Element) -> int:
    """Return the child index where a new ``<hyperlinks>`` block belongs."""
    insert_at = len(list(ws_elm))
    for i, child in enumerate(ws_elm):
        local = etree.QName(child).localname
        if local == "hyperlinks":
            return i
        if local in _BEFORE_HYPERLINKS:
            insert_at = i + 1
    return insert_at


def _fix_hyperlinks_position(ws_elm: etree._Element, block: etree._Element) -> None:
    """Move ``<hyperlinks>`` after ``<sheetData>`` if it was inserted too early."""
    sheet_data = ws_elm.find(_SHEET_DATA)
    if sheet_data is None:
        return
    block_idx = list(ws_elm).index(block)
    target_idx = _hyperlinks_insert_index(ws_elm)
    if block_idx != target_idx:
        ws_elm.remove(block)
        ws_elm.insert(target_idx, block)


def _ensure_hyperlinks_block(ws_elm: etree._Element) -> etree._Element:
    block = ws_elm.find(_HYPERLINKS)
    if block is None:
        block = etree.Element(_HYPERLINKS)
        ws_elm.insert(_hyperlinks_insert_index(ws_elm), block)
    else:
        _fix_hyperlinks_position(ws_elm, block)
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

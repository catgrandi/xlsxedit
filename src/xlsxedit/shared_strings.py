"""Shared string table helpers."""

from __future__ import annotations

from lxml import etree
from lxml.etree import _Element

from xlsxedit.opc.constants import SML_NS

_SI = f"{{{SML_NS}}}si"
_T = f"{{{SML_NS}}}t"
_R = f"{{{SML_NS}}}r"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _t_nodes_in_si(si: _Element) -> list[_Element]:
    """Return all ``<t>`` text nodes in a shared-string ``<si>`` (plain or rich)."""
    direct = list(si.findall(_T))
    if direct:
        return direct
    nodes: list[_Element] = []
    for r in si.findall(_R):
        nodes.extend(r.findall(_T))
    return nodes


def si_display_text(si: _Element) -> str:
    return "".join((t.text or "") for t in _t_nodes_in_si(si))


def replace_in_si(si: _Element, old: str, new: str) -> int:
    """Replace ``old`` with ``new`` inside one SST entry."""
    if not old:
        return 0
    total = 0
    for t_elm in _t_nodes_in_si(si):
        text = t_elm.text or ""
        if old not in text:
            continue
        count = text.count(old)
        t_elm.text = text.replace(old, new)
        if " " in (t_elm.text or "") or (t_elm.text or "").startswith(" ") or (
            t_elm.text or ""
        ).endswith(" "):
            t_elm.set(_XML_SPACE, "preserve")
        total += count
    return total


class SharedStringTable:
    """Proxy over the sharedStrings part element."""

    def __init__(self, element: _Element, *, defer_index: bool = False):
        self._element = element
        self._sis: list[_Element] | None = None if defer_index else list(element.findall(_SI))

    def _ensure_sis(self) -> list[_Element]:
        if self._sis is None:
            self._sis = list(self._element.findall(_SI))
        return self._sis

    def __len__(self) -> int:
        return len(self._ensure_sis())

    def get(self, index: int) -> _Element:
        return self._ensure_sis()[index]

    def text(self, index: int) -> str:
        return si_display_text(self._ensure_sis()[index])

    def replace_in_index(self, index: int, old: str, new: str) -> int:
        return replace_in_si(self._ensure_sis()[index], old, new)

    def replace_all(self, old: str, new: str) -> int:
        """Replace in every SST entry. Prefer cell-driven replace for worksheets."""
        return sum(replace_in_si(si, old, new) for si in self._ensure_sis())

    def add(self, text: str) -> int:
        """Append a plain string entry; return its 0-based index."""
        si = etree.SubElement(self._element, _SI)
        t_elm = etree.SubElement(si, _T)
        t_elm.text = text
        if " " in text or text.startswith(" ") or text.endswith(" "):
            t_elm.set(_XML_SPACE, "preserve")
        sis = self._ensure_sis()
        sis.append(si)
        index = len(sis) - 1
        self._element.set("count", str(int(self._element.get("count", "0")) + 1))
        self._element.set("uniqueCount", str(len(sis)))
        return index

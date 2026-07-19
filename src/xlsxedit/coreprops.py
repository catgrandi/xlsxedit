"""Core (Dublin Core) document properties — ``/docProps/core.xml``."""

from __future__ import annotations

from datetime import datetime, timezone

from lxml import etree
from lxml.etree import _Element

_CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
_DC_NS = "http://purl.org/dc/elements/1.1/"
_DCTERMS_NS = "http://purl.org/dc/terms/"
_XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"

CORE_PROPS_NSMAP = {
    "cp": _CP_NS,
    "dc": _DC_NS,
    "dcterms": _DCTERMS_NS,
    "xsi": _XSI_NS,
}

_XSI_TYPE = f"{{{_XSI_NS}}}type"

# Schema-defined child order of <cp:coreProperties>.
_CHILD_ORDER = (
    f"{{{_CP_NS}}}category",
    f"{{{_CP_NS}}}contentStatus",
    f"{{{_DCTERMS_NS}}}created",
    f"{{{_DC_NS}}}creator",
    f"{{{_DC_NS}}}description",
    f"{{{_DC_NS}}}identifier",
    f"{{{_CP_NS}}}keywords",
    f"{{{_DC_NS}}}language",
    f"{{{_CP_NS}}}lastModifiedBy",
    f"{{{_CP_NS}}}lastPrinted",
    f"{{{_DCTERMS_NS}}}modified",
    f"{{{_CP_NS}}}revision",
    f"{{{_DC_NS}}}subject",
    f"{{{_DC_NS}}}title",
    f"{{{_CP_NS}}}version",
)


def new_core_properties_element() -> _Element:
    """Return an empty ``<cp:coreProperties>`` root element."""
    return etree.Element(f"{{{_CP_NS}}}coreProperties", nsmap=CORE_PROPS_NSMAP)


def _parse_w3cdtf(text: str) -> datetime | None:
    text = text.strip()
    if not text:
        return None
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        return datetime.fromisoformat(text)
    except ValueError:
        return None


def _format_w3cdtf(value: datetime) -> str:
    if value.tzinfo is not None:
        value = value.astimezone(timezone.utc).replace(tzinfo=None)
    return value.strftime("%Y-%m-%dT%H:%M:%SZ")


class CoreProperties:
    """Read/write proxy over the ``<cp:coreProperties>`` element."""

    def __init__(self, element: _Element):
        self._element = element

    def _get_text(self, tag: str) -> str | None:
        child = self._element.find(tag)
        if child is None:
            return None
        return child.text

    def _get_or_add(self, tag: str) -> _Element:
        child = self._element.find(tag)
        if child is not None:
            return child
        child = etree.Element(tag)
        order_idx = _CHILD_ORDER.index(tag)
        for existing in self._element:
            if existing.tag in _CHILD_ORDER and _CHILD_ORDER.index(existing.tag) > order_idx:
                existing.addprevious(child)
                return child
        self._element.append(child)
        return child

    def _set_text(self, tag: str, value: str) -> None:
        self._get_or_add(tag).text = value

    def _get_datetime(self, tag: str) -> datetime | None:
        text = self._get_text(tag)
        if text is None:
            return None
        return _parse_w3cdtf(text)

    def _set_datetime(self, tag: str, value: datetime) -> None:
        child = self._get_or_add(tag)
        child.set(_XSI_TYPE, "dcterms:W3CDTF")
        child.text = _format_w3cdtf(value)

    @property
    def author(self) -> str | None:
        return self._get_text(f"{{{_DC_NS}}}creator")

    @author.setter
    def author(self, value: str) -> None:
        self._set_text(f"{{{_DC_NS}}}creator", value)

    @property
    def title(self) -> str | None:
        return self._get_text(f"{{{_DC_NS}}}title")

    @title.setter
    def title(self, value: str) -> None:
        self._set_text(f"{{{_DC_NS}}}title", value)

    @property
    def subject(self) -> str | None:
        return self._get_text(f"{{{_DC_NS}}}subject")

    @subject.setter
    def subject(self, value: str) -> None:
        self._set_text(f"{{{_DC_NS}}}subject", value)

    @property
    def keywords(self) -> str | None:
        return self._get_text(f"{{{_CP_NS}}}keywords")

    @keywords.setter
    def keywords(self, value: str) -> None:
        self._set_text(f"{{{_CP_NS}}}keywords", value)

    @property
    def comments(self) -> str | None:
        return self._get_text(f"{{{_DC_NS}}}description")

    @comments.setter
    def comments(self, value: str) -> None:
        self._set_text(f"{{{_DC_NS}}}description", value)

    @property
    def category(self) -> str | None:
        return self._get_text(f"{{{_CP_NS}}}category")

    @category.setter
    def category(self, value: str) -> None:
        self._set_text(f"{{{_CP_NS}}}category", value)

    @property
    def last_modified_by(self) -> str | None:
        return self._get_text(f"{{{_CP_NS}}}lastModifiedBy")

    @last_modified_by.setter
    def last_modified_by(self, value: str) -> None:
        self._set_text(f"{{{_CP_NS}}}lastModifiedBy", value)

    @property
    def created(self) -> datetime | None:
        return self._get_datetime(f"{{{_DCTERMS_NS}}}created")

    @created.setter
    def created(self, value: datetime) -> None:
        self._set_datetime(f"{{{_DCTERMS_NS}}}created", value)

    @property
    def modified(self) -> datetime | None:
        return self._get_datetime(f"{{{_DCTERMS_NS}}}modified")

    @modified.setter
    def modified(self, value: datetime) -> None:
        self._set_datetime(f"{{{_DCTERMS_NS}}}modified", value)

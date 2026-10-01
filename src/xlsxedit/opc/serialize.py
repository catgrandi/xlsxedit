"""Serialize the XML items the package layer writes, in the form Excel writes them."""

from __future__ import annotations

from lxml import etree

XML_DECLARATION = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'


def serialize_part_xml(element: etree._Element) -> bytes:
    """``element`` as a package item: Excel's double-quoted declaration, CRLF, then the tree."""
    return XML_DECLARATION + etree.tostring(element, encoding="UTF-8", xml_declaration=False)

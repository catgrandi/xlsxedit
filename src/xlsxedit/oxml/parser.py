"""XML parse and serialize for package parts.

Adapted from the oxml layer of python-docx / python-pptx by Steve Canny
(MIT License, Copyright (c) 2013 Steve Canny). See NOTICE and
THIRD_PARTY_LICENSES at the project root.
"""

from __future__ import annotations

from lxml import etree

_PARSER = etree.XMLParser(remove_blank_text=False, huge_tree=True)


def parse_xml(xml: bytes | str) -> etree._Element:
    if isinstance(xml, str):
        xml = xml.encode("utf-8")
    return etree.fromstring(xml, parser=_PARSER)


def serialize_xml(element: etree._Element) -> bytes:
    """Serialize an XML element as a package part blob."""
    return etree.tostring(
        element,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )

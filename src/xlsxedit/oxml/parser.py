"""XML parse and serialize for package parts.

Adapted from the oxml layer of python-docx / python-pptx by Steve Canny
(MIT License, Copyright (c) 2013 Steve Canny). See NOTICE and
THIRD_PARTY_LICENSES at the project root.
"""

from __future__ import annotations

from lxml import etree

from xlsxedit.exceptions import DTDForbiddenError

# Entities stay unexpanded and nothing is fetched, so hostile parts can neither
# read local files (XXE) nor multiply text (entity bombs); huge_tree only lifts
# libxml2's size limits for large worksheets.
_PARSER = etree.XMLParser(
    remove_blank_text=False, huge_tree=True, resolve_entities=False, no_network=True
)
_TEMPLATE_PARSER = etree.XMLParser(
    remove_blank_text=True, huge_tree=True, resolve_entities=False, no_network=True
)


def _refuse_dtd(root: etree._Element) -> etree._Element:
    """Reject a DTD: its unexpanded entity references would not survive a save."""
    doctype = root.getroottree().docinfo.doctype
    if doctype:
        raise DTDForbiddenError(
            f"XML part declares a DTD ({doctype}); Excel never writes one, so it is not loaded"
        )
    return root


def parse_xml(xml: bytes | str) -> etree._Element:
    if isinstance(xml, str):
        xml = xml.encode("utf-8")
    return _refuse_dtd(etree.fromstring(xml, parser=_PARSER))


def parse_template_xml(xml: bytes | str) -> etree._Element:
    """Parse a bundled template fragment; strip indent so saved parts stay compact."""
    if isinstance(xml, str):
        xml = xml.encode("utf-8")
    return _refuse_dtd(etree.fromstring(xml, parser=_TEMPLATE_PARSER))


def serialize_xml(element: etree._Element) -> bytes:
    """Serialize an XML element as a package part blob."""
    return etree.tostring(
        element,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )

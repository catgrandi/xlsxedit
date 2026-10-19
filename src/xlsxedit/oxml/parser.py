"""XML parse and serialize for package parts.

Adapted from the oxml layer of python-docx / python-pptx by Steve Canny
(MIT License, Copyright (c) 2013 Steve Canny). See NOTICE and
THIRD_PARTY_LICENSES at the project root.
"""

from __future__ import annotations

import codecs
import re

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


_BYTE_ORDER_MARKS = (  # UTF-32 first: its little-endian mark starts like UTF-16's
    (codecs.BOM_UTF32_LE, "utf-32-le"),
    (codecs.BOM_UTF32_BE, "utf-32-be"),
    (codecs.BOM_UTF8, "latin-1"),
    (codecs.BOM_UTF16_LE, "utf-16-le"),
    (codecs.BOM_UTF16_BE, "utf-16-be"),
)
_PROLOG_SPACE = re.compile(r"[ \t\r\n]*")
_DOCTYPE_NAME = re.compile(r"[ \t\r\n]*([^ \t\r\n[>]*)")
_DECLARED_ENCODING = re.compile(
    r"<\?xml[^>]*?encoding[ \t\r\n]*=[ \t\r\n]*[\x22']([A-Za-z0-9._-]+)"
)
_MORE = object()


def _prolog_codec(blob: bytes) -> tuple[str, int]:
    """The codec that decodes the markup of ``blob`` and the length of its byte-order mark.

    Without a mark, zero bytes around the first character tell UTF-16 and
    UTF-32 apart; Latin-1 then stands in for UTF-8 and every other encoding
    that writes markup as ASCII.
    """
    for mark, codec in _BYTE_ORDER_MARKS:
        if blob.startswith(mark):
            return codec, len(mark)
    head = blob[:4]
    if len(head) == 4 and head[1:] == bytes(3):
        return "utf-32-le", 0
    if len(head) == 4 and head[:3] == bytes(3):
        return "utf-32-be", 0
    if len(head) >= 2 and head[0] and not head[1]:
        return "utf-16-le", 0
    if len(head) >= 2 and head[1] and not head[0]:
        return "utf-16-be", 0
    return "latin-1", 0


def _doctype_in_prolog(text: str, complete: bool):
    """The root name a DTD in the prolog that starts ``text`` declares.

    ``None`` when the prolog ends without one; ``_MORE`` when ``text`` ends
    inside the prolog and ``complete`` says more text follows.
    """
    i = 0
    while True:
        i = _PROLOG_SPACE.match(text, i).end()
        if text.startswith("<?", i):
            end, close = text.find("?>", i + 2), 2
        elif text.startswith("<!--", i):
            end, close = text.find("-->", i + 4), 3
        elif text.startswith("<!DOCTYPE", i):
            name = _DOCTYPE_NAME.match(text, i + 9)
            return _MORE if name.end() == len(text) and not complete else name.group(1)
        else:
            return _MORE if len(text) - i < 9 and not complete else None
        if end < 0:
            return None if complete else _MORE
        i = end + close


def refuse_dtd_in_prolog(blob: bytes, partname: str) -> None:
    """Raise ``DTDForbiddenError`` if the XML in ``blob`` declares a DTD, without parsing it.

    A DTD can only come before the root element, so only the prolog is read:
    a byte-order mark, the XML declaration, comments, processing instructions
    and white space, in UTF-8, UTF-16 or UTF-32. Anything else ends the scan;
    markup that is not well formed is left for the parser to report.
    """
    codec, start = _prolog_codec(blob)
    if codec == "latin-1":
        # UTF-7 writes "<" as "+ADw-", so its markup is only seen once decoded.
        declared = _DECLARED_ENCODING.match(blob[start : start + 256].decode("latin-1"))
        if declared and declared.group(1).lower().replace("_", "-") == "utf-7":
            codec = "utf-7"
    size = 4096
    while True:
        complete = start + size >= len(blob)
        name = _doctype_in_prolog(blob[start : start + size].decode(codec, "ignore"), complete)
        if name is not _MORE:
            break
        size *= 4
    if name is not None:
        raise DTDForbiddenError(
            f"{partname} declares a DTD (<!DOCTYPE {name}>); "
            "Excel never writes one, so it is not loaded"
        )


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

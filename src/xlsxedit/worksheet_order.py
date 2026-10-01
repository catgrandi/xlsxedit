"""OOXML worksheet child-element order (ECMA-376).

Excel rejects worksheets when children are out of schema sequence
(e.g. ``tableParts`` before ``conditionalFormatting`` / ``drawing``).
"""

from __future__ import annotations

from lxml import etree

# Sequence from SpreadsheetML CT_Worksheet (unbounded kinds may repeat).
_WS_CHILD_ORDER: tuple[str, ...] = (
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
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "drawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
)

_WS_CHILD_RANK: dict[str, int] = {name: i for i, name in enumerate(_WS_CHILD_ORDER)}


def _child_rank(child: etree._Element) -> int | None:
    if not isinstance(child.tag, str):  # comment or processing instruction
        return None
    return _WS_CHILD_RANK.get(etree.QName(child).localname)


def worksheet_child_insert_index(ws_elm: etree._Element, localname: str) -> int:
    """Index at which to insert a new child with ``localname``.

    The new child goes after every child the sequence places at or before it
    and before the first child it places later. Children the sequence does not
    name never move and do not affect the result.

    Raises ``ValueError`` when ``localname`` is not a ``CT_Worksheet`` child:
    guessing a position would leave the sheet schema-invalid.
    """
    rank = _WS_CHILD_RANK.get(localname)
    if rank is None:
        raise ValueError(f"<{localname}> is not a CT_Worksheet child element")
    for i, child in enumerate(ws_elm):
        child_rank = _child_rank(child)
        if child_rank is not None and child_rank > rank:
            return i
    return len(ws_elm)


def insert_worksheet_child(ws_elm: etree._Element, elm: etree._Element) -> None:
    """Insert ``elm`` into ``ws_elm`` at the schema-correct position."""
    localname = etree.QName(elm).localname
    ws_elm.insert(worksheet_child_insert_index(ws_elm, localname), elm)


def reposition_worksheet_child(ws_elm: etree._Element, elm: etree._Element) -> None:
    """Move an existing child to its schema-correct index if it is out of order.

    Leaves the tree alone when ``elm`` is already in order or is not a
    ``CT_Worksheet`` child.
    """
    if elm.getparent() is not ws_elm:
        return
    localname = etree.QName(elm).localname
    rank = _WS_CHILD_RANK.get(localname)
    if rank is None:
        return
    preceding = (_child_rank(c) for c in elm.itersiblings(preceding=True))
    following = (_child_rank(c) for c in elm.itersiblings())
    if all(r is None or r <= rank for r in preceding) and all(
        r is None or r >= rank for r in following
    ):
        return
    ws_elm.remove(elm)
    ws_elm.insert(worksheet_child_insert_index(ws_elm, localname), elm)

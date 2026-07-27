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
    "drawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
)

_WS_CHILD_RANK: dict[str, int] = {name: i for i, name in enumerate(_WS_CHILD_ORDER)}


def worksheet_child_insert_index(ws_elm: etree._Element, localname: str) -> int:
    """Index at which to insert a new child with ``localname``."""
    rank = _WS_CHILD_RANK.get(localname)
    if rank is None:
        return len(list(ws_elm))
    for i, child in enumerate(ws_elm):
        child_rank = _WS_CHILD_RANK.get(etree.QName(child).localname)
        if child_rank is not None and child_rank > rank:
            return i
    return len(list(ws_elm))


def insert_worksheet_child(ws_elm: etree._Element, elm: etree._Element) -> None:
    """Insert ``elm`` into ``ws_elm`` at the schema-correct position."""
    localname = etree.QName(elm).localname
    ws_elm.insert(worksheet_child_insert_index(ws_elm, localname), elm)


def reposition_worksheet_child(ws_elm: etree._Element, elm: etree._Element) -> None:
    """Move an existing child to its schema-correct index if needed."""
    if elm.getparent() is not ws_elm:
        return
    localname = etree.QName(elm).localname
    current = list(ws_elm).index(elm)
    # Compute target as if elm were not present
    ws_elm.remove(elm)
    target = worksheet_child_insert_index(ws_elm, localname)
    ws_elm.insert(target, elm)
    if current != target:
        pass  # moved

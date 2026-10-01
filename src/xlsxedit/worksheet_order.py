"""Schema order for the child elements of OOXML parts (ECMA-376).

Excel repairs a part whose children are out of schema sequence, for example
``tableParts`` before ``conditionalFormatting`` or ``drawing`` in a worksheet.
``insert_ordered`` and ``reposition_ordered`` place a child by one of the
sequences in ``xlsxedit._ooxml_order.ORDER``; the ``*_worksheet_child``
functions apply them to ``CT_Worksheet``.

Children match the sequence by local name. A child the sequence does not name,
such as a comment or a foreign element, never moves and does not affect where
a new child goes.
"""

from __future__ import annotations

from functools import lru_cache

from lxml import etree

from xlsxedit._ooxml_order import ORDER

_WS_CHILD_ORDER: tuple[str, ...] = ORDER["CT_Worksheet"]
_WS_CHILD_RANK: dict[str, int] = {name: i for i, name in enumerate(_WS_CHILD_ORDER)}


@lru_cache(maxsize=None)
def _ranks(order: tuple[str, ...]) -> dict[str, int]:
    return {name: i for i, name in enumerate(order)}


def _child_rank(child: etree._Element, ranks: dict[str, int]) -> int | None:
    if not isinstance(child.tag, str):  # comment or processing instruction
        return None
    return ranks.get(etree.QName(child).localname)


def ordered_insert_index(
    parent: etree._Element, localname: str, order: tuple[str, ...]
) -> int:
    """Index at which a new ``localname`` child of ``parent`` keeps ``order``.

    The new child goes after every child that ``order`` places at or before it
    and before the first child that ``order`` places later.

    Raises ``ValueError`` when ``order`` does not name ``localname``: guessing
    a position would leave the part schema-invalid.
    """
    ranks = _ranks(order)
    rank = ranks.get(localname)
    if rank is None:
        raise ValueError(f"<{localname}> is not a child element in this schema sequence")
    for i, child in enumerate(parent):
        child_rank = _child_rank(child, ranks)
        if child_rank is not None and child_rank > rank:
            return i
    return len(parent)


def insert_ordered(
    parent: etree._Element, elm: etree._Element, order: tuple[str, ...]
) -> None:
    """Insert ``elm`` into ``parent`` at its position in ``order``."""
    parent.insert(ordered_insert_index(parent, etree.QName(elm).localname, order), elm)


def reposition_ordered(
    parent: etree._Element, elm: etree._Element, order: tuple[str, ...]
) -> bool:
    """Move child ``elm`` of ``parent`` to its position in ``order``.

    Returns whether ``elm`` moved. Leaves the tree alone when ``elm`` is
    already in order, is not a child of ``parent``, or is not named in
    ``order``.
    """
    if elm.getparent() is not parent:
        return False
    ranks = _ranks(order)
    localname = etree.QName(elm).localname
    rank = ranks.get(localname)
    if rank is None:
        return False
    preceding = (_child_rank(c, ranks) for c in elm.itersiblings(preceding=True))
    following = (_child_rank(c, ranks) for c in elm.itersiblings())
    if all(r is None or r <= rank for r in preceding) and all(
        r is None or r >= rank for r in following
    ):
        return False
    parent.remove(elm)
    parent.insert(ordered_insert_index(parent, localname, order), elm)
    return True


def worksheet_child_insert_index(ws_elm: etree._Element, localname: str) -> int:
    """``ordered_insert_index`` for a ``CT_Worksheet`` child."""
    return ordered_insert_index(ws_elm, localname, _WS_CHILD_ORDER)


def insert_worksheet_child(ws_elm: etree._Element, elm: etree._Element) -> None:
    """``insert_ordered`` for a ``CT_Worksheet`` child."""
    insert_ordered(ws_elm, elm, _WS_CHILD_ORDER)


def reposition_worksheet_child(ws_elm: etree._Element, elm: etree._Element) -> bool:
    """``reposition_ordered`` for a ``CT_Worksheet`` child."""
    return reposition_ordered(ws_elm, elm, _WS_CHILD_ORDER)

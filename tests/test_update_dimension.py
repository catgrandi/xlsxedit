"""``update_dimension`` reads every cell address in one pass of the regex engine
(#42); inserts and bulk writes keep ``<dimension>`` without reading them again (#56)."""

from __future__ import annotations

import pytest
from lxml import etree

from xlsxedit import Workbook
from xlsxedit.exceptions import InvalidRangeError
from xlsxedit.opc.constants import SML_NS
from xlsxedit.parts import WorksheetPart
from tests.conftest import BOOK1, INSPECT_FIXTURES

_NS = f'xmlns="{SML_NS}"'

# The sheets below hold out-of-order and invalid addresses on purpose.
pytestmark = pytest.mark.usefixtures("unchecked_workbooks")


def _sheet_with(cells: str, dimension: str | None = None):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    root = ws._part.element
    old = root.find(f"{{{SML_NS}}}sheetData")
    root.replace(old, etree.fromstring(f"<sheetData {_NS}>{cells}</sheetData>"))
    if dimension is not None:
        root.insert(0, etree.Element(f"{{{SML_NS}}}dimension", ref=dimension))
    return ws


def _dimension(ws) -> str | None:
    dim = ws._part.element.find(f"{{{SML_NS}}}dimension")
    return None if dim is None else dim.get("ref")


def _rescanned(ws) -> str | None:
    """``<dimension>`` as ``update_dimension`` reads the cells, leaving the sheet as it was."""
    dim = ws._part.element.find(f"{{{SML_NS}}}dimension")
    kept = None if dim is None else (dim.get("ref"), ws._part.element.index(dim))
    ws.update_dimension()
    found = _dimension(ws)
    dim = ws._part.element.find(f"{{{SML_NS}}}dimension")
    if kept is None and dim is not None:
        ws._part.element.remove(dim)
    elif kept is not None:
        dim.set("ref", kept[0])
    return found


# Numbers only, and nothing in a header cell, so that no fixture refuses them.
_INSERTS = {
    "rows at 1": lambda ws: ws.insert_rows([[1, 2]], at_row=1),
    "rows at 3": lambda ws: ws.insert_rows([[1], [None, 2]], at_row=3),
    "rows past the cells": lambda ws: ws.insert_rows([[1]], at_row=40),
    "empty rows": lambda ws: ws.insert_rows([[], []], at_row=2),
    "columns at A": lambda ws: ws.insert_columns([[1, 2]], at_col="A"),
    "columns at C": lambda ws: ws.insert_columns([[None], [None, 3]], at_col="C"),
    "columns at K": lambda ws: ws.insert_columns([[4]], at_cell="K4"),
}


@pytest.mark.parametrize("insert", _INSERTS.values(), ids=_INSERTS)
@pytest.mark.parametrize("name", ["Book1", *INSPECT_FIXTURES])
def test_an_insert_leaves_the_dimension_update_dimension_would_write(name: str, insert):
    path = BOOK1 if name == "Book1" else INSPECT_FIXTURES[name]
    for sheet in Workbook.open(path).sheetnames:
        wb = Workbook.open(path)
        ws = wb[sheet]
        if not isinstance(ws._part, WorksheetPart):
            continue
        insert(ws)
        assert _dimension(ws) == _rescanned(ws), sheet


_IRREGULAR = {
    "rows without r": '<row><c r="A1"><v>1</v></c></row><row><c><v>2</v></c></row>'
    '<row r="5"><c r="B5"><v>5</v></c></row>',
    "cells without r": '<row r="1"><c><v>1</v></c><c><v>2</v></c></row><row r="2"><c r="D2"/><c/></row>',
    "lowercase": '<row r="1"><c r="a1"/><c r="b1"/></row><row r="4"><c r="c4"/></row>',
    "a cell below its row": '<row r="2"><c r="A2"/><c r="B9"/></row><row r="5"><c r="C5"/></row>',
    "a cell above its row": '<row r="7"><c r="A1"/></row><row r="8"><c r="B8"/></row>',
    "leading zeros": '<row r="003"><c r="A003"/><c r="B3"/></row><row r="5"><c r="A5"/></row>',
    "a row extension": '<row r="2"><c r="A2"/><extLst><ext uri="{X}"/></extLst></row>'
    '<row r="4"><c r="B4"/></row>',
    "rows out of order": '<row r="5"><c r="B5"/></row><row r="1"><c r="A1"/></row>',
    "a row at the last row": '<row r="1"><c r="A1"/></row><row r="1048576" ht="20" customHeight="1"/>',
    "a cell at the last column": '<row r="1"><c r="A1"/><c r="XFD1" s="1"/></row>',
    "no cells": '<row r="2" ht="20" customHeight="1"/>',
    "one zero-padded cell": '<row r="02"><c r="B02"/></row>',
    "a row listed twice": '<row r="2"><c r="B2"/></row><row r="2"/>',
}


@pytest.mark.parametrize("insert", _INSERTS.values(), ids=_INSERTS)
@pytest.mark.parametrize("dimension", [None, "A1", "B2:E9"])
@pytest.mark.parametrize("cells", _IRREGULAR.values(), ids=_IRREGULAR)
def test_an_insert_on_an_irregular_sheet_leaves_the_dimension_update_dimension_would_write(
    cells: str, dimension: str | None, insert
):
    ws = _sheet_with(cells, dimension)
    insert(ws)
    assert _dimension(ws) == _rescanned(ws)


def test_a_bulk_write_grows_the_dimension():
    ws = _sheet_with('<row r="2"><c r="B2"/></row><row r="9"><c r="E9"/></row>', "B2:E9")
    ws.write_rows([[1]], at_cell="C4")
    assert _dimension(ws) == "B2:E9"
    ws.write_rows([[1, 2], [], [3]], at_cell="A12")
    assert _dimension(ws) == "A2:E14"
    ws.write_rows([[], [None, None, None, None, None, None]], at_cell="A20")
    assert _dimension(ws) == "A2:F21"


@pytest.mark.parametrize("stale", [None, "A1", "Q99", "A1:", "B2:E"])
def test_a_bulk_write_replaces_the_dimension_of_a_sheet_without_cells(stale):
    """Excel writes ``A1`` for a sheet without cells."""
    ws = _sheet_with('<row r="3" ht="20" customHeight="1"/>', stale)
    ws.write_rows([[1, 2]], at_cell="C5")
    assert _dimension(ws) == "C5:D5"


@pytest.mark.parametrize("ref", [None, "A1:", "A:C", "2:4", "A1:XFD1048577", "A1:B" + "9" * 5000])
def test_a_bulk_write_leaves_a_dimension_it_cannot_grow_on_a_sheet_with_cells(ref):
    """Only reading every address could tell what the other cells cover; a ref of
    the written cells alone would leave them out."""
    ws = _sheet_with('<row r="2"><c r="B2"/></row><row r="4"><c r="C4"/></row>', ref)
    ws.write_rows([[7]], at_cell="E6")
    assert _dimension(ws) == ref


def test_a_new_sheet_written_cell_by_cell_and_then_in_bulk_gets_no_partial_dimension():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Title"
    ws.write_rows([[1, 2], [3, 4]], at_row=3)
    assert _dimension(ws) is None
    ws.insert_rows([[0]], at_row=2)
    assert _dimension(ws) == "A1:B5"


def test_a_bulk_write_keeps_the_one_cell_it_does_not_find_by_its_address():
    """A sheet whose only cell is ``B02`` has ``<dimension ref="B2"/>``."""
    ws = _sheet_with('<row r="02"><c r="B02"/></row>', "B2")
    ws.write_rows([[7]], at_cell="A5")
    assert _dimension(ws) == "A2:B5"


@pytest.mark.parametrize("before", [True, False], ids=["before", "after"])
def test_a_dimension_in_its_place_stays_there(before: bool):
    """A comment next to ``<dimension>`` is not a reason to move it."""
    ws = _sheet_with('<row r="1"><c r="A1"/></row><row r="2"><c r="B2"/></row>', "A1:B2")
    root = ws._part.element
    dim = root.find(f"{{{SML_NS}}}dimension")
    (dim.addprevious if before else dim.addnext)(etree.Comment(" generated "))
    order = [etree.QName(child).localname if isinstance(child.tag, str) else "#" for child in root]
    ws.update_dimension()
    ws.write_rows([[5]], at_cell="A2")
    assert [etree.QName(c).localname if isinstance(c.tag, str) else "#" for c in root] == order
    assert _dimension(ws) == "A1:B2"


def test_a_dimension_out_of_its_place_moves_into_it():
    ws = _sheet_with('<row r="1"><c r="A1"/></row>')
    root = ws._part.element
    root.insert(0, etree.Element(f"{{{SML_NS}}}sheetPr"))
    root.append(etree.Element(f"{{{SML_NS}}}dimension", ref="A1"))
    ws.update_dimension()
    names = [etree.QName(child).localname for child in root]
    assert names.index("dimension") == names.index("sheetPr") + 1


@pytest.mark.parametrize("node", ["<!-- a comment -->", "<?generator x?>"])
@pytest.mark.parametrize("write", _INSERTS.values(), ids=_INSERTS)
def test_a_comment_in_the_worksheet_root_leaves_the_dimension_its_place(node: str, write):
    ws = _sheet_with('<row r="1"><c r="A1"/></row><row r="2"><c r="B2"/></row>', "A1:B2")
    ws._part.element.insert(0, etree.fromstring(f"<w>{node}</w>")[0])
    write(ws)
    names = [etree.QName(child).localname for child in ws._part.element if isinstance(child.tag, str)]
    assert names.index("dimension") < names.index("sheetData")
    assert _dimension(ws) == _rescanned(ws)


def test_a_bulk_write_covers_the_merge_anchors_it_writes():
    ws = _sheet_with("")
    ws.merge_cells("A1:C3")
    ws.write_rows([["to the anchor"]], at_cell="B2")
    assert ws["A1"].value == "to the anchor"
    assert _dimension(ws) == "A1:B2"


@pytest.mark.parametrize(
    ("cells", "ref"),
    [
        ('<row r="5"><c r="Z5"/><c r="AA5"/></row><row r="3"><c r="b3"/></row>', "B3:AA5"),
        ('<row r="2"><c r="C2"/></row>', "C2"),
        ('<row r="7"><c r="XFD1048576"/><c r="A007"/></row>', "A7:XFD1048576"),
        ('<row r="4"><c/><c r=""/><c r="D4"/></row><row r="9"/>', "D4"),
        ('<row r="1"><c r="AB12"/><c r="ZZ3"/><c r="AAA2"/></row>', "AB2:AAA12"),
    ],
)
def test_the_dimension_spans_every_addressed_cell(cells: str, ref: str):
    ws = _sheet_with(cells)
    ws.update_dimension()
    assert _dimension(ws) == ref


def test_a_sheet_without_addressed_cells_keeps_no_dimension():
    ws = _sheet_with('<row r="1"><c/></row>')
    ws.update_dimension()
    assert _dimension(ws) is None


@pytest.mark.parametrize("bad", ["1A", "A1B", "A 1", "A-1", "AB", "12"])
def test_an_invalid_address_raises(bad: str):
    ws = _sheet_with(f'<row r="1"><c r="A1"/><c r="{bad}"/><c r="B2"/></row>')
    with pytest.raises(InvalidRangeError, match="invalid cell address"):
        ws.update_dimension()


@pytest.mark.parametrize("row_ref", ["x", ""])
def test_a_row_number_that_is_not_one_raises(row_ref: str):
    ws = _sheet_with(f'<row r="{row_ref}"><c r="A1"/></row>')
    with pytest.raises(ValueError):
        ws.update_dimension()

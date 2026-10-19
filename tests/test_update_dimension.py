"""``update_dimension`` reads every cell address in one pass of the regex engine (#42)."""

from __future__ import annotations

import pytest
from lxml import etree

from xlsxedit import Workbook
from xlsxedit.exceptions import InvalidRangeError
from xlsxedit.opc.constants import SML_NS

_NS = f'xmlns="{SML_NS}"'

# The sheets below hold out-of-order and invalid addresses on purpose.
pytestmark = pytest.mark.usefixtures("unchecked_workbooks")


def _sheet_with(cells: str):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    root = ws._part.element
    old = root.find(f"{{{SML_NS}}}sheetData")
    root.replace(old, etree.fromstring(f"<sheetData {_NS}>{cells}</sheetData>"))
    return ws


def _dimension(ws) -> str | None:
    dim = ws._part.element.find(f"{{{SML_NS}}}dimension")
    return None if dim is None else dim.get("ref")


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

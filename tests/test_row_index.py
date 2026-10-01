"""Cell-API and bulk writes share one row index (#6).

Every test reads back the saved ``sheetData``: one ``<row>`` per number and
one ``<c>`` per address, both in ascending order, plus ``check_consistency``.
"""

from __future__ import annotations

import io

import pytest
from lxml import etree

from xlsxedit import Workbook
from xlsxedit.opc.constants import SML_NS
from tests.preservation import check_consistency, read_pkg
from tests.schema_validate import assert_valid_package

SHEET = "xl/worksheets/sheet1.xml"
_ROW = f"{{{SML_NS}}}row"
_C = f"{{{SML_NS}}}c"


def _layout(wb: Workbook) -> list[tuple[int, list[str]]]:
    """Saved ``sheetData`` as ``(row r, [cell r, ...])`` in document order."""
    pkg = read_pkg(wb)
    check_consistency(pkg)
    sheet_data = etree.fromstring(pkg[SHEET]).find(f"{{{SML_NS}}}sheetData")
    return [
        (int(row.get("r")), [c.get("r") for c in row.iterchildren(_C)])
        for row in sheet_data.iterchildren(_ROW)
    ]


def _reopen(wb: Workbook) -> Workbook:
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return Workbook.open(buf)


def test_cell_write_then_bulk_write_on_a_new_row():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.write_rows([["Item", "Qty"]], at_cell="A1")
    ws["A6"].value = "KEEP ME"
    ws.write_rows([["r5", 5]], at_cell="A6")

    assert _layout(wb) == [(1, ["A1", "B1"]), (6, ["A6", "B6"])]
    assert [ws["A6"].value, ws["B6"].value] == ["r5", 5]


def test_bulk_write_then_cell_write_on_a_new_row():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.write_rows([["Item", "Qty"]], at_cell="A1")
    ws.write_rows([["r5", 5]], at_cell="A6")
    ws["A6"].value = "edited"
    ws["C6"].value = "note"

    assert _layout(wb) == [(1, ["A1", "B1"]), (6, ["A6", "B6", "C6"])]
    assert [ws[a].value for a in ("A6", "B6", "C6")] == ["edited", 5, "note"]


@pytest.mark.parametrize("cell_first", [True, False], ids=["cell-then-bulk", "bulk-then-cell"])
def test_interleaved_writes_on_new_and_existing_rows(cell_first: bool):
    template = Workbook.create()
    template["Sheet1"].write_rows([["b2", "c2"], ["b3", "c3"]], at_cell="B2")
    wb = _reopen(template)
    ws = wb["Sheet1"]

    def cell_writes():
        ws["A2"].value = "cell"
        ws["C2"].value = "cell"
        ws["F6"].value = "cell"
        ws["A4"].value = "cell"

    def bulk_writes():
        ws.write_rows([["bulk", "bulk"]], at_cell="C2")
        ws.write_rows([["bulk", "bulk"]], at_cell="A6")

    steps = [cell_writes, bulk_writes]
    for step in steps if cell_first else reversed(steps):
        step()

    assert _layout(wb) == [
        (2, ["A2", "B2", "C2", "D2"]),
        (3, ["B3", "C3"]),
        (4, ["A4"]),
        (6, ["A6", "B6", "F6"]),
    ]
    assert ws["C2"].value == ("bulk" if cell_first else "cell")


def test_rows_written_out_of_order_are_saved_in_order():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.write_rows([[10], [11], [12]], at_row=10)
    ws["A5"].value = 5
    ws.write_rows([[8]], at_row=8)
    ws["A20"].value = 20
    ws["B11"].value = "b"
    ws.write_rows([[1]], at_row=1)

    assert _layout(wb) == [
        (1, ["A1"]),
        (5, ["A5"]),
        (8, ["A8"]),
        (10, ["A10"]),
        (11, ["A11", "B11"]),
        (12, ["A12"]),
        (20, ["A20"]),
    ]


def test_cells_written_right_to_left_are_saved_in_column_order():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.write_rows([["c", "d"]], at_cell="C1")
    ws["A1"].value = "a"
    ws.write_rows([["b"]], at_cell="B1")
    ws["AA1"].value = "aa"
    ws["Z1"].value = "z"

    assert _layout(wb) == [(1, ["A1", "B1", "C1", "D1", "Z1", "AA1"])]


def test_insert_columns_writes_the_new_column_in_place():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.write_rows([["a", "b", "c"]], at_cell="A1")
    ws.insert_columns([["new"]], at_col="B")

    assert _layout(wb) == [(1, ["A1", "B1", "C1", "D1"])]
    assert [ws[a].value for a in ("A1", "B1", "C1", "D1")] == ["a", "new", "b", "c"]


def test_row_height_and_bulk_write_share_the_row():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.write_rows([["h"]], at_cell="A1")
    ws.row_dimensions[6].height = 30
    ws.write_rows([["r6"]], at_cell="A6")

    assert _layout(wb) == [(1, ["A1"]), (6, ["A6"])]
    assert ws.row_dimensions[6].height == 30


def test_lowercase_addresses_reach_the_same_cell():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.write_rows([[1]], at_cell="A6")
    ws["a6"].value = 2
    ws["b6"].value = 3

    assert _layout(wb) == [(6, ["A6", "B6"])]
    assert ws["A6"].value == 2


def test_new_cells_go_before_the_row_extension_list():
    wb = Workbook.create()
    wb["Sheet1"]["A2"].value = "a"
    etree.SubElement(wb["Sheet1"]["A2"]._element.getparent(), f"{{{SML_NS}}}extLst")
    wb = _reopen(wb)
    ws = wb["Sheet1"]
    ws["C2"].value = "c"
    ws.write_rows([["b"]], at_cell="B2")

    row = ws["A2"]._element.getparent()
    assert [etree.QName(child).localname for child in row] == ["c", "c", "c", "extLst"]
    assert _layout(wb) == [(2, ["A2", "B2", "C2"])]
    assert_valid_package(wb)


def test_older_worksheet_object_sees_writes_through_a_newer_one():
    wb = Workbook.create()
    old = wb["Sheet1"]
    old.write_rows([["a1", "b1"]], at_cell="A1")
    wb.add_worksheet("Other")  # hands out new Worksheet objects
    new = wb["Sheet1"]
    new["C1"].value = "new"
    new.write_rows([["new"]], at_cell="A6")

    old["C1"].value = "old"
    old["B6"].value = "old"
    old.write_rows([["old"]], at_cell="A6")
    assert _layout(wb) == [(1, ["A1", "B1", "C1"]), (6, ["A6", "B6"])]

    new.insert_rows([["inserted"]], at_row=1)
    old["B1"].value = "old"
    old["D2"].value = "old"
    old.write_rows([["old"]], at_cell="C7")
    assert _layout(wb) == [
        (1, ["A1", "B1"]),
        (2, ["A2", "B2", "C2", "D2"]),
        (7, ["A7", "B7", "C7"]),
    ]
    assert [old[a].value for a in ("A1", "B1", "A2", "C2", "D2")] == [
        "inserted",
        "old",
        "a1",
        "old",
        "old",
    ]

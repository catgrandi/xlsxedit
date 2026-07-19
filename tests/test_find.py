"""Tests for Workbook/Worksheet.find, findall, and Cell.offset."""

from __future__ import annotations

from pathlib import Path

import pytest

from xlsxedit import Workbook


def test_worksheet_find_and_findall():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "hello"
    ws["B2"].value = "hello"
    ws["C3"].value = "other"

    found = ws.find("hello")
    assert found is not None
    assert found.address == "A1"

    all_hello = ws.findall("hello")
    assert [c.address for c in all_hello] == ["A1", "B2"]
    assert ws.find("missing") is None
    assert ws.findall("missing") == []


def test_workbook_find_sheet_scope():
    wb = Workbook.create()
    wb["Sheet1"]["A1"].value = "shared"
    other = wb.add_worksheet("Other")
    other["B2"].value = "shared"
    other["C3"].value = "only-other"

    first = wb.find("shared")
    assert first is not None
    assert first.worksheet.name == "Sheet1"
    assert first.address == "A1"

    scoped = wb.find("shared", sheet="Other")
    assert scoped is not None
    assert scoped.address == "B2"

    assert wb.find("only-other", sheet="Sheet1") is None
    assert [c.address for c in wb.findall("shared")] == ["A1", "B2"]
    assert [c.address for c in wb.findall("shared", sheet="Other")] == ["B2"]


def test_find_typed_number_not_string():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = 888
    ws["A2"].value = "888"

    found = ws.find(888)
    assert found is not None
    assert found.address == "A1"
    assert ws.find("888").address == "A2"


def test_find_skips_formula_cells():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].formula = "=1+1"
    # Cached value may look numeric; find must skip formula cells.
    from lxml import etree
    from xlsxedit.opc.constants import SML_NS

    v = etree.SubElement(ws["A1"]._element, f"{{{SML_NS}}}v")
    v.text = "2"
    ws["B1"].value = 2

    assert ws.find(2).address == "B1"
    assert len(ws.findall(2)) == 1


def test_offset_diagonal_round_trip(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A9"].value = "Anchor"
    cell = wb.find("Anchor", sheet="Sheet1")
    assert cell is not None
    assert cell.address == "A9"

    target = cell.offset(cols=1, rows=1)
    assert target.address == "B10"
    target.value = "written via find+offset"

    out = tmp_path / "find_offset.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    assert wb2["Sheet1"]["B10"].value == "written via find+offset"
    assert wb2.find("Anchor").address == "A9"


def test_offset_out_of_range():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "x"
    cell = ws.find("x")
    assert cell is not None
    with pytest.raises(ValueError):
        cell.offset(cols=-1, rows=0)
    with pytest.raises(ValueError):
        cell.offset(cols=0, rows=-1)


def test_cell_worksheet_property():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "x"
    cell = ws.find("x")
    assert cell is not None
    assert cell.worksheet is ws

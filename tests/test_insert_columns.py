"""Tests for Worksheet.insert_columns."""

from __future__ import annotations

from pathlib import Path

import pytest

from xlsxedit import Workbook
from xlsxedit.row_shift import shift_range_ref_cols


def test_shift_range_ref_cols_expand_and_shift():
    # Insert at B (index 1): A1:C1 expands to A1:D1
    assert shift_range_ref_cols("A1:C1", 1, 1) == "A1:D1"
    # Entirely to the right of insert at B: C1 → D1
    assert shift_range_ref_cols("C1", 1, 1) == "D1"
    # Entirely left of insert at C: A1:B1 unchanged
    assert shift_range_ref_cols("A1:B1", 2, 1) == "A1:B1"


def test_insert_columns_shifts_cells_and_writes():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "left"
    ws["C1"].value = "right"
    n = ws.insert_columns([("mid1", "mid2")], at_col="C")
    assert n == 1
    assert ws["A1"].value == "left"
    assert ws["C1"].value == "mid1"
    assert ws["C2"].value == "mid2"
    assert ws["D1"].value == "right"


def test_insert_columns_at_cell_start_row():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["C5"].value = "was-c5"
    ws.insert_columns([("x", "y")], at_cell="C2")
    assert ws["C2"].value == "x"
    assert ws["C3"].value == "y"
    assert ws["D5"].value == "was-c5"


def test_insert_columns_two_columns():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["B1"].value = "B"
    ws.insert_columns([("c1", "c2"), ("d1", "d2")], at_col="B")
    assert ws["B1"].value == "c1"
    assert ws["C1"].value == "d1"
    assert ws["D1"].value == "B"


def test_insert_columns_merge_shifts():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.merge_cells("A1:C1")
    ws.insert_columns([("x",)], at_col="B")
    assert "A1:D1" in ws.merged_ranges


def test_insert_columns_cf_sqref_shifts():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["C1"].value = 1
    ws.add_conditional_formatting("C1:C3", operator="greaterThan", formula="0")
    ws.insert_columns([("x",)], at_col="C")
    assert ws.conditional_formatting[0].cell_range == "D1:D3"


def test_insert_columns_requires_at_col_or_at_cell():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    with pytest.raises(ValueError, match="at_col or at_cell"):
        ws.insert_columns([("x",)])


def test_insert_columns_round_trip(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "A"
    ws["B1"].value = "B"
    ws.insert_columns([("X",)], at_col="B")
    out = tmp_path / "insert_cols.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    assert wb2["Sheet1"]["A1"].value == "A"
    assert wb2["Sheet1"]["B1"].value == "X"
    assert wb2["Sheet1"]["C1"].value == "B"

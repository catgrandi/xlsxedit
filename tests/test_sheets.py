from __future__ import annotations

from pathlib import Path

import pytest

from xlsxedit import Workbook


def test_add_worksheet():
    wb = Workbook.create()
    ws2 = wb.add_worksheet("Sheet2")
    assert wb.sheetnames == ["Sheet1", "Sheet2"]
    assert ws2.name == "Sheet2"
    assert list(ws2.cells) == []


def test_add_worksheet_duplicate_name_raises():
    wb = Workbook.create()
    with pytest.raises(ValueError, match="already exists"):
        wb.add_worksheet("Sheet1")


def test_add_worksheet_save_roundtrip(tmp_path: Path):
    out = tmp_path / "two_sheets.xlsx"
    wb = Workbook.create()
    ws2 = wb.add_worksheet("Report")
    ws2["A1"].value = "Title"
    wb.save(out)

    wb2 = Workbook.open(out)
    assert wb2.sheetnames == ["Sheet1", "Report"]
    assert wb2["Report"]["A1"].value == "Title"
    assert list(wb2["Sheet1"].cells) == []

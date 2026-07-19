from __future__ import annotations

import zipfile
from pathlib import Path

from xlsxedit import Workbook
from xlsxedit.api import _default_xlsx_path


def test_default_template_path_exists():
    path = _default_xlsx_path()
    assert path.is_file()
    assert path.name == "default.xlsx"


def test_create_blank_workbook():
    wb = Workbook.create()
    assert wb.sheetnames == ["Sheet1"]
    assert list(wb["Sheet1"].cells) == []
    assert wb.orphan_partnames == ()


def test_constructor_blank_workbook():
    wb = Workbook()
    assert wb.sheetnames == ["Sheet1"]
    assert list(wb["Sheet1"].cells) == []


def test_constructor_open_workbook(book1_path: Path):
    wb = Workbook(book1_path)
    assert "Sheet1" in wb.sheetnames
    assert wb["Sheet1"]["A1"].value == "this is a1"


def test_create_save_roundtrip(tmp_path: Path):
    out = tmp_path / "new.xlsx"
    wb = Workbook.create()
    wb.save(out)

    assert out.is_file()
    with zipfile.ZipFile(out) as z:
        names = set(z.namelist())
        assert "xl/workbook.xml" in names
        assert "xl/worksheets/sheet1.xml" in names
        assert "xl/styles.xml" in names
        assert "xl/theme/theme1.xml" in names
        assert "xl/sharedStrings.xml" in names

    wb2 = Workbook.open(out)
    assert wb2.sheetnames == ["Sheet1"]
    assert list(wb2["Sheet1"].cells) == []

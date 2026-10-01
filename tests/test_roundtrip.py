from __future__ import annotations

import zipfile
from pathlib import Path

from xlsxedit import Workbook


def test_open_sheet_names(book1_path: Path):
    wb = Workbook.open(book1_path)
    assert wb.sheetnames == ["Sheet1", "test"]


def test_roundtrip_preserves_opaque_parts(book1_path: Path, tmp_path: Path):
    out = tmp_path / "out.xlsx"
    wb = Workbook.open(book1_path)
    wb.save(out)

    with zipfile.ZipFile(book1_path) as orig, zipfile.ZipFile(out) as saved:
        assert set(orig.namelist()) == set(saved.namelist())
        assert orig.read("xl/styles.xml") == saved.read("xl/styles.xml")
        assert orig.read("xl/theme/theme1.xml") == saved.read("xl/theme/theme1.xml")


def test_open_unpacked_directory(book1_path: Path, tmp_path: Path):
    book1_dir = tmp_path / "Book1"
    with zipfile.ZipFile(book1_path) as z:
        z.extractall(book1_dir)
    assert (book1_dir / "[Content_Types].xml").is_file()

    wb = Workbook.open(book1_dir)
    assert wb.sheetnames == Workbook.open(book1_path).sheetnames


def test_numeric_cells_unchanged_on_save(book1_path: Path, tmp_path: Path):
    out = tmp_path / "out.xlsx"
    wb = Workbook.open(book1_path)
    before = {c.address: c.value for c in wb["Sheet1"].cells}
    wb.save(out)
    after = {c.address: c.value for c in Workbook.open(out)["Sheet1"].cells}
    assert before == after
    assert after["A5"] == 100

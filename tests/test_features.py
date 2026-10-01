"""Tests for Phase 4+ features."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pytest

from xlsxedit import Workbook
from tests.conftest import INSPECT_FIXTURES


def test_datetime_read_write(tmp_path: Path):
    wb = Workbook.open(INSPECT_FIXTURES["DifferentCellTypes"])
    b5 = wb["Sheet1"]["B5"]
    assert isinstance(b5.value, datetime)
    assert b5.value.day == 7 and b5.value.month == 8 and b5.value.year == 2025
    b6 = wb["Sheet1"]["B6"]
    assert b6.value.day == 21
    out = tmp_path / "dates.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    assert isinstance(wb2["Sheet1"]["B5"].value, datetime)


def test_replace_value_type_date(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "{ship}"
    ws["A1"].style  # ensure style exists on template cell
    n = wb.replace("{ship}", "2025-08-07", value_type="date")
    assert n == 1
    out = tmp_path / "d.xlsx"
    wb.save(out)


def test_merge_write_redirects_to_anchor(tmp_path: Path):
    wb = Workbook.open(INSPECT_FIXTURES["MergedCells"])
    wb["Sheet1"]["B1"].value = "via B1"
    assert wb["Sheet1"]["A1"].value == "via B1"
    out = tmp_path / "merged.xlsx"
    wb.save(out)


def test_cell_style_read():
    wb = Workbook.open(INSPECT_FIXTURES["DifferentCellTypes"])
    b5 = wb["Sheet1"]["B5"]
    assert b5.style.is_date
    b2 = wb["Sheet1"]["B2"]
    assert b2.style.is_percent


def test_rename_worksheet(tmp_path: Path):
    wb = Workbook.open(INSPECT_FIXTURES["TestBook"])
    wb.rename_worksheet("Sheet2", "Renamed")
    assert "Renamed" in wb.sheetnames
    out = tmp_path / "renamed.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    assert "Renamed" in wb2.sheetnames


def test_remove_worksheet(tmp_path: Path):
    wb = Workbook.open(INSPECT_FIXTURES["TestBook"])
    wb.remove_worksheet("Sheet2")
    assert wb.sheetnames == ["Sheet1"]
    with pytest.raises(ValueError):
        wb.remove_worksheet("Sheet1")


def test_images_list_and_replace(tmp_path: Path):
    wb = Workbook.open(INSPECT_FIXTURES["images"])
    pics = wb["Sheet1"].images
    assert len(pics) >= 1
    assert pics[0].anchor
    assert pics[0].media_path


def test_charts_and_tables_read():
    wb = Workbook.open(INSPECT_FIXTURES["ChartsAndTables"])
    assert len(wb["bar chart"].charts) >= 1
    assert wb["bar chart"].charts[0].name
    tbl_ws = wb["Table"]
    assert len(tbl_ws.tables) >= 1
    assert tbl_ws.tables[0].columns


def test_add_image_on_create(tmp_path: Path, assets_dir: Path):
    jpg = assets_dir / "coco-happy-swiss-nature.jpg"
    assert jpg.is_file()
    wb = Workbook.create()
    pic = wb["Sheet1"].add_image(jpg, anchor="C4", name="testpic")
    assert pic.name == "testpic"
    assert pic.height > pic.width
    out = tmp_path / "with_img.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    assert len(wb2["Sheet1"].images) == 1


def test_hyperlink_fixture_external_and_internal():
    wb = Workbook.open(INSPECT_FIXTURES["hyperlink"])
    assert wb["Sheet1"]["B1"].hyperlink.url == "https://jonasruilong.com/"
    assert wb["Sheet1"]["B1"].hyperlink.location is None
    assert wb["Sheet1"]["B2"].hyperlink.url is None
    assert wb["Sheet1"]["B2"].hyperlink.location == "Sheet2!B1"
    assert wb["Sheet1"]["B2"].hyperlink.display == "link to other sheet"


def test_conditional_formatting_sheet2_cell_is_dxf():
    wb = Workbook.open(INSPECT_FIXTURES["ConditionalFormatting"])
    ws = wb["Sheet2"]
    assert len(ws.conditional_formatting) >= 1
    rule = ws.conditional_formatting[0].rules[0]
    assert rule.type == "cellIs"
    assert rule.operator == "greaterThan"
    assert rule._element.get("dxfId") == "0"
    assert ws.conditional_formatting[0].cell_range == "B6:B7"


def test_custom_cell_size_with_added_image_roundtrip(assets_dir: Path, tmp_path: Path):
    wb = Workbook.open(INSPECT_FIXTURES["CustomCellSize"])
    wb["Sheet1"].add_image(assets_dir / "coco-happy-swiss-nature.jpg", anchor="E2")
    out = tmp_path / "sized_with_image.xlsx"
    wb.save(out)

    ws = Workbook.open(out)["Sheet1"]
    assert len(ws.images) == 1
    assert ws.column_dimensions["C"].width == 32.0
    assert ws.row_dimensions[4].height == 47.0

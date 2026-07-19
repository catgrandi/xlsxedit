"""Tests for cell font/fill style writes."""

from __future__ import annotations

from pathlib import Path

import pytest

from xlsxedit import Workbook
from tests.conftest import INSPECT_FIXTURES


def test_text_style_fixture_read():
    wb = Workbook.open(INSPECT_FIXTURES["TextSytle"])
    ws = wb["Sheet1"]
    assert ws["A1"].style.bold
    assert ws["A3"].style.font_color == "FFFF0000"
    assert ws["A4"].style.font_size == 16
    assert ws["A5"].style.bg_color == "FFFFFF00"
    assert ws["A6"].style.horizontal_align == "center"
    assert ws["A7"].style.italic
    assert ws["A8"].style.underline
    assert ws["A9"].style.font_name == "Arial"


def test_text_style_fixture_round_trip(tmp_path: Path):
    wb = Workbook.open(INSPECT_FIXTURES["TextSytle"])
    out = tmp_path / "textstyle.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    assert wb2["Sheet1"]["A1"].style.bold
    assert wb2["Sheet1"]["A5"].style.bg_color == "FFFFFF00"


def test_apply_style_write(tmp_path: Path):
    wb = Workbook.create()
    cell = wb["Sheet1"]["A1"]
    cell.value = "styled"
    cell.apply_style(bold=True, font_color="FF0000", bg_color="FFFF00", font_size=14)
    assert cell.style.bold
    assert cell.style.font_color == "FFFF0000"
    assert cell.style.bg_color == "FFFFFF00"
    assert cell.style.font_size == 14
    wb.save(tmp_path / "styled.xlsx")
    wb2 = Workbook.open(tmp_path / "styled.xlsx")
    assert wb2["Sheet1"]["A1"].style.bold


def test_apply_style_live_part_round_trip(tmp_path: Path):
    """Regression: styles live on StylesPart.element and serialize on save."""
    wb = Workbook.open(INSPECT_FIXTURES["TextSytle"])
    ws = wb["Sheet1"]
    ws["B1"].value = "new"
    ws["B1"].apply_style(bold=True, bg_color="00FF00")
    out = tmp_path / "live.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    assert wb2["Sheet1"]["B1"].style.bold
    assert wb2["Sheet1"]["B1"].style.bg_color == "FF00FF00"
    # Pre-existing styles untouched.
    assert wb2["Sheet1"]["A1"].style.bold


def test_colored_conditional_formatting(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    for i, val in enumerate([10, 50, 120], start=2):
        ws[f"B{i}"].value = val
    ws.add_conditional_formatting(
        "B2:B4",
        operator="greaterThan",
        formula="100",
        bg_color="FFC6EFCE",
        font_color="FF006100",
    )
    wb.save(tmp_path / "cf.xlsx")
    wb2 = Workbook.open(tmp_path / "cf.xlsx")
    rule = wb2["Sheet1"].conditional_formatting[0].rules[0]
    assert rule.type == "cellIs"
    assert rule._element.get("dxfId") is not None


def test_default_conditional_formatting_has_dxf(tmp_path: Path):
    from lxml import etree

    from xlsxedit.opc.constants import SML_NS

    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = 5
    ws.add_conditional_formatting("A1", operator="greaterThan", formula="0")
    wb.save(tmp_path / "cf_default.xlsx")
    wb2 = Workbook.open(tmp_path / "cf_default.xlsx")
    rule = wb2["Sheet1"].conditional_formatting[0].rules[0]
    assert rule.type == "cellIs"
    assert rule._element.get("dxfId") is not None

    styles = etree.fromstring(wb2._workbook_part.styles_part.blob)
    ns = f"{{{SML_NS}}}"
    dxf_id = int(rule._element.get("dxfId"))
    dxf = styles.find(f"{ns}dxfs")[dxf_id]
    bg = dxf.find(f".//{ns}bgColor")
    assert bg is not None
    assert bg.get("rgb") == "FFC6EFCE"


def test_color_scale_formatting(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    for i, val in enumerate([15, 45, 72, 91, 33], start=2):
        ws[f"B{i}"].value = val
    ws.add_color_scale_formatting("B2:B6")
    wb.save(tmp_path / "colorscale.xlsx")
    wb2 = Workbook.open(tmp_path / "colorscale.xlsx")
    assert wb2["Sheet1"].conditional_formatting[0].rules[0].type == "colorScale"

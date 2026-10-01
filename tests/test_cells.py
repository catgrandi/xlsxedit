from __future__ import annotations

import zipfile
from pathlib import Path

from xlsxedit import Workbook
from xlsxedit.oxml.parser import parse_xml
from xlsxedit.opc.constants import SML_NS

_C = f"{{{SML_NS}}}c"


def test_create_set_string_and_number(tmp_path: Path):
    out = tmp_path / "new.xlsx"
    wb = Workbook.create()
    wb["Sheet1"]["A1"].value = "hello"
    wb["Sheet1"]["A2"].value = 888
    wb.save(out)

    wb2 = Workbook.open(out)
    assert wb2["Sheet1"]["A1"].value == "hello"
    assert wb2["Sheet1"]["A2"].value == 888

    with zipfile.ZipFile(out) as z:
        sheet = parse_xml(z.read("xl/worksheets/sheet1.xml"))
    a2 = next(c for c in sheet.iter(_C) if c.get("r") == "A2")
    assert a2.get("t") is None
    assert a2.find(f"{{{SML_NS}}}v").text == "888"


def test_set_formula(tmp_path: Path):
    out = tmp_path / "formula.xlsx"
    wb = Workbook.create()
    wb["Sheet1"]["A1"].value = 10
    wb["Sheet1"]["A2"].formula = "=A1*2"
    wb.save(out)

    wb2 = Workbook.open(out)
    assert wb2["Sheet1"]["A2"].formula == "=A1*2"


def _sheet1_cell(path: Path, address: str):
    with zipfile.ZipFile(path) as z:
        sheet = parse_xml(z.read("xl/worksheets/sheet1.xml"))
    return next(c for c in sheet.iter(_C) if c.get("r") == address)


def test_typed_number_replace_emits_no_type_attribute(tmp_path: Path):
    template = tmp_path / "template.xlsx"
    wb = Workbook.create()
    wb["Sheet1"]["A2"].value = "{number1}"
    wb.save(template)
    # The placeholder is a typed text cell; the replace must drop that type.
    assert _sheet1_cell(template, "A2").get("t") is not None

    out = tmp_path / "filled.xlsx"
    wb = Workbook.open(template)
    n = wb.replace("{number1}", 888, value_type="number")
    assert n == 1
    wb.save(out)

    wb2 = Workbook.open(out)
    assert wb2["Sheet1"]["A2"].value == 888
    assert _sheet1_cell(out, "A2").get("t") is None


def test_style_preserved_on_value_set(book1_path: Path, tmp_path: Path):
    wb = Workbook.open(book1_path)
    b2 = next(c for c in wb["Sheet1"].cells if c.address == "B2")
    style_before = b2.style_index
    b2.value = "updated text"
    out = tmp_path / "styled.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    b2_after = next(c for c in wb2["Sheet1"].cells if c.address == "B2")
    assert b2_after.style_index == style_before
    assert b2_after.value == "updated text"

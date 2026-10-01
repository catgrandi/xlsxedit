"""Round-trip and XML structural checks for the API smoke workbook."""

from __future__ import annotations

import zipfile
from datetime import datetime
from pathlib import Path

from lxml import etree

from xlsxedit import Workbook
from xlsxedit.opc.constants import SML_NS
from tests.smoke_workbook import build_smoke_workbook


def _assert_hyperlinks_after_sheet_data(blob: bytes) -> None:
    ws = etree.fromstring(blob)
    ns = f"{{{SML_NS}}}"
    sheet_data = ws.find(f"{ns}sheetData")
    hyperlinks = ws.find(f"{ns}hyperlinks")
    if hyperlinks is not None:
        assert sheet_data is not None
        assert list(ws).index(hyperlinks) > list(ws).index(sheet_data)


def _assert_styles_xml_order(blob: bytes) -> None:
    styles = etree.fromstring(blob)
    fonts = styles.find(f"{{{SML_NS}}}fonts")
    num_fmts = styles.find(f"{{{SML_NS}}}numFmts")
    if num_fmts is not None:
        assert list(styles).index(num_fmts) < list(styles).index(fonts)


def _assert_content_types(path: Path) -> None:
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        assert "xl/styles.xml" in names
        assert "xl/workbook.xml" in names
        assert any(n.startswith("xl/worksheets/") for n in names)
        assert "xl/drawings/drawing1.xml" in names


def _assert_chart_xml_valid(wb: Workbook, sheet_name: str = "Chart") -> None:
    from xlsxedit.drawing import CHART_NS

    chart = wb[sheet_name].charts[0]
    chart_part = chart._chart_part
    assert chart_part is not None
    root = etree.fromstring(chart_part.blob)
    ns = f"{{{CHART_NS}}}"

    for f_elm in root.iter(f"{ns}f"):
        formula = f_elm.text or ""
        if ":" in formula and "!" in formula:
            assert formula.count("!") == 1, formula

    title = root.find(f".//{ns}chart/{ns}title")
    assert title is not None
    assert etree.QName(title[0]).localname == "tx"
    assert title.find(f"{ns}tx/{ns}rich") is not None

    str_refs = list(root.iter(f"{ns}strRef"))
    assert len(str_refs) == 2
    assert str_refs[1].find(f"{ns}f").text == f"{sheet_name}!$A$2:$A$3"
    assert str_refs[1].find(f"{ns}strCache/{ns}ptCount").get("val") == "2"

    num_ref = root.find(f".//{ns}numRef")
    assert num_ref.find(f"{ns}f").text == f"{sheet_name}!$B$2:$B$3"
    assert num_ref.find(f"{ns}numCache/{ns}ptCount").get("val") == "2"


def _assert_cell_is_has_dxf(wb: Workbook, sheet_name: str = "CF") -> None:
    cf_blocks = wb[sheet_name].conditional_formatting
    cell_is = next(
        (rule for block in cf_blocks for rule in block.rules if rule.type == "cellIs"),
        None,
    )
    assert cell_is is not None
    assert cell_is._element.get("dxfId") is not None

    styles_part = wb._workbook_part.styles_part
    assert styles_part is not None
    styles = etree.fromstring(styles_part.blob)
    dxfs = styles.find(f"{{{SML_NS}}}dxfs")
    assert dxfs is not None
    assert int(dxfs.get("count", "0")) > 0


def test_smoke_workbook_round_trip_and_xml(tmp_path: Path) -> None:
    out = tmp_path / "api_smoke.xlsx"
    wb = build_smoke_workbook()
    wb.save(out)

    assert out.is_file()
    _assert_content_types(out)

    wb2 = Workbook.open(out)
    expected_sheets = [
        "Index", "Values", "Layout", "Styles", "Formulas", "SAR",
        "Links", "CF", "Chart", "Table", "Photos",
    ]
    assert wb2.sheetnames == expected_sheets
    assert "TempRemove" not in wb2.sheetnames

    values = wb2["Values"]
    assert values["B2"].value == "hello"
    assert values["B3"].value == 42
    assert values["B4"].value == 3.14
    assert values["B5"].value is True
    assert isinstance(values["B6"].value, datetime)
    assert values["B7"].style.num_format == "$#,##0.00"
    assert values["B8"].style.num_format == "0.00%"
    assert values["B9"].value is None

    layout = wb2["Layout"]
    assert layout["A1"].value == "written via B1"
    assert layout.merged_ranges == []

    assert wb2["SAR"]["B1"].value == "Acme"
    assert wb2["SAR"]["B2"].value == 7
    assert isinstance(wb2["SAR"]["B3"].value, datetime)
    assert wb2["SAR"]["B4"].value == "Item Widget ok"
    assert wb2["SAR"]["B5"].value == "partial new text"

    assert wb2["Links"]["A1"].hyperlink.url == "https://xlsxedit.jonasruilong.com"
    _assert_hyperlinks_after_sheet_data(wb2["Links"]._part.blob)
    assert wb2["Formulas"]["C1"].formula == "=A1+B1"
    assert len(wb2["CF"].conditional_formatting) >= 2
    _assert_cell_is_has_dxf(wb2, "CF")
    assert len(wb2["Chart"].charts) >= 1
    assert wb2["Chart"].charts[0].title == "Sales updated"
    _assert_chart_xml_valid(wb2, "Chart")
    assert wb2["Table"].tables[0].ref == "A1:B4"
    assert len(wb2["Photos"].images) >= 2

    styles_part = wb2._workbook_part.styles_part
    assert styles_part is not None
    _assert_styles_xml_order(styles_part.blob)

    print("Smoke workbook OK")

"""Worksheet child element order must follow OOXML sequence."""

from __future__ import annotations

from pathlib import Path

import pytest
from lxml import etree

from xlsxedit import Workbook
from xlsxedit.opc.constants import OFFICE_REL_NS, SML_NS
from xlsxedit.worksheet_order import (
    _WS_CHILD_RANK,
    insert_worksheet_child,
    reposition_worksheet_child,
)

ASSETS = Path(__file__).resolve().parents[1] / "assets"


def _child_names(ws) -> list[str]:
    return [etree.QName(c).localname for c in ws._part.element]


def _assert_order_ok(names: list[str]) -> None:
    ranks = [_WS_CHILD_RANK[n] for n in names if n in _WS_CHILD_RANK]
    assert ranks == sorted(ranks), f"out of order: {names}"


def test_table_then_cf_then_drawing_order(tmp_path: Path):
    """Regression: tableParts before CF/drawing made Excel repair the file."""
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Day"
    ws["B1"].value = "N"
    ws["A2"].value = "Mon"
    ws["B2"].value = 1
    ws.merge_cells("A1:B1")
    ws.add_table("A1:B2", ["Day", "N"], name="T1")
    ws.add_conditional_formatting("B2", operator="greaterThan", formula="0")
    ws.add_color_scale_formatting("B2:B2")
    ws.add_chart("bar", anchor="D2", data_range="A1:B2", title="T")
    ws["A5"].value = "link"
    ws["A5"].hyperlink.url = "https://example.com"

    out = tmp_path / "order.xlsx"
    wb.save(out)

    wb2 = Workbook.open(out)
    names = _child_names(wb2["Sheet1"])
    _assert_order_ok(names)
    assert "tableParts" in names
    assert "conditionalFormatting" in names
    assert "drawing" in names
    assert "hyperlinks" in names
    assert names.index("mergeCells") < names.index("conditionalFormatting")
    assert names.index("conditionalFormatting") < names.index("hyperlinks")
    assert names.index("hyperlinks") < names.index("drawing")
    assert names.index("drawing") < names.index("tableParts")


def test_new_children_go_before_legacy_drawing(tmp_path: Path):
    """Regression: on a sheet with comments or form controls, new blocks landed
    after ``<legacyDrawing>``, which Excel repairs."""
    wb = Workbook.create()
    ws = wb["Sheet1"]
    root = ws._part.element
    for name, r_id in (("legacyDrawing", "rId90"), ("legacyDrawingHF", "rId91")):
        elm = etree.SubElement(root, f"{{{SML_NS}}}{name}")
        elm.set(f"{{{OFFICE_REL_NS}}}id", r_id)

    ws["A1"].value = "Day"
    ws["B1"].value = "N"
    ws["A2"].value = "Mon"
    ws["B2"].value = 1
    ws.merge_cells("C1:D1")
    ws.add_conditional_formatting("B2", operator="greaterThan", formula="0")
    ws["A5"].value = "link"
    ws["A5"].hyperlink.url = "https://example.com"
    ws.add_image(ASSETS / "coco-happy-swiss-nature.jpg", anchor="F2")
    ws.add_table("A1:B2", ["Day", "N"], name="T1")

    out = tmp_path / "legacy.xlsx"
    wb.save(out)

    names = _child_names(Workbook.open(out)["Sheet1"])
    _assert_order_ok(names)
    assert names[-5:] == [
        "hyperlinks",
        "drawing",
        "legacyDrawing",
        "legacyDrawingHF",
        "tableParts",
    ]
    assert names.index("mergeCells") < names.index("conditionalFormatting")


def test_unknown_child_names_leave_the_tree_alone():
    wb = Workbook.create()
    root = wb["Sheet1"]._part.element
    foreign = etree.SubElement(root, "{urn:example:ext}note")
    before = list(root)

    with pytest.raises(ValueError, match="bogus"):
        insert_worksheet_child(root, etree.Element(f"{{{SML_NS}}}bogus"))
    reposition_worksheet_child(root, foreign)

    assert list(root) == before


def test_reposition_moves_only_an_out_of_order_child():
    wb = Workbook.create()
    root = wb["Sheet1"]._part.element
    hyperlinks = etree.SubElement(root, f"{{{SML_NS}}}hyperlinks")
    root.append(etree.Comment("kept in place"))
    etree.SubElement(root, f"{{{SML_NS}}}pageMargins")
    before = list(root)

    reposition_worksheet_child(root, hyperlinks)
    assert list(root) == before

    table_parts = etree.Element(f"{{{SML_NS}}}tableParts")
    root.insert(0, table_parts)
    reposition_worksheet_child(root, table_parts)
    assert list(root) == before + [table_parts]

"""Worksheet child element order must follow OOXML sequence."""

from __future__ import annotations

from pathlib import Path

from lxml import etree

from xlsxedit import Workbook
from xlsxedit.opc.constants import SML_NS
from xlsxedit.worksheet_order import _WS_CHILD_RANK


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

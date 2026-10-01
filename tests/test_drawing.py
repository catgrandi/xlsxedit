"""Drawing and chart parts: one live tree per part."""

from __future__ import annotations

import io
from pathlib import Path

import pytest

from tests.conftest import INSPECT_FIXTURES
from tests.preservation import assert_preserved, read_pkg, worksheet_members
from xlsxedit import Workbook
from xlsxedit.drawing import drawing_parts_for_worksheet
from xlsxedit.parts import ChartPart, DrawingPart

PNG = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
    b"\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc\xf8\x0f\x00"
    b"\x00\x01\x01\x00\x05\x18\xd8N\x00\x00\x00\x00IEND\xaeB`\x82"
)


@pytest.fixture
def png(tmp_path: Path) -> Path:
    path = tmp_path / "dot.png"
    path.write_bytes(PNG)
    return path


def _reopen(wb: Workbook) -> Workbook:
    buf = io.BytesIO()
    wb.save(buf)
    return Workbook.open(io.BytesIO(buf.getvalue()))


def _drawing_payload(member: str) -> bool:
    """A drawing, chart or media part; their ``.rels`` are regenerated on every save."""
    return member.startswith(("xl/drawings/", "xl/charts/", "xl/media/")) and "/_rels/" not in member


def _chart_sheet(wb: Workbook, png: Path):
    ws = wb["Sheet1"]
    ws["A1"].value = "Item"
    ws["B1"].value = "Qty"
    ws["A2"].value = "Apples"
    ws["B2"].value = 3
    ws.add_chart("bar", anchor="D2", data_range="A1:B2", title="Sales")
    ws.add_image(png, anchor="H2", name="logo")
    return ws


# --- one live tree per part -------------------------------------------------


def test_registered_part_classes():
    wb = Workbook.open(INSPECT_FIXTURES["ChartsAndTables"])
    ws = wb["bar chart"]
    assert isinstance(drawing_parts_for_worksheet(ws)[0], DrawingPart)
    assert isinstance(ws.charts[0]._chart_part, ChartPart)


def test_chart_and_picture_edits_on_one_sheet_both_survive(png: Path):
    wb = Workbook.create()
    _chart_sheet(wb, png)
    wb = _reopen(wb)

    ws = wb["Sheet1"]
    ch = ws.charts[0]
    pic = ws.images[0]
    ch.anchor = "Z50"
    pic.anchor = "B40"

    ws2 = _reopen(wb)["Sheet1"]
    assert ws2.charts[0].anchor == "Z50"
    assert ws2.images[0].anchor == "B40"


def test_proxies_share_the_live_tree(png: Path):
    wb = Workbook.create()
    ws = _chart_sheet(wb, png)
    first, second = ws.images[0], ws.images[0]
    first.name = "renamed"
    assert second.name == "renamed"
    ws.charts[0].offset_x = 4
    ws.charts[0].name = "chart renamed"
    assert ws.charts[0].offset_x == 4
    ws2 = _reopen(wb)["Sheet1"]
    assert ws2.images[0].name == "renamed"
    assert (ws2.charts[0].name, ws2.charts[0].offset_x) == ("chart renamed", 4)


def test_reading_does_not_reserialise_drawing_or_chart_parts():
    path = INSPECT_FIXTURES["ChartsAndTables"]
    wb = Workbook.open(path)
    for ws in wb.worksheets:
        for chart in ws.charts:
            assert chart.anchor and chart.name
            chart.title
    before, after = read_pkg(path), read_pkg(wb)
    for member in filter(_drawing_payload, before):
        assert after[member] == before[member], member


@pytest.mark.parametrize("fixture", ["ChartsAndTables", "images"])
def test_unrelated_cell_edit_keeps_drawing_and_chart_parts_byte_identical(fixture: str):
    path = INSPECT_FIXTURES[fixture]
    before = read_pkg(path)
    wb = Workbook.open(path)
    ws = wb.worksheets[0]
    ws.images, ws.charts  # parse the drawing part
    ws["Z99"].value = 5

    report = assert_preserved(before, wb, expected_changed={worksheet_members(before)[ws.name]})
    after = read_pkg(wb)
    for member in filter(_drawing_payload, before):
        assert member not in report.reserialised
        assert after[member] == before[member], member


def test_insert_rows_on_another_sheet_keeps_drawing_and_chart_parts_byte_identical():
    path = INSPECT_FIXTURES["ChartsAndTables"]
    before = read_pkg(path)
    wb = Workbook.open(path)
    for ws in wb.worksheets:
        ws.charts, ws.images
    wb["Table"].insert_rows([[1, 2]], at_row=40)

    report = assert_preserved(before, wb, expected_changed={worksheet_members(before)["Table"]})
    after = read_pkg(wb)
    drawing_members = list(filter(_drawing_payload, before))
    assert drawing_members
    for member in drawing_members:
        assert member not in report.reserialised
        assert after[member] == before[member], member


def test_failed_series_formula_leaves_the_chart_unchanged(png: Path):
    wb = Workbook.create()
    ws = _chart_sheet(wb, png)
    chart = ws.charts[0]
    before = read_pkg(wb)["xl/charts/chart1.xml"]
    with pytest.raises(ValueError, match="invalid chart formula"):
        chart.set_series_formula(2, "B2:B3", ws)
    assert read_pkg(wb)["xl/charts/chart1.xml"] == before


def test_chart_title_edit_changes_only_its_chart_part():
    path = INSPECT_FIXTURES["ChartsAndTables"]
    wb = Workbook.open(path)
    chart = wb["bar chart"].charts[0]
    chart.title = "Renamed"
    assert_preserved(path, wb, expected_changed={chart.partname.lstrip("/")})
    assert _reopen(wb)["bar chart"].charts[0].title == "Renamed"

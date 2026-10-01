"""Style allocation: dedupe, growth bounds, font building and effective styles."""

from __future__ import annotations

from lxml import etree

from xlsxedit import Workbook
from xlsxedit.opc.constants import SML_NS
from xlsxedit.styles import Styles
from tests.conftest import INSPECT_FIXTURES
from tests.preservation import assert_preserved, read_pkg
from tests.schema_validate import assert_valid_package

_NS = f"{{{SML_NS}}}"


def _stylesheet(wb: Workbook) -> etree._Element:
    return wb._workbook_part.styles_part.element


def _counts(wb: Workbook) -> dict[str, int]:
    stylesheet = _stylesheet(wb)
    return {
        name: len(stylesheet.find(f"{_NS}{name}"))
        for name in ("cellXfs", "fonts", "fills", "dxfs")
    }


# --- growth bounds -----------------------------------------------------------


def test_template_rows_with_column_styles_allocate_per_template_column():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    for col in "ABCD":
        ws[f"{col}2"].apply_style(bg_color="FFFFFFFF")
        ws[f"{col}3"].apply_style(bg_color="FFF2F2F2")
    before = _counts(wb)
    rows = [(f"item {i}", i, i * 1.5, i % 7) for i in range(5_000)]
    column_styles = [{}, {"num_format": "0"}, {"bold": True, "num_format": "0.00"}, {}]

    ws.write_rows(rows, at_row=4, template_rows=[2, 3], column_styles=column_styles)
    after = _counts(wb)

    assert after["cellXfs"] < 50
    assert after["cellXfs"] - before["cellXfs"] <= 2 * 2  # template rows x styled columns
    assert after["fonts"] - before["fonts"] <= 1
    assert after["fills"] == before["fills"]
    assert ws["B4"].style.num_format == "0"
    assert ws["C5"].style.bold
    assert ws["C4"].style.bg_color == "FFFFFFFF"
    assert ws["C5"].style.bg_color == "FFF2F2F2"

    ws.write_rows(rows, at_row=4, template_rows=[2, 3], column_styles=column_styles)
    assert _counts(wb) == after
    assert_valid_package(wb)


def test_bulk_export_allocates_once_per_template_row_and_column(monkeypatch):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A2"].apply_style(bg_color="FFFFFFFF")
    ws["A3"].apply_style(bg_color="FFF2F2F2")
    calls = []
    allocate = Styles.allocate_cell_style

    def counting(self, **kwargs):
        calls.append(kwargs)
        return allocate(self, **kwargs)

    monkeypatch.setattr(Styles, "allocate_cell_style", counting)
    rows = [(i, i * 1.5) for i in range(1_000)]
    ws.write_rows(rows, at_row=4, template_rows=[2, 3], column_styles=[{"bold": True}, {}])

    assert len(calls) == 2


def test_identical_row_styles_calls_allocate_once():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    kwargs = {
        "row_styles": [{"bg_color": "FFFFFFFF"}, {"bg_color": "FF9DC3E6", "bold": True}],
        "column_styles": [{}, {"num_format": "$#,##0.00"}],
    }
    rows = [(f"r{i}", i * 2.5) for i in range(100)]

    ws.write_rows(rows, at_row=1, **kwargs)
    first = _counts(wb)
    ws.write_rows(rows, at_row=1, **kwargs)
    ws.write_rows(rows, at_row=200, **kwargs)

    assert _counts(wb) == first
    assert_valid_package(wb)


def test_identical_apply_style_calls_add_one_fill():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    before = _counts(wb)

    for row in range(1, 301):
        ws[f"A{row}"].apply_style(bg_color="FFFFFF00")

    after = _counts(wb)
    assert after["fills"] == before["fills"] + 1
    assert after["cellXfs"] == before["cellXfs"] + 1
    assert ws["A300"].style_index == ws["A1"].style_index


# --- dedupe ------------------------------------------------------------------


def test_apply_style_reuses_an_identical_existing_xf():
    """TextSytle's xf 1 is the default font in bold, so bolding A2 adds nothing."""
    wb = Workbook.open(INSPECT_FIXTURES["TextSytle"])
    before = read_pkg(wb)
    ws = wb["Sheet1"]

    ws["A2"].apply_style(bold=True)

    assert ws["A2"].style_index == "1"
    assert not wb.styles.dirty
    assert_preserved(before, wb, expected_changed={"xl/worksheets/sheet1.xml"})


def test_existing_duplicates_keep_their_indices():
    red = (
        '<fill><patternFill patternType="solid">'
        '<fgColor rgb="FFFF0000"/><bgColor indexed="64"/></patternFill></fill>'
    )
    stylesheet = etree.fromstring(
        f'<styleSheet xmlns="{SML_NS}"><fills count="3">'
        f'<fill><patternFill patternType="none"/></fill>{red}{red}</fills>'
        '<cellXfs count="2"><xf numFmtId="0" fillId="1"/><xf numFmtId="0" fillId="2"/>'
        "</cellXfs></styleSheet>"
    )
    styles = Styles(stylesheet)

    assert styles.ensure_fill("FF0000") == 1
    assert styles.clone_xf(1) == 1
    assert [xf.get("fillId") for xf in stylesheet.find(f"{_NS}cellXfs")] == ["1", "2"]
    assert len(stylesheet.find(f"{_NS}fills")) == 3


def test_entry_edited_in_place_is_not_reused():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].apply_style(bg_color="FF00FF00")
    edited = int(ws["A1"].style_index)
    xf = _stylesheet(wb).find(f"{_NS}cellXfs")[edited]
    etree.SubElement(xf, f"{_NS}alignment").set("horizontal", "right")

    ws["A2"].apply_style(bg_color="FF00FF00")

    assert int(ws["A2"].style_index) != edited
    assert ws["A2"].style.horizontal_align is None


def test_identical_conditional_formats_share_a_dxf():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    before = _counts(wb)

    for cell_range in ("A1:A5", "B1:B5", "C1:C5"):
        ws.add_conditional_formatting(cell_range, formula="10", bg_color="FFC6EFCE")

    assert _counts(wb)["dxfs"] == before["dxfs"] + 1
    dxf_ids = {cf.rules[0]._element.get("dxfId") for cf in ws.conditional_formatting}
    assert len(dxf_ids) == 1


def test_recreated_dxfs_restarts_numbering():
    wb = Workbook.open(INSPECT_FIXTURES["ConditionalFormatting"])
    stylesheet = _stylesheet(wb)
    stylesheet.remove(stylesheet.find(f"{_NS}dxfs"))

    ws = wb.worksheets[0]
    ws.add_conditional_formatting("H1", formula="0", bg_color="FFFFC7CE")

    added = [cf for cf in ws.conditional_formatting if cf.cell_range == "H1"]
    assert added[0].rules[0]._element.get("dxfId") == "0"
    assert stylesheet.find(f"{_NS}dxfs").get("count") == "1"
    assert_valid_package(wb)

"""Tests for bulk export APIs (write_rows, write_dataframe, large mode)."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pytest

from xlsxedit import Workbook
from xlsxedit.parts import WorksheetPart


class SimpleFrame:
    """Minimal DataFrame stand-in for tests without pandas."""

    def __init__(self, columns, rows):
        self.columns = columns
        self._rows = rows

    def itertuples(self, index=False):
        for row in self._rows:
            yield tuple(row)


def test_write_rows_numeric(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "H1"
    ws["B1"].value = "H2"
    ws["A2"].value = 1
    ws["B2"].value = 2
    ws["A2"].apply_style(bold=True)

    n = ws.write_rows(
        [(10, 20), (30, 40)],
        at_row=3,
        template_rows=2,
    )
    assert n == 2
    assert ws["A3"].value == 10
    assert ws["B4"].value == 40
    out = tmp_path / "bulk.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    assert wb2["Sheet1"]["A3"].value == 10
    assert wb2["Sheet1"]["B4"].value == 40


def test_write_dataframe_with_simple_frame(tmp_path: Path):
    wb = Workbook.create()
    df = SimpleFrame(["Item", "Qty"], [("Apple", 3), ("Pear", 5)])
    rows_written = wb.write_dataframe(df, at_cell="A1", template_rows=1)
    assert rows_written == 2
    ws = wb["Sheet1"]
    assert ws["A1"].value == "Item"
    assert ws["A2"].value == "Apple"
    assert ws["B3"].value == 5
    out = tmp_path / "df.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    assert wb2["Sheet1"]["A2"].value == "Apple"


def test_iter_rows():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "a"
    ws["B2"].value = 2
    rows = list(ws.iter_rows())
    assert len(rows) == 2
    assert rows[0][0] == 1
    assert rows[1][0] == 2


def test_clear_range(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "keep style"
    ws["A1"].apply_style(bold=True)
    ws["A2"].value = "remove"
    cleared = ws.clear_range("A2:A2")
    assert cleared == 1
    assert ws["A2"].value is None
    assert ws["A1"].style.bold


def test_expand_conditional_formatting(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.add_conditional_formatting("B2:B5", operator="greaterThan", formula="0")
    updated = ws.expand_conditional_formatting(100)
    assert updated == 1
    assert ws.conditional_formatting[0].cell_range == "B2:B100"


def test_expand_conditional_formatting_skips_single_cell():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.add_conditional_formatting("D7", operator="greaterThan", formula="80")
    ws.add_color_scale_formatting("C10:C12")
    updated = ws.expand_conditional_formatting(16)
    assert updated == 1  # only the multi-row color scale
    ranges = {b.cell_range for b in ws.conditional_formatting}
    assert "D7" in ranges
    assert "C10:C16" in ranges


def test_table_resize_preserves_header_when_header_false():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Item"
    ws["B1"].value = "Qty"
    for i in range(2, 9):
        ws[f"A{i}"].value = f"r{i}"
        ws[f"B{i}"].value = i
    ws.add_table("A1:B8", ["Item", "Qty"], name="TBody")
    rows = [(f"n{i}", i * 10) for i in range(1, 8)]
    wb.write_dataframe(
        rows,
        at_cell="A2",
        header=False,
        mode="overwrite",
        resize_table=True,
    )
    assert ws.tables[0].ref == "A1:B8"


def test_large_mode_deferred_worksheet_parse(tmp_path: Path):
    wb = Workbook.create()
    wb.add_worksheet("Data")
    wb["Sheet1"]["A1"].value = "s1"
    wb["Data"]["A1"].value = "s2"
    out = tmp_path / "two_sheets.xlsx"
    wb.save(out)

    wb2 = Workbook.open(out, large=True)
    assert wb2._large is True
    data_part = wb2["Data"]._part
    assert isinstance(data_part, WorksheetPart)
    assert data_part._element is None
    assert data_part.blob is data_part._orig_blob
    # Access triggers parse
    assert wb2["Data"]["A1"].value == "s2"
    assert data_part._element is not None


def test_export_to_template(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Item"
    ws["B1"].value = "Qty"
    template = tmp_path / "template.xlsx"
    wb.save(template)

    df = SimpleFrame(["Item", "Qty"], [("X", 1)])
    out = tmp_path / "filled.xlsx"
    result = Workbook.export_to_template(
        template,
        df,
        save_path=out,
        at_cell="A1",
        template_rows=1,
        header=False,
    )
    assert result is not None
    wb2 = Workbook.open(out)
    assert wb2["Sheet1"]["A1"].value == "X"


def test_table_resize_on_write_dataframe(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Item"
    ws["B1"].value = "Qty"
    ws["A2"].value = "Old"
    ws["B2"].value = 1
    table = ws.add_table("A1:B2", ["Item", "Qty"], name="T1")
    assert table.name == "T1"
    df = SimpleFrame(["Item", "Qty"], [("New", 9), ("More", 2)])
    wb.write_dataframe(df, at_cell="A1", template_rows=1, header=True, resize_table=True)
    assert ws.tables[0].ref == "A1:B3"


def test_write_rows_zebra_template_styles():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A5"].value = "a"
    ws["A5"].apply_style(bg_color="FFFFFFFF")
    ws["A6"].value = "b"
    ws["A6"].apply_style(bg_color="FFF2F2F2")
    s5 = ws["A5"].style_index
    s6 = ws["A6"].style_index
    assert s5 != s6
    ws.write_rows([("r1",), ("r2",), ("r3",)], at_row=7, template_rows=[5, 6])
    assert ws["A7"].style_index == s5
    assert ws["A8"].style_index == s6
    assert ws["A9"].style_index == s5


def test_inline_row_and_column_styles():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.write_rows(
        [("Item", 2, 9.5), ("Other", 1, 3.25)],
        at_row=2,
        row_styles=[{"bg_color": "FFFFFFFF"}, {"bg_color": "FFF2F2F2"}],
        column_styles=[{}, {"num_format": "0"}, {"num_format": "0.00"}],
    )
    assert ws["B2"].style_index != ws["B3"].style_index
    assert ws["C2"].style.num_format == "0.00"


def test_insert_rows_pushes_footer():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A12"].value = "footer"
    ws.insert_rows([("d1",), ("d2",), ("d3",)], at_row=6)
    assert ws["A15"].value == "footer"


def test_write_dataframe_insert_mode():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A10"].value = "below"
    df = SimpleFrame(["X"], [("a",), ("b",)])
    wb.write_dataframe(df, at_cell="A5", header=False, mode="insert")
    assert ws["A5"].value == "a"
    assert ws["A6"].value == "b"
    assert ws["A12"].value == "below"


def test_dimension_before_sheet_views(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "x"
    ws.update_dimension()
    out = tmp_path / "dim.xlsx"
    wb.save(out)

    from lxml import etree
    import zipfile

    root = etree.fromstring(zipfile.ZipFile(out).read("xl/worksheets/sheet1.xml"))
    names = [etree.QName(c).localname for c in root]
    assert names.index("dimension") < names.index("sheetViews")


def test_write_dataframe_pandas_all_columns(tmp_path: Path):
    pd = pytest.importorskip("pandas")
    wb = Workbook.create()
    df = pd.DataFrame(
        {
            "Item": ["Widget"],
            "Qty": [2],
            "Unit": [10.0],
            "Line total": [20.0],
        }
    )
    wb.write_dataframe(df, at_cell="A1", header=False, mode="overwrite")
    ws = wb["Sheet1"]
    assert ws["A1"].value == "Widget"
    assert ws["B1"].value == 2
    assert ws["C1"].value == 10.0
    assert ws["D1"].value == 20.0


def test_style_source_conflict():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    with pytest.raises(ValueError, match="not both"):
        ws.write_rows(
            [("x",)],
            at_row=2,
            template_rows=1,
            row_styles=[{"bold": True}],
        )


def test_write_many_rows_sequential(tmp_path: Path):
    """Regression: sequential bulk write must stay O(n), not O(n²) per row."""
    n = 3000
    rows = [(i, i * 2, i * 3) for i in range(n)]
    wb = Workbook.create()
    ws = wb["Sheet1"]
    written = ws.write_rows(rows, at_row=2)
    assert written == n
    assert ws["A2"].value == 0
    assert ws["C2"].value == 0
    assert ws["A2001"].value == 1999
    assert ws["C2001"].value == 5997
    last_row = 2 + n - 1
    assert ws[f"A{last_row}"].value == n - 1
    assert ws[f"C{last_row}"].value == (n - 1) * 3

    out = tmp_path / "many_rows.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    assert wb2["Sheet1"]["A2"].value == 0
    assert wb2["Sheet1"][f"C{last_row}"].value == (n - 1) * 3

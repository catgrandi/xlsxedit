"""Tests for the optional pandas ExcelWriter / reader engine."""

from __future__ import annotations

from pathlib import Path

import pytest

pytest.importorskip("pandas")

import pandas as pd
from pandas.io.excel._base import ExcelFile

from xlsxedit import Workbook
from xlsxedit.pandas_io import read_excel, register


@pytest.fixture(scope="module", autouse=True)
def _register_engine():
    register()


def test_to_excel_create_mode(tmp_path: Path):
    out = tmp_path / "created.xlsx"
    df = pd.DataFrame({"Item": ["Widget", "Gadget"], "Qty": [2, 5]})
    with pd.ExcelWriter(out, engine="xlsxedit") as writer:
        df.to_excel(writer, sheet_name="Data", index=False)

    wb = Workbook.open(out)
    assert wb.sheetnames == ["Data"]
    assert wb["Data"]["A1"].value == "Item"
    assert wb["Data"]["B2"].value == 2
    assert wb["Data"]["A3"].value == "Gadget"


def test_read_excel_round_trip(tmp_path: Path):
    path = tmp_path / "round.xlsx"
    src = pd.DataFrame({"Item": ["Widget", "Gadget"], "Qty": [2, 5]})
    with pd.ExcelWriter(path, engine="xlsxedit") as writer:
        src.to_excel(writer, sheet_name="Data", index=False)

    got = pd.read_excel(path, engine="xlsxedit")
    pd.testing.assert_frame_equal(got, src)


def test_read_excel_multi_sheet(tmp_path: Path):
    path = tmp_path / "multi.xlsx"
    wb = Workbook.create()
    wb.rename_worksheet("Sheet1", "One")
    wb["One"]["A1"].value = "a"
    wb.add_worksheet("Two")
    wb["Two"]["A1"].value = "b"
    wb.save(path)

    one = pd.read_excel(path, engine="xlsxedit", sheet_name="One", header=None)
    two = pd.read_excel(path, engine="xlsxedit", sheet_name="Two", header=None)
    assert one.iloc[0, 0] == "a"
    assert two.iloc[0, 0] == "b"


def test_read_excel_sparse_sheet(tmp_path: Path):
    path = tmp_path / "sparse.xlsx"
    wb = Workbook.create()
    wb["Sheet1"]["C3"].value = "c3"
    wb["Sheet1"]["E5"].value = "e5"
    wb.save(path)

    df = pd.read_excel(path, engine="xlsxedit", header=None)
    assert df.shape == (3, 5)
    assert df.iloc[0, 2] == "c3"
    assert df.iloc[2, 4] == "e5"


def test_read_excel_formula_cached_value(tmp_path: Path):
    path = tmp_path / "formula.xlsx"
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = 10
    ws["B1"].value = 11
    ws["B1"].formula = "=A1+1"
    wb.save(path)

    df = pd.read_excel(path, engine="xlsxedit", header=None)
    assert df.iloc[0, 0] == 10
    assert df.iloc[0, 1] == 11
    assert df.iloc[0, 1] != "=A1+1"


def test_read_excel_formula_error_ref(tmp_path: Path):
    from xlsxedit.cell import _ensure_v

    path = tmp_path / "ref_error.xlsx"
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "ok"
    cell = ws["B1"]
    cell.formula = "=Z99"
    cell._element.set("t", "e")
    _ensure_v(cell._element).text = "#REF!"
    wb.save(path)

    df = pd.read_excel(path, engine="xlsxedit", header=None)
    assert df.iloc[0, 0] == "ok"
    assert pd.isna(df.iloc[0, 1])


def test_read_excel_merge_value_only_at_anchor(tmp_path: Path):
    path = tmp_path / "merge.xlsx"
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "v"
    ws.merge_cells("A1:A3")
    ws["B1"].value = 1
    ws["B2"].value = 2
    ws["B3"].value = 3
    wb.save(path)

    df = pd.read_excel(path, engine="xlsxedit", header=None)
    assert df.iloc[0, 0] == "v"
    assert pd.isna(df.iloc[1, 0])
    assert pd.isna(df.iloc[2, 0])
    assert list(df.iloc[:, 1]) == [1, 2, 3]


def test_read_excel_nrows(tmp_path: Path):
    path = tmp_path / "nrows.xlsx"
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "h"
    ws["A2"].value = "a"
    ws["A3"].value = "b"
    ws["A4"].value = "c"
    wb.save(path)

    df = pd.read_excel(path, engine="xlsxedit", nrows=1)
    assert len(df) == 1
    assert df.iloc[0, 0] == "a"


def test_register_patches_engines():
    assert "xlsxedit" in ExcelFile._engines
    assert ExcelFile._engines["xlsxedit"].__name__ == "XlsxeditReader"


def test_read_excel_helper(tmp_path: Path):
    path = tmp_path / "helper.xlsx"
    wb = Workbook.create()
    wb["Sheet1"]["A1"].value = "x"
    wb.save(path)
    df = read_excel(path, header=None)
    assert df.iloc[0, 0] == "x"


def test_to_excel_append_overlay_template(tmp_path: Path):
    template = tmp_path / "template.xlsx"
    wb = Workbook.create()
    ws = wb.rename_worksheet("Sheet1", "Report")
    ws["A1"].value = "Title"
    ws["A4"].value = "Item"
    ws["B4"].value = "Qty"
    ws["A10"].value = "footer"
    wb.save(template)

    df = pd.DataFrame({"Item": ["A", "B"], "Qty": [1, 2]})
    with pd.ExcelWriter(
        template,
        engine="xlsxedit",
        mode="a",
        if_sheet_exists="overlay",
    ) as writer:
        df.to_excel(writer, sheet_name="Report", startrow=4, header=False, index=False)

    wb2 = Workbook.open(template)
    assert wb2["Report"]["A1"].value == "Title"
    assert wb2["Report"]["A5"].value == "A"
    assert wb2["Report"]["B6"].value == 2
    assert wb2["Report"]["A10"].value == "footer"


def test_engine_kwargs_template_rows(tmp_path: Path):
    template = tmp_path / "styled.xlsx"
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "h"
    ws["A1"].apply_style(bold=True, bg_color="FFD9E1F2")
    ws["A2"].value = "sample"
    ws["A2"].apply_style(bg_color="FF9DC3E6")
    wb.save(template)

    df = pd.DataFrame({"X": [10, 20]})
    with pd.ExcelWriter(
        template,
        engine="xlsxedit",
        mode="a",
        if_sheet_exists="overlay",
        engine_kwargs={"template_rows": [2]},
    ) as writer:
        df.to_excel(writer, sheet_name="Sheet1", startrow=2, header=False, index=False)

    wb2 = Workbook.open(template)
    assert wb2["Sheet1"]["A3"].value == 10
    assert wb2["Sheet1"]["A3"].style_index == wb2["Sheet1"]["A2"].style_index


def test_if_sheet_exists_error(tmp_path: Path):
    path = tmp_path / "err.xlsx"
    Workbook.create().save(path)
    df = pd.DataFrame({"A": [1]})
    with pytest.raises(ValueError, match="already exists"):
        with pd.ExcelWriter(path, engine="xlsxedit", mode="a", if_sheet_exists="error") as w:
            df.to_excel(w, sheet_name="Sheet1", index=False)


def test_write_rows_at_helper():
    wb = Workbook.create()
    n = wb["Sheet1"]._write_rows_at([("a", 1), ("b", 2)], start_row=3, start_col_idx=1)
    assert n == 2
    assert wb["Sheet1"]["B3"].value == "a"
    assert wb["Sheet1"]["C4"].value == 2


def test_date_datetime_formats_round_trip(tmp_path: Path):
    from datetime import date, datetime

    path = tmp_path / "dates.xlsx"
    df = pd.DataFrame(
        {
            "d": [date(2025, 8, 7)],
            "dt": [datetime(2025, 8, 7, 14, 30)],
        }
    )
    with pd.ExcelWriter(path, engine="xlsxedit") as writer:
        df.to_excel(writer, index=False)

    wb = Workbook.open(path)
    assert wb["Sheet1"]["A2"].style.is_date
    assert wb["Sheet1"]["A2"].style.num_format == "YYYY-MM-DD"
    assert isinstance(wb["Sheet1"]["A2"].value, datetime)
    assert wb["Sheet1"]["B2"].style.num_format == "YYYY-MM-DD HH:MM:SS"
    assert wb["Sheet1"]["B2"].value == datetime(2025, 8, 7, 14, 30)

    got = pd.read_excel(path, engine="xlsxedit")
    assert pd.api.types.is_datetime64_any_dtype(got["d"])
    assert pd.api.types.is_datetime64_any_dtype(got["dt"])
    assert got["dt"].iloc[0] == pd.Timestamp("2025-08-07 14:30:00")


def test_custom_date_formats(tmp_path: Path):
    from datetime import date, datetime

    path = tmp_path / "custom_dates.xlsx"
    df = pd.DataFrame(
        {
            "d": [date(2025, 8, 7)],
            "dt": [datetime(2025, 8, 7, 14, 30)],
        }
    )
    with pd.ExcelWriter(
        path,
        engine="xlsxedit",
        date_format="DD/MM/YYYY",
        datetime_format="YYYY-MM-DD HH:MM",
    ) as writer:
        df.to_excel(writer, index=False)

    wb = Workbook.open(path)
    assert wb["Sheet1"]["A2"].style.num_format == "DD/MM/YYYY"
    assert wb["Sheet1"]["B2"].style.num_format == "YYYY-MM-DD HH:MM"


def test_multiindex_column_merges(tmp_path: Path):
    path = tmp_path / "mi_cols.xlsx"
    cols = pd.MultiIndex.from_product([["A", "B"], ["x", "y"]])
    df = pd.DataFrame([[1, 2, 3, 4]], columns=cols)
    with pd.ExcelWriter(path, engine="xlsxedit") as writer:
        df.to_excel(writer)  # index=True required by pandas for MultiIndex columns

    wb = Workbook.open(path)
    refs = set(wb["Sheet1"].merged_ranges)
    assert "B1:C1" in refs
    assert "D1:E1" in refs
    assert wb["Sheet1"]["B1"].value == "A"
    assert wb["Sheet1"]["D1"].value == "B"
    assert wb["Sheet1"]["B2"].value == "x"
    assert wb["Sheet1"]["C2"].value == "y"


def test_multiindex_index_merges(tmp_path: Path):
    path = tmp_path / "mi_idx.xlsx"
    idx = pd.MultiIndex.from_product([["G1", "G2"], ["a", "b"]])
    df = pd.DataFrame({"v": [1, 2, 3, 4]}, index=idx)
    with pd.ExcelWriter(path, engine="xlsxedit") as writer:
        df.to_excel(writer)

    wb = Workbook.open(path)
    refs = set(wb["Sheet1"].merged_ranges)
    assert "A2:A3" in refs
    assert "A4:A5" in refs
    assert wb["Sheet1"]["A2"].value == "G1"
    assert wb["Sheet1"]["A4"].value == "G2"

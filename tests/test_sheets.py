from __future__ import annotations

from pathlib import Path

import pytest

from xlsxedit import Workbook, WorksheetNotFoundError
from tests.preservation import assert_preserved, read_pkg
from tests.smoke_workbook import SMOKE_IMAGE


def test_add_worksheet():
    wb = Workbook.create()
    ws2 = wb.add_worksheet("Sheet2")
    assert wb.sheetnames == ["Sheet1", "Sheet2"]
    assert ws2.name == "Sheet2"
    assert list(ws2.cells) == []


def test_add_worksheet_duplicate_name_raises():
    wb = Workbook.create()
    with pytest.raises(ValueError, match="already exists"):
        wb.add_worksheet("Sheet1")


def test_add_worksheet_save_roundtrip(tmp_path: Path):
    out = tmp_path / "two_sheets.xlsx"
    wb = Workbook.create()
    ws2 = wb.add_worksheet("Report")
    ws2["A1"].value = "Title"
    wb.save(out)

    wb2 = Workbook.open(out)
    assert wb2.sheetnames == ["Sheet1", "Report"]
    assert wb2["Report"]["A1"].value == "Title"
    assert list(wb2["Sheet1"].cells) == []


# --- one Worksheet per sheet (#44) -------------------------------------------

_SHEET_LIST_CHANGES = {
    "add": lambda wb: wb.add_worksheet("Added"),
    "copy": lambda wb: wb.copy_worksheet("Sheet1", "Copy"),
    "rename": lambda wb: wb.rename_worksheet("Sheet1", "Renamed"),
    "remove other": lambda wb: wb.remove_worksheet("Other"),
}


@pytest.mark.parametrize("change", _SHEET_LIST_CHANGES.values(), ids=_SHEET_LIST_CHANGES)
def test_a_held_worksheet_sees_writes_through_wb_name_and_the_reverse(change):
    wb = Workbook.create()
    wb.add_worksheet("Other")
    held = wb["Sheet1"]
    held.write_rows([["head", "line"]], at_row=1)
    held["A1"].apply_style(bold=True)
    change(wb)
    current = wb[held.name]
    assert current is held

    current["B5"].value = "through wb"
    current.row_dimensions[5].height = 30
    held["C6"].value = "through held"
    held.row_dimensions[6].height = 40
    assert held["B5"].value == "through wb"
    assert held.row_dimensions[5].height == 30
    assert current["C6"].value == "through held"
    assert current.row_dimensions[6].height == 40
    held.write_rows([["x", "y"]], at_row=7, template_rows=1)
    assert current["A7"].style.bold
    assert held.clear_range("A1:C7") == 6
    assert [current[a].value for a in ("A1", "B5", "C6", "B7")] == [None] * 4


def test_rename_renames_the_held_worksheet():
    wb = Workbook.create()
    held = wb["Sheet1"]
    assert wb.rename_worksheet("Sheet1", "Data") is held
    assert held.name == "Data"
    assert wb.sheetnames == ["Data"]
    with pytest.raises(WorksheetNotFoundError):
        wb["Sheet1"]


def test_a_removed_sheets_worksheet_raises_on_use():
    wb = Workbook.create()
    gone = wb.add_worksheet("Gone")
    gone.write_rows([[1, 2], [3, 4]], at_row=1)
    gone.merge_cells("A3:B3")
    cell, missing = gone["A1"], gone["C9"]
    assert gone.row_dimensions[1].height is None  # builds the row index
    assert gone.merged_ranges == ["A3:B3"]
    wb.remove_worksheet("Gone")
    replacement = wb.add_worksheet("Gone")
    before = read_pkg(wb)

    for use in (
        lambda: gone["A1"],
        lambda: gone.cells,
        lambda: gone.values(),
        lambda: gone.merged_ranges,
        lambda: gone.row_dimensions[1].height,
        lambda: gone.column_dimensions["A"].width,
        lambda: gone.tables,
        lambda: gone.images,
        lambda: gone.write_rows([[5]], at_row=9),
        lambda: gone.clear_range("A1:B2"),
        lambda: setattr(cell, "value", 5),
        lambda: cell.apply_style(bold=True),
        lambda: missing.value,
        lambda: setattr(missing, "value", 5),
        lambda: gone.add_image(SMOKE_IMAGE, anchor="B2"),
        lambda: gone.add_chart(anchor="E2", data_range="A1:B2"),
    ):
        with pytest.raises(WorksheetNotFoundError, match="'Gone' was removed"):
            use()
    assert gone.name == "Gone"
    assert (repr(missing), missing.address) == ("<Cell 'C9'>", "C9")
    assert wb["Gone"] is replacement
    assert_preserved(before, read_pkg(wb))

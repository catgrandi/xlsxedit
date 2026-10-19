#!/usr/bin/env python3
"""Author the Excel-reference pairs in ``fixtures/excel-reference/`` through Excel.

Windows only; needs Microsoft Excel and pywin32. Run from the repository root:

    python scripts/excel_fixtures/build_reference.py                    # every pair
    python scripts/excel_fixtures/build_reference.py table-append-row   # just one

For each operation the script builds ``before.xlsx`` in Excel, opens it again,
applies the operation through Excel's object model, saves ``after.xlsx``, and
writes the folder's ``README.md`` with the UI steps, the call that stood in
for them, what Excel did, and the Excel build. ``tests/test_excel_reference.py``
applies the matching xlsxedit operation to ``before.xlsx`` and compares it
with ``after.xlsx``.
"""

from __future__ import annotations

import argparse
import datetime as dt
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from _excel import FIXTURES, Excel, xl
from build_fixtures import _at, _rows

REFERENCE = FIXTURES / "excel-reference"


@dataclass(frozen=True)
class Operation:
    name: str
    summary: str
    build: Callable[[Excel], object]
    apply: Callable[[object], None]
    steps: tuple[str, ...]
    call: str
    observed: tuple[str, ...]


def _shape(ws, name, top_left, bottom_right, placement):
    (left, top), (right, bottom) = _at(ws, top_left), _at(ws, bottom_right)
    shp = ws.Shapes.AddShape(xl.ShapeRectangle, left + 3, top + 2, right - left - 6, bottom - top - 4)
    shp.Name = name
    shp.Placement = placement
    return shp


# table-append-row ---------------------------------------------------------


def _table_append_row_before(excel: Excel):
    wb = excel.new_workbook("Orders")
    ws = wb.Worksheets("Orders")
    _rows(
        ws,
        "A1",
        [
            ("Item", "Qty", "Price", "Amount", "Shipped"),
            ("Widget", 4, 2.5, None, dt.date(2026, 3, 2)),
            ("Gadget", 10, 1.25, None, dt.date(2026, 3, 5)),
            ("Bolt", 100, 0.1, None, dt.date(2026, 3, 9)),
            ("Bracket", 12, 3.75, None, dt.date(2026, 3, 12)),
            ("Hinge", 8, 4.2, None, dt.date(2026, 3, 16)),
        ],
    )
    lo = ws.ListObjects.Add(xl.SrcRange, ws.Range("A1:E6"), None, xl.Yes)
    lo.Name = "Orders"
    lo.ListColumns("Amount").DataBodyRange.Formula = "=[@Qty]*[@Price]"
    lo.ListColumns("Price").DataBodyRange.NumberFormat = "#,##0.00"
    lo.ListColumns("Amount").DataBodyRange.NumberFormat = "#,##0.00"
    lo.ListColumns("Shipped").DataBodyRange.NumberFormat = "yyyy-mm-dd"
    rule = lo.ListColumns("Amount").DataBodyRange.FormatConditions.Add(
        xl.CellValue, xl.Greater, "=40"
    )
    rule.Interior.Color = 0xCEEFC6
    lo.ListColumns("Qty").DataBodyRange.Validation.Add(
        xl.ValidateWholeNumber, xl.ValidAlertStop, xl.GreaterEqual, "0"
    )
    ws.Range("A8").Value = "Note two rows below the table."
    ws.Range("G2").Value = "Beside the table."
    return wb


def _table_append_row_apply(wb) -> None:
    wb.Worksheets("Orders").ListObjects("Orders").ListRows.Add()


# insert-rows-above-table --------------------------------------------------


def _insert_rows_above_table_before(excel: Excel):
    wb = excel.new_workbook("Stock", "Lists")
    ws, lists = wb.Worksheets("Stock"), wb.Worksheets("Lists")
    _rows(lists, "A1", [("Yes",), ("No",), ("Later",)])
    ws.Range("A1:D1").Merge()
    ws.Range("A1").Value = "Inventory"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Counted on the first Monday of each month."
    _rows(
        ws,
        "A4",
        [
            ("Item", "Qty", "Price"),
            ("Bolt", 120, 0.1),
            ("Nut", 300, 0.05),
            ("Washer", 80, 0.02),
            ("Hinge", 15, 4.2),
            ("Clamp", 9, 7.5),
        ],
    )
    lo = ws.ListObjects.Add(xl.SrcRange, ws.Range("A4:C9"), None, xl.Yes)
    lo.Name = "Stock"
    lo.ListColumns("Qty").DataBodyRange.FormatConditions.Add(
        xl.CellValue, xl.Greater, "=100"
    ).Font.Bold = True
    ws.Range("E4").Value = "Qty total"
    ws.Range("E5").Formula = "=SUM(B5:B9)"
    ws.Range("F2").Value = "Reorder?"
    _rows(ws, "F3", [(None,)] * 8)
    ws.Range("F3:F10").FormatConditions.Add(xl.Expression, None, '=F3="Yes"').Interior.Color = (
        0x9CEBFF
    )
    ws.Range("F3:F10").Validation.Add(
        xl.ValidateList, xl.ValidAlertStop, xl.Between, "=Lists!$A$1:$A$3"
    )
    ws.Range("G3:G10").Validation.Add(xl.ValidateList, xl.ValidAlertStop, xl.Between, "Low,High")
    ws.Range("A12:B12").Merge()
    ws.Range("A12").Value = "Checked by"
    wb.Names.Add("Reorder", "=Stock!$F$3:$F$10")
    wb.Names.Add("Title", "=Stock!$A$1")
    return wb


def _insert_rows_above_table_apply(wb) -> None:
    wb.Worksheets("Stock").Rows("3:4").Insert()


# insert-rows-through-drawings ---------------------------------------------


def _insert_rows_through_drawings_before(excel: Excel):
    wb = excel.new_workbook("Drawings")
    ws = wb.Worksheets("Drawings")
    for row in range(1, 31):
        ws.Cells(row, 1).Value = row
    for row, height in {6: 30, 8: 36, 15: 24}.items():
        ws.Rows(row).RowHeight = height
    placements = (
        ("two-cell", xl.MoveAndSize, "C", "D"),
        ("one-cell", xl.Move, "E", "F"),
        ("absolute", xl.FreeFloating, "G", "H"),
    )
    _shape(ws, "Above two-cell", "C1", "D3", xl.MoveAndSize)
    for label, placement, first, last in placements:
        _shape(ws, f"Straddle {label}", f"{first}5", f"{last}11", placement)
        _shape(ws, f"Below {label}", f"{first}14", f"{last}18", placement)
    left, top = _at(ws, "J5")
    right, bottom = _at(ws, "N11")
    chart = ws.ChartObjects().Add(left, top, right - left, bottom - top)
    chart.Name = "Straddle chart"
    chart.Chart.ChartType = xl.ColumnClustered
    chart.Chart.SetSourceData(ws.Range("A1:A5"))
    return wb


def _insert_rows_through_drawings_apply(wb) -> None:
    wb.Worksheets("Drawings").Rows("7:9").Insert()


# delete-rows-under-cf -----------------------------------------------------


def _delete_rows_under_cf_before(excel: Excel):
    wb = excel.new_workbook("Sheet1", "Lists")
    ws, lists = wb.Worksheets("Sheet1"), wb.Worksheets("Lists")
    _rows(lists, "A1", [("Yes",), ("No",)])
    for row in range(1, 21):
        for col in range(1, 9):
            ws.Cells(row, col).Value = row * 10 + col

    def rule(address, threshold):
        ws.Range(address).FormatConditions.Add(xl.CellValue, xl.Greater, f"={threshold}").Font.Bold = True

    rule("A2:A20", 50)  # contains the deleted rows
    rule("B5:B12", 60)  # starts at the first deleted row
    rule("C3:C6", 40)  # ends inside the deleted rows
    rule("D5:D7", 0)  # entirely inside the deleted rows
    rule("E10:E15", 100)  # below the deleted rows
    rule("F2:F4,F6:F9", 30)  # two areas, one cut by the delete
    ws.Range("G2:G10").FormatConditions.AddDatabar()
    ws.Range("H5:H7").Validation.Add(xl.ValidateWholeNumber, xl.ValidAlertStop, xl.GreaterEqual, "0")
    ws.Range("H9:H12").Validation.Add(xl.ValidateWholeNumber, xl.ValidAlertStop, xl.GreaterEqual, "0")
    ws.Range("I2:I10").Validation.Add(
        xl.ValidateList, xl.ValidAlertStop, xl.Between, "=Lists!$A$1:$A$2"
    )
    ws.Range("J5:K6").Merge()
    ws.Range("J9:K10").Merge()
    wb.Names.Add("Watch", "=Sheet1!$A$6")
    wb.Names.Add("Block", "=Sheet1!$A$3:$B$9")
    for label, placement, first, last in (
        ("two-cell", xl.MoveAndSize, "M", "N"),
        ("one-cell", xl.Move, "O", "P"),
        ("absolute", xl.FreeFloating, "Q", "R"),
    ):
        _shape(ws, f"Inside {label}", f"{first}5", f"{last}8", placement)
        _shape(ws, f"Below {label}", f"{first}10", f"{last}13", placement)
    _shape(ws, "Across two-cell", "S4", "T10", xl.MoveAndSize)
    return wb


def _delete_rows_under_cf_apply(wb) -> None:
    wb.Worksheets("Sheet1").Rows("5:7").Delete()


OPERATIONS = {
    op.name: op
    for op in (
        Operation(
            name="table-append-row",
            summary="A row added to the bottom of a table by pressing Tab in its last cell.",
            build=_table_append_row_before,
            apply=_table_append_row_apply,
            steps=(
                "Select E6, the last cell of the `Orders` table.",
                "Press Tab.",
            ),
            call='Worksheets("Orders").ListObjects("Orders").ListRows.Add()',
            observed=(
                "`ListRows.Add()` stood in for Tab. With `AlwaysInsert` omitted it shifted the "
                "cells below the table down one row within the table's columns, although row 7 "
                "was empty: the note in A8 moved to A9, while G2 beside the table stayed. Whether "
                "Tab does the same when the row below is empty was not checked in the UI.",
                "The table and its autoFilter grow from A1:E6 to A1:E7. The new row is empty "
                "except D7, which gets the calculated-column formula.",
                "C7, D7 and E7 take their columns' number formats (the `dataDxfId` formats).",
                "The conditional format on the Amount column grows from D2:D6 to D2:D7, and the "
                "validation on the Qty column from B2:B6 to B2:B7.",
            ),
        ),
        Operation(
            name="insert-rows-above-table",
            summary="Two whole rows inserted above a table, at the first row of a conditional "
            "format and two validations.",
            build=_insert_rows_above_table_before,
            apply=_insert_rows_above_table_apply,
            steps=(
                "Select rows 3:4 on `Stock` by their row headers.",
                "Home > Insert > Insert Sheet Rows.",
            ),
            call='Worksheets("Stock").Rows("3:4").Insert()',
            observed=(
                "Everything from row 3 down moves two rows: the table (A4:C9 to A6:C11) and its "
                "autoFilter, the merge A12:B12, the defined name `Reorder`, and the formula in E5 "
                "(now E7, `SUM(B5:B9)` rewritten to `SUM(B7:B11)`).",
                "A conditional format or validation whose first row is the insert row moves; it "
                "does not grow. F3:F10 becomes F5:F12 for the conditional format and the x14 list "
                "validation, and G3:G10 becomes G5:G12 for the base list validation.",
                "The expression rule's relative formula follows its range: `F3=\"Yes\"` becomes "
                "`F5=\"Yes\"`.",
                "Rows 1 and 2, the merged title and the note, stay where they are, and so does "
                "the defined name `Title`.",
                "The inserted rows take the format of row 2 above them, which is the default.",
            ),
        ),
        Operation(
            name="insert-rows-through-drawings",
            summary="Three whole rows inserted through and above drawings of every placement.",
            build=_insert_rows_through_drawings_before,
            apply=_insert_rows_through_drawings_apply,
            steps=(
                "Select rows 7:9 on `Drawings` by their row headers.",
                "Home > Insert > Insert Sheet Rows.",
            ),
            call='Worksheets("Drawings").Rows("7:9").Insert()',
            observed=(
                "The inserted rows 7 to 9 copy the 30 pt height of row 6; the custom heights of "
                "rows 8 and 15 move to rows 11 and 18.",
                "A two-cell anchor without `editAs` that straddles the insert grows: `from` stays "
                "and `to` moves down three rows. The chart does the same.",
                '`editAs="oneCell"` and `editAs="absolute"` anchors that straddle the insert keep '
                "their size: `from` stays and Excel recomputes `to`, cell and offset, over the new "
                "row heights.",
                "Two-cell and one-cell anchors below the insert move down three rows and keep "
                "their offsets.",
                '`editAs="absolute"` below the insert stays where it was on the page: Excel '
                "rewrites `from` and `to`, cells and offsets, to describe the same position over "
                "the new rows (`from` row 13 + 25400 EMU becomes row 10 + 254000 EMU).",
                "The anchor above the insert does not change.",
            ),
        ),
        Operation(
            name="delete-rows-under-cf",
            summary="Three whole rows deleted under conditional formats, validations, merges, "
            "defined names and drawings.",
            build=_delete_rows_under_cf_before,
            apply=_delete_rows_under_cf_apply,
            steps=(
                "Select rows 5:7 on `Sheet1` by their row headers.",
                "Home > Delete > Delete Sheet Rows.",
            ),
            call='Worksheets("Sheet1").Rows("5:7").Delete()',
            observed=(
                "Conditional formats: a range that contains the deleted rows shrinks (A2:A20 to "
                "A2:A17); one that starts at the first deleted row keeps its start (B5:B12 to "
                "B5:B9); one that ends inside them is cut (C3:C6 to C3:C4); one entirely inside "
                "them is removed (D5:D7); one below moves up (E10:E15 to E7:E12); the x14 data "
                "bar shrinks (G2:G10 to G2:G7).",
                "A multi-area range whose areas become adjacent is merged: `F2:F4 F6:F9` becomes "
                "`F2:F6`.",
                "Validations follow the same rules, base and x14: H5:H7 is removed, H9:H12 "
                "becomes H6:H9, and the x14 list I2:I10 becomes I2:I7. Excel had already stored "
                "H5:H7 and H9:H12, added as two identical validations, as one with sqref "
                "`H5:H7 H9:H12`.",
                "Merges: J5:K6, inside the deleted rows, is removed; J9:K10 becomes J6:K7.",
                "Defined names: `Block` shrinks from A3:B9 to A3:B6; `Watch` named A6, a deleted "
                "cell, and becomes `Sheet1!#REF!`.",
                "Drawings inside the deleted rows: the two-cell (move and size) shape is deleted; "
                'the `editAs="oneCell"` shape survives at the top of row 5 with its size kept; '
                'the `editAs="absolute"` shape survives where it was.',
                "Drawings below the deleted rows: two-cell and one-cell shapes move up three rows; "
                "the absolute one stays where it was on the page, which here leaves its anchor "
                "cells unchanged because the rows are all the same height. The two-cell shape "
                "across the deleted rows shrinks by three rows.",
            ),
        ),
    )
}


def _readme(op: Operation, excel: Excel) -> str:
    lines = [
        f"# {op.name}",
        "",
        op.summary,
        "",
        "## Steps in Excel",
        "",
        "1. Open `before.xlsx`.",
        *(f"{n}. {step}" for n, step in enumerate(op.steps, start=2)),
        f"{len(op.steps) + 2}. Save as `after.xlsx`.",
        "",
        "## How this pair was made",
        "",
        f"`python scripts/excel_fixtures/build_reference.py {op.name}` built `before.xlsx` "
        "in Excel, opened it again and ran the steps through Excel's object model:",
        "",
        "```vb",
        op.call,
        "```",
        "",
        f"- Excel: {excel.version}",
        f"- Platform: {excel.platform}",
        f"- Date: {dt.date.today().isoformat()}",
    ]
    if op.observed:
        lines += ["", "## What Excel did", "", *(f"- {item}" for item in op.observed)]
    return "\n".join(lines) + "\n"


def build(op: Operation, excel: Excel) -> Path:
    folder = REFERENCE / op.name
    before = excel.save(op.build(excel), folder / "before.xlsx")
    wb = excel.open(before)
    op.apply(wb)
    excel.save(wb, folder / "after.xlsx")
    (folder / "README.md").write_text(_readme(op, excel), encoding="utf-8", newline="\n")
    return folder


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument(
        "names", nargs="*", metavar="NAME", help=f"pairs to build (default all): {', '.join(OPERATIONS)}"
    )
    args = parser.parse_args(argv)
    unknown = sorted(set(args.names) - OPERATIONS.keys())
    if unknown:
        parser.error(f"unknown operation: {', '.join(unknown)}")
    with Excel() as excel:
        print(f"Excel {excel.version}; {excel.platform}")
        for name in args.names or list(OPERATIONS):
            folder = build(OPERATIONS[name], excel)
            print(f"wrote {folder.relative_to(FIXTURES.parent)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

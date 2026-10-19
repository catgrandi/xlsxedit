#!/usr/bin/env python3
"""Author the Excel-written fixtures in ``fixtures/`` by driving Excel through COM.

Windows only; needs Microsoft Excel and pywin32. Run from the repository root:

    python scripts/excel_fixtures/build_fixtures.py            # every fixture
    python scripts/excel_fixtures/build_fixtures.py TableRich  # just one

Each builder makes a new workbook in a private Excel instance and saves it
through Excel, so every part is Excel's own XML. Features Excel's object model
cannot create are listed in ``scripts/excel_fixtures/README.md`` with the
steps to add them by hand. Regenerate ``fixtures/MANIFEST.md`` afterwards with
``python scripts/describe_fixtures.py``.
"""

from __future__ import annotations

import argparse
import datetime as dt
import struct
import sys
import tempfile
import zlib
from pathlib import Path

import pythoncom
from _excel import FIXTURES, Excel, personal_data, personal_strings, xl

ORDERS = [
    ("Widget", 4, 2.5, dt.date(2026, 3, 2)),
    ("Gadget", 10, 1.25, dt.date(2026, 3, 5)),
    ("Bolt", 100, 0.1, dt.date(2026, 3, 9)),
    ("Nut", 250, 0.05, dt.date(2026, 3, 9)),
    ("Bracket", 12, 3.75, dt.date(2026, 3, 12)),
    ("Hinge", 8, 4.2, dt.date(2026, 3, 16)),
    ("Spring", 30, 0.6, dt.date(2026, 3, 20)),
    ("Clamp", 6, 7.5, dt.date(2026, 3, 23)),
]


def _serial(value):
    """Dates as Excel serial numbers; the caller gives the cells a date format."""
    if isinstance(value, dt.date):
        return (value - dt.date(1899, 12, 30)).days
    return value


def _rows(ws, top_left: str, rows) -> None:
    """Write a rectangular block of values starting at ``top_left``."""
    rows = tuple(tuple(_serial(v) for v in r) for r in rows)
    first = ws.Range(top_left)
    last = ws.Cells(first.Row + len(rows) - 1, first.Column + len(rows[0]) - 1)
    ws.Range(first, last).Value = rows


def _png(path: Path, size: int = 16) -> Path:
    """A small two-colour PNG, so the drawing fixtures need no image asset."""
    raw = b"".join(
        b"\x00"
        + b"".join(
            (b"\x1f\x6f\xb4" if (x // 4 + y // 4) % 2 else b"\xf2\xa9\x00") for x in range(size)
        )
        for y in range(size)
    )

    def chunk(kind: bytes, data: bytes) -> bytes:
        body = kind + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body))

    header = struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0)
    path.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", header)
        + chunk(b"IDAT", zlib.compress(raw, 9))
        + chunk(b"IEND", b"")
    )
    return path


def _at(ws, address: str) -> tuple[float, float]:
    """Left and top of a cell, in points."""
    cell = ws.Range(address)
    return cell.Left, cell.Top


def table_rich(excel: Excel) -> Path:
    """A table with every table feature, and a second table directly below it."""
    wb = excel.new_workbook("Orders")
    ws = wb.Worksheets("Orders")
    _rows(ws, "A1", [("Item", "Qty", "Price", "Amount", "Shipped")])
    _rows(ws, "A2", [(item, qty, price, None, shipped) for item, qty, price, shipped in ORDERS])
    lo = ws.ListObjects.Add(xl.SrcRange, ws.Range(f"A1:E{1 + len(ORDERS)}"), None, xl.Yes)
    lo.Name = "Orders"
    lo.TableStyle = "TableStyleMedium2"
    lo.ShowTableStyleRowStripes = True
    lo.ListColumns("Amount").DataBodyRange.Formula = "=[@Qty]*[@Price]"
    lo.ListColumns("Price").DataBodyRange.NumberFormat = "#,##0.00"
    lo.ListColumns("Amount").DataBodyRange.NumberFormat = "#,##0.00"
    lo.ListColumns("Shipped").DataBodyRange.NumberFormat = "yyyy-mm-dd"
    lo.Sort.SortFields.Clear()
    lo.Sort.SortFields.Add(lo.ListColumns("Qty").DataBodyRange, xl.SortOnValues, xl.Descending)
    lo.Sort.Header = xl.Yes
    lo.Sort.Apply()
    # One cell whose formula differs from the column's. A single-cell write would
    # refill the whole column, so write it together with its consistent neighbour.
    ws.Range("D4:D5").Formula = (("=[@Qty]*[@Price]*0.9",), ("=[@Qty]*[@Price]",))
    lo.ShowTotals = True
    lo.ListColumns("Item").TotalsCalculation = xl.TotalsNone
    lo.TotalsRowRange.Cells(1, 1).Value = "Total"
    lo.ListColumns("Qty").TotalsCalculation = xl.TotalsSum
    lo.ListColumns("Amount").TotalsCalculation = xl.TotalsSum
    lo.ListColumns("Shipped").TotalsCalculation = xl.TotalsNone
    below = lo.TotalsRowRange.Row + 1
    _rows(
        ws,
        f"A{below}",
        [("Code", "Owner", "Note"), ("A-1", "Ops", "restock"), ("B-7", "Sales", "promo")],
    )
    notes = ws.ListObjects.Add(xl.SrcRange, ws.Range(f"A{below}:C{below + 2}"), None, xl.Yes)
    notes.Name = "Notes"
    notes.TableStyle = "TableStyleLight9"
    return excel.save(wb, FIXTURES / "TableRich.xlsx")


def table_with_content_below(excel: Excel) -> Path:
    """A table with a note, a merged title and a chart below it, and content to its right."""
    wb = excel.new_workbook("Report")
    ws = wb.Worksheets("Report")
    _rows(
        ws,
        "A1",
        [
            ("Region", "Q1", "Q2", "Total"),
            ("North", 120, 135, None),
            ("South", 98, 104, None),
            ("East", 143, 150, None),
            ("West", 87, 95, None),
            ("Central", 110, 118, None),
        ],
    )
    # Written before the table exists: text typed directly below a table extends it.
    ws.Range("A7").Value = "Figures in thousands of units."
    ws.Range("A7").Font.Italic = True
    ws.Range("A8:D8").Merge()
    ws.Range("A8").Value = "Quarterly sales by region"
    ws.Range("A8").Font.Bold = True
    ws.Range("A8").HorizontalAlignment = -4108  # xlCenter
    lo = ws.ListObjects.Add(xl.SrcRange, ws.Range("A1:D6"), None, xl.Yes)
    lo.Name = "Sales"
    lo.ListColumns("Total").DataBodyRange.Formula = "=[@Q1]+[@Q2]"
    _rows(ws, "F1", [("Target", "Gap")] + [(t, None) for t in (250, 200, 290, 190, 225)])
    ws.Range("G2:G6").Formula = "=D2-F2"
    left, top = _at(ws, "A10")
    right, bottom = _at(ws, "F24")
    chart = ws.ChartObjects().Add(left, top, right - left, bottom - top)
    chart.Name = "Sales chart"
    chart.Chart.ChartType = xl.ColumnClustered
    chart.Chart.SetSourceData(ws.Range("A1:C6"))
    return excel.save(wb, FIXTURES / "TableWithContentBelow.xlsx")


def data_validation(excel: Excel) -> Path:
    """Every kind of validation the roadmap touches, including Excel's x14 form."""
    wb = excel.new_workbook("Form", "Lists")
    ws, lists = wb.Worksheets("Form"), wb.Worksheets("Lists")
    _rows(lists, "A1", [("Red",), ("Green",), ("Blue",), ("Yellow",)])
    _rows(ws, "A1", [("Size", "Colour", "Due", "Weight", "Code", "Count", "Approved")])
    _rows(ws, "J1", [("Sizes",), ("S",), ("M",), ("L",), ("XL",)])

    def add(address, kind, operator, formula1, formula2=None, **messages):
        v = ws.Range(address).Validation
        v.Delete()
        if formula2 is None:
            v.Add(kind, xl.ValidAlertStop, operator, formula1)
        else:
            v.Add(kind, xl.ValidAlertStop, operator, formula1, formula2)
        for name, text in messages.items():
            setattr(v, name, text)
        return v

    add("A2:A20", xl.ValidateList, xl.Between, "=$J$2:$J$5")
    add("B2:B20", xl.ValidateList, xl.Between, "=Lists!$A$1:$A$4")
    add(
        "C2:C20",
        xl.ValidateDate,
        xl.Between,
        "=DATE(2026,1,1)",
        "=DATE(2026,12,31)",
        InputTitle="Due date",
        InputMessage="A date in 2026.",
    )
    add("D2:D20", xl.ValidateDecimal, xl.Between, "0", "99.5")
    add(
        "E2:E20",
        xl.ValidateCustom,
        xl.Between,
        "=LEN(E2)=5",
        ErrorTitle="Code",
        ErrorMessage="Codes have five characters.",
    )
    add("F2:F5,F8:F10", xl.ValidateWholeNumber, xl.GreaterEqual, "0")
    add("G2:G20", xl.ValidateList, xl.Between, "Yes,No")
    return excel.save(wb, FIXTURES / "DataValidation.xlsx")


def conditional_formatting_rich(excel: Excel) -> Path:
    """Multi-area, expression, x14 icon-set and data-bar rules, and CF on a table column."""
    wb = excel.new_workbook("Scores")
    ws = wb.Worksheets("Scores")
    scores = [
        ("Ann", 12, 55, 71, -4),
        ("Ben", -8, 61, 40, 88),
        ("Cy", 45, 23, 90, 67),
        ("Dee", 0, 77, 15, 52),
        ("Eve", 93, -15, 62, 30),
        ("Fay", 27, 48, 84, -11),
        ("Gus", -20, 39, 57, 74),
        ("Hal", 66, 81, 9, 46),
        ("Ida", 34, 5, 49, 95),
        ("Jo", 58, 70, 28, 13),
    ]
    _rows(ws, "A1", [("Name", "Q1", "Q2", "Q3", "Q4")] + scores)
    expression = ws.Range("A2:E11").FormatConditions.Add(xl.Expression, None, "=$B2>0")
    expression.Interior.Color = 0xCEEFC6
    multi = ws.Range("C2:C11,E2:E11").FormatConditions.Add(xl.CellValue, xl.Greater, "=50")
    multi.Font.Bold = True
    icons = ws.Range("D2:D11").FormatConditions.AddIconSetCondition()
    icons.IconSet = wb.IconSets(xl.IconSet3TrafficLights1)
    icons.IconCriteria(1).Icon = xl.IconRedDownArrow
    icons.IconCriteria(2).Icon = xl.IconYellowFlag
    icons.IconCriteria(3).Icon = xl.IconGreenCheck
    ws.Range("B2:B11").FormatConditions.AddDatabar()
    _rows(
        ws,
        "G1",
        [("Team", "Target", "Actual")]
        + [(t, a, b) for t, a, b in zip("PQRSTUV", (50, 60, 70, 80, 55, 65, 75), (48, 66, 70, 91, 40, 64, 80))],
    )
    lo = ws.ListObjects.Add(xl.SrcRange, ws.Range("G1:I8"), None, xl.Yes)
    lo.Name = "Targets"
    lo.ListColumns("Actual").DataBodyRange.FormatConditions.AddColorScale(3)
    return excel.save(wb, FIXTURES / "ConditionalFormattingRich.xlsx")


def drawing_anchors(excel: Excel) -> Path:
    """Every anchor Excel writes for a drawing, over uneven column widths and row heights."""
    wb = excel.new_workbook("Anchors")
    ws = wb.Worksheets("Anchors")
    for column, width in {"B": 20, "C": 4, "D": 32, "E": 11, "F": 2.5, "G": 18, "H": 9}.items():
        ws.Columns(column).ColumnWidth = width
    for row, height in {2: 30, 4: 6, 6: 42, 9: 21, 12: 9, 15: 36}.items():
        ws.Rows(row).RowHeight = height
    for row in range(1, 25):
        ws.Cells(row, 1).Value = row

    def box(name, top_left, bottom_right, placement=None, shape=xl.ShapeRectangle):
        (left, top), (right, bottom) = _at(ws, top_left), _at(ws, bottom_right)
        shp = ws.Shapes.AddShape(shape, left + 3, top + 2, right - left - 6, bottom - top - 4)
        shp.Name = name
        if placement is not None:
            shp.Placement = placement
        return shp

    box("Move and size", "B2", "C5")
    box("Move", "D2", "E5", xl.Move)
    box("Absolute", "E2", "H5", xl.FreeFloating)
    with tempfile.TemporaryDirectory() as tmp:
        left, top = _at(ws, "B7")
        pic = ws.Shapes.AddPicture(str(_png(Path(tmp) / "tile.png")), False, True, left, top, 48, 48)
        pic.Name = "Picture"
    left, top = _at(ws, "D7")
    text = ws.Shapes.AddTextbox(xl.TextHorizontal, left + 4, top + 4, 150, 40)
    text.Name = "Text box"
    text.TextFrame2.TextRange.Text = "A text box"
    oval = box("Group oval", "B10", "C13", shape=xl.ShapeOval)
    rect = box("Group rectangle", "C10", "E13")
    group = ws.Shapes.Range([oval.Name, rect.Name]).Group()
    group.Name = "Group"
    left, top = _at(ws, "E8")
    line = ws.Shapes.AddLine(left, top, left + 120, top + 60)
    line.Name = "Connector"
    left, top = _at(ws, "B16")
    right, bottom = _at(ws, "E24")
    chart = ws.ChartObjects().Add(left, top, right - left, bottom - top)
    chart.Name = "Chart"
    chart.Chart.ChartType = xl.ColumnClustered
    chart.Chart.SetSourceData(ws.Range("A1:A6"))
    return excel.save(wb, FIXTURES / "DrawingAnchors.xlsx")


def formulas(excel: Excel) -> Path:
    """Shared, array, dynamic-array, data-table, cross-sheet and future-function formulas."""
    wb = excel.new_workbook("Calc", "Data", "Other Sheet")
    calc, data, other = (wb.Worksheets(n) for n in ("Calc", "Data", "Other Sheet"))
    _rows(
        data,
        "A1",
        [("Key", "Qty", "Price"), ("a", 3, 2.5), ("b", 7, 1.5), ("c", 2, 9.0), ("d", 5, 4.0), ("e", 1, 12.0)],
    )
    data.Range("E1").Value = "Rate"
    data.Range("E2").Value = 0.2
    wb.Names.Add("Rate", "=Data!$E$2")
    _rows(other, "A1", [(10,), (20,), (30,)])
    _rows(
        calc,
        "A1",
        [("n", "Shared", "Array", "Spill", "Sorted", "Lookup", "Cross-sheet", "", "Input", "", "x", "x*Rate")]
        + [(n,) + (None,) * 11 for n in range(1, 6)],
    )
    # Excel shares a filled-down formula only when it has no reference to another sheet.
    calc.Range("B2").Formula = "=A2*2"
    calc.Range("B2:B6").FillDown()
    calc.Range("C2:C4").FormulaArray = "=Data!B2:B4*2"
    calc.Range("C6").FormulaArray = "=SUM(Data!B2:B6*Data!C2:C6)"
    calc.Range("D2").Formula2 = "=SEQUENCE(4)"
    calc.Range("E2").Formula2 = "=SORT(Data!B2:B6,,-1)"
    calc.Range("F2").Formula2 = '=XLOOKUP("c",Data!A2:A6,Data!C2:C6)'
    calc.Range("G2").Formula = "=SUM('Other Sheet'!A1:A3)*Rate"
    calc.Range("G3").Formula = "=Data!C2+'Other Sheet'!A2"
    calc.Range("I2").Value = 100
    _rows(calc, "K3", [(10,), (20,), (30,), (40,)])
    calc.Range("L2").Formula = "=I2*Rate"
    calc.Range("K2:L6").Table(None, calc.Range("I2"))
    return excel.save(wb, FIXTURES / "Formulas.xlsx")


def preservation(excel: Excel) -> Path:
    """A macro-enabled book with the parts xlsxedit must carry through untouched."""
    wb = excel.new_workbook("Data", "Report", "Pivot")
    data, report, pivot = (wb.Worksheets(n) for n in ("Data", "Report", "Pivot"))
    regions, products = ("North", "South", "East", "West"), ("Apples", "Pears", "Plums")
    sales = [
        (regions[i % 4], products[i % 3], dt.date(2026, 1 + i % 6, 1), 40 + (i * 37) % 160)
        for i in range(20)
    ]
    _rows(data, "A1", [("Region", "Product", "Month", "Sales")] + sales)
    data.Range("C2:C21").NumberFormat = "mmm yyyy"
    data.Activate()
    data.Range("A2").Select()
    wb.Windows(1).FreezePanes = True  # View > Freeze Panes with A2 selected
    data.Range("A1:D21").AutoFilter(4, ">100")

    _rows(
        report,
        "A1",
        [("Product", "Jan", "Feb", "Mar", "Apr", "Trend")]
        + [(p, 10 + i, 14 + i, 9 + 2 * i, 17 - i, None) for i, p in enumerate(products)]
        + [("Total", None, None, None, None, None)],
    )
    report.Range("B5:E5").Formula = "=SUM(B2:B4)"
    report.Range("F2:F4").SparklineGroups.Add(xl.SparkLine, "Report!B2:E4")
    report.Range("F5").SparklineGroups.Add(xl.SparkColumn, "Report!B5:E5")
    report.Range("A2").AddComment("Legacy note: prices exclude tax.")
    report.Range("A7").Value = "Detail"
    _rows(report, "A8", [("North", 1), ("South", 2), ("East", 3)])
    report.Rows("8:10").Group()
    report.Outline.ShowLevels(1)
    report.Range("A12").Value = "Hidden row below"
    report.Rows(13).Hidden = True
    report.Range("A13").Value = "hidden"

    cache = wb.PivotCaches().Create(xl.Database, data.Range("A1:D21"))
    table = cache.CreatePivotTable(pivot.Range("A3"), "SalesPivot")
    table.PivotFields("Region").Orientation = xl.RowField
    table.AddDataField(table.PivotFields("Sales"), "Total sales", xl.Sum)
    left, top = _at(pivot, "E3")
    slicers = wb.SlicerCaches.Add2(table, "Product")
    slicers.Slicers.Add(pivot, pythoncom.Missing, "Product", "Product", top, left, 144, 160)

    wb.CustomXMLParts.Add(
        '<fixture xmlns="urn:xlsxedit:fixtures"><name>Preservation</name></fixture>'
    )
    report.Range("A4").AddCommentThreaded("Threaded comment: check the plum figures.")
    _add_vba_module(wb)
    return excel.save(wb, FIXTURES / "Preservation.xlsm")


VBA_MODULE = """Public Function Doubled(ByVal n As Double) As Double
    Doubled = n * 2
End Function
"""


def _add_vba_module(wb) -> None:
    """Add a standard module, which makes Excel write ``xl/vbaProject.bin``.

    Needs "Trust access to the VBA project object model" (Excel Options >
    Trust Center > Macro Settings). This script never changes that setting;
    without it the module is skipped and the README gives the manual step.
    """
    try:
        module = wb.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
    except pythoncom.com_error:
        print(
            "  VBA project access is not trusted: Preservation.xlsm gets no vbaProject.bin "
            "(see scripts/excel_fixtures/README.md)"
        )
        return
    module.Name = "FixtureMacros"
    module.CodeModule.AddFromString(VBA_MODULE)


BUILDERS = {
    "TableRich": table_rich,
    "TableWithContentBelow": table_with_content_below,
    "DataValidation": data_validation,
    "ConditionalFormattingRich": conditional_formatting_rich,
    "DrawingAnchors": drawing_anchors,
    "Formulas": formulas,
    "Preservation": preservation,
}


def check_private() -> int:
    """Scan every fixture for the user's names and local paths, without starting Excel."""
    needles = personal_strings()
    paths = sorted(
        p for p in FIXTURES.rglob("*") if p.suffix.lower() in (".xlsx", ".xlsm")
    )
    status = 0
    for path in paths:
        leaks = personal_data(path, needles)
        if leaks:
            print(f"{path.relative_to(FIXTURES.parent)}: {', '.join(leaks)}")
            status = 1
    print(f"checked {len(paths)} workbooks" if status == 0 else "personal data found")
    return status


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument(
        "names", nargs="*", metavar="NAME", help=f"fixtures to build (default all): {', '.join(BUILDERS)}"
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="only scan every fixture for the user's names and local paths",
    )
    args = parser.parse_args(argv)
    if args.check:
        return check_private()
    unknown = sorted(set(args.names) - BUILDERS.keys())
    if unknown:
        parser.error(f"unknown fixture: {', '.join(unknown)}")
    names = args.names or list(BUILDERS)
    with Excel() as excel:
        print(f"Excel {excel.version}; {excel.platform}")
        for name in names:
            path = BUILDERS[name](excel)
            print(f"wrote {path.relative_to(FIXTURES.parent)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

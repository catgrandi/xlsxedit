"""New rows are placed through the row index, not by scanning ``sheetData`` (#42).

Counting the walks over a sheet, not timing them, keeps an insert from growing
quadratic again without making the suite depend on the machine's speed.
"""

from __future__ import annotations

import pytest
from lxml import etree

from xlsxedit import Workbook, row_shift
from xlsxedit.opc.constants import SML_NS
from xlsxedit.worksheet import Worksheet
from tests.preservation import check_consistency, read_pkg

SHEET = "xl/worksheets/sheet1.xml"
_ROW = f"{{{SML_NS}}}row"


class _Walks:
    """Count every pass over a sheet's rows that an insert can make."""

    def __init__(self, monkeypatch):
        self.counts = {"scan": 0, "index": 0, "sort": 0, "walk": 0, "dimension": 0}
        self._count(monkeypatch, Worksheet, "_place_row_by_scan", "scan")
        self._count(monkeypatch, row_shift, "_walk_sheet_data", "walk")
        self._count(monkeypatch, Worksheet, "update_dimension", "dimension")
        build = Worksheet._build_bulk_indexes

        def build_index(ws):
            if not ws._bulk_indexes_built:
                self.counts["index"] += 1
            build(ws)

        monkeypatch.setattr(Worksheet, "_build_bulk_indexes", build_index)
        place = Worksheet._place_row_after_indexed

        def place_after(ws, *args):
            unsorted = ws._row_numbers is None
            placed = place(ws, *args)
            if unsorted and ws._row_numbers is not None:
                self.counts["sort"] += 1
            return placed

        monkeypatch.setattr(Worksheet, "_place_row_after_indexed", place_after)

    def _count(self, monkeypatch, owner, name: str, key: str) -> None:
        real = getattr(owner, name)

        def counted(*args, **kwargs):
            self.counts[key] += 1
            return real(*args, **kwargs)

        monkeypatch.setattr(owner, name, counted)


def _sheet(rows: int) -> Worksheet:
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.write_rows([[n, f"row {n}"] for n in range(1, rows + 1)], at_row=1)
    return ws


def _row_numbers(ws: Worksheet) -> list[int]:
    pkg = read_pkg(ws._workbook)
    check_consistency(pkg)
    sheet_data = etree.fromstring(pkg[SHEET]).find(f"{{{SML_NS}}}sheetData")
    return [int(row.get("r")) for row in sheet_data.iterchildren(_ROW)]


@pytest.mark.parametrize("at_row", [1, 150, 300, 301])
def test_insert_walks_the_sheet_a_fixed_number_of_times(monkeypatch, at_row: int):
    counts = []
    for inserted in (5, 200):
        ws = _sheet(300)
        with monkeypatch.context() as patched:
            walks = _Walks(patched)
            ws.insert_rows([[n, "new"] for n in range(inserted)], at_row=at_row)
        counts.append(walks.counts)
        assert _row_numbers(ws) == list(range(1, 301 + inserted))
        assert [ws[f"A{at_row + n}"].value for n in range(inserted)] == list(range(inserted))
    assert counts[0] == counts[1]
    assert counts[0]["scan"] == 0


def test_rows_written_in_any_order_are_placed_without_a_scan(monkeypatch):
    ws = _sheet(0)
    ws.write_rows([[n] for n in range(0, 1000, 10)], at_row=1)
    walks = _Walks(monkeypatch)
    for row in [500, 250, 750, 251, 2, 1001, 499, 125]:
        ws[f"B{row}"].value = row
    assert walks.counts["scan"] == 0
    assert walks.counts["sort"] <= 1
    assert _row_numbers(ws) == sorted({*range(1, 101), 125, 250, 251, 499, 500, 750, 1001})


def test_a_worksheet_held_across_add_worksheet_shares_the_row_index(monkeypatch):
    wb = Workbook.create()
    old = wb["Sheet1"]
    for n in (1, 5, 9):
        old[f"A{n}"].value = n
    wb.add_worksheet("Other")
    new = wb["Sheet1"]
    assert new is old
    walks = _Walks(monkeypatch)

    new["A3"].value = 3
    old["B4"].value = "old 4"
    new["A6"].value = 6
    old["B6"].value = "old 6"
    assert walks.counts["scan"] == 0
    assert _row_numbers(old) == [1, 3, 4, 5, 6, 9]
    assert [new[a].value for a in ("A3", "B4", "A6", "B6")] == [3, "old 4", 6, "old 6"]


def test_rows_added_to_the_xml_behind_the_index_are_respected(monkeypatch):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    for n in (1, 5, 9):
        ws[f"A{n}"].value = n
    rows = {int(row.get("r")): row for row in ws._sheet_data().iterchildren(_ROW)}
    for n, before in ((3, 5), (6, 9)):
        row = etree.Element(_ROW, r=str(n))
        etree.SubElement(row, f"{{{SML_NS}}}c", r=f"A{n}")
        rows[before].addprevious(row)
    walks = _Walks(monkeypatch)

    ws["B4"].value = 4  # the index puts row 4 after row 1, where row 3 now is
    ws["B6"].value = 6  # found once the scan above has rebuilt the index
    assert walks.counts["scan"] == 1
    assert _row_numbers(ws) == [1, 3, 4, 5, 6, 9]
    assert [ws[a].value for a in ("B4", "B6")] == [4, 6]


def test_rows_a_bulk_write_appends_are_indexed_on_first_lookup():
    ws = _sheet(0)
    ws.write_rows([[n, n * 2, f"s{n}"] for n in range(1, 501)], at_row=1)
    assert ws._cell_index == {}
    assert ws["B250"].value == 500
    assert set(ws._cell_index) == {"A250", "B250", "C250"}
    ws["D250"].value = "new"
    assert ws.clear_range("A10:C10") == 3
    ws.write_rows([["again"]], at_cell="C250")
    assert [ws[a].value for a in ("A10", "C250", "D250", "C251")] == [None, "again", "new", "s251"]
    assert _row_numbers(ws) == list(range(1, 501))


def test_cells_after_a_merge_anchor_in_a_new_row_are_found():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.merge_cells("A5:C5")
    ws.write_rows([["x", "y", "z", "w", "v"]], at_cell="B5")  # B5 and C5 go to the anchor A5
    assert ws._find_cell_element("D5") is not None
    assert ws.clear_range("D5:F5") == 3
    assert [ws[a].value for a in ("A5", "D5", "E5", "F5")] == ["y", None, None, None]

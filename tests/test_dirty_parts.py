"""Write paths mark the parts they change (#42), so saving them skips a second parse."""

from __future__ import annotations

import io

import pytest

from xlsxedit import Workbook
from xlsxedit.opc import part as part_module
from xlsxedit.opc.constants import CT
from xlsxedit.opc.part import XmlPart
from tests.conftest import BOOK1, INSPECT_FIXTURES
from tests.preservation import assert_preserved, content_types, read_pkg, worksheet_members

SHEET = "Sheet1"


def _saved_bytes(wb: Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _member(pkg: dict[str, bytes], content_type: str) -> str:
    return next(m for m, ct in content_types(pkg).items() if ct == content_type)


def _dirty_parts(wb: Workbook) -> list[str]:
    return [
        str(p.partname)
        for p in wb._package.iter_parts()
        if isinstance(p, XmlPart) and p.is_dirty
    ]


@pytest.mark.parametrize("value", [42, "probe"], ids=["number", "string"])
def test_saving_an_edited_sheet_does_not_parse_it_again(monkeypatch, value):
    before = read_pkg(BOOK1)
    sheet = worksheet_members(before)[SHEET]
    sst = _member(before, CT.SHARED_STRINGS)
    wb = Workbook.open(BOOK1)
    wb[SHEET]["Z99"].value = value

    parsed: list[bytes] = []
    real = part_module.parse_xml
    monkeypatch.setattr(part_module, "parse_xml", lambda blob: parsed.append(blob) or real(blob))
    after = read_pkg(_saved_bytes(wb))

    assert before[sheet] not in parsed
    if isinstance(value, str):
        assert before[sst] not in parsed
    assert after[sheet] != before[sheet]
    assert Workbook.open(io.BytesIO(_saved_bytes(wb)))[SHEET]["Z99"].value == value


def test_an_edited_part_releases_its_source_bytes():
    wb = Workbook.open(BOOK1)
    ws = wb[SHEET]
    ws["Z99"].value = 42
    assert ws._part.is_dirty
    assert ws._part._orig_blob is None


@pytest.mark.parametrize("large", [False, True], ids=["eager", "deferred"])
@pytest.mark.parametrize("name", INSPECT_FIXTURES)
def test_reading_marks_no_part_dirty(name: str, large: bool):
    wb = Workbook.open(INSPECT_FIXTURES[name], large=large)
    for ws in wb.worksheets:
        ws.values()
        for cell in ws.cells:
            cell.value, cell.formula, cell.style.num_format
        ws.merged_ranges, ws.images, ws.charts, ws.tables
    assert _dirty_parts(wb) == []


def _filled() -> bytes:
    """Book1 after API writes of a number, a string, a row of values, a row
    height, a column width and a number format, saved once."""
    wb = Workbook.open(BOOK1)
    ws = wb[SHEET]
    ws["Z99"].value = 42
    ws["Z100"].value = "probe"
    ws.write_rows([[1, 2.5, True, None]], at_row=120)
    ws.row_dimensions[120].height = 30
    ws.column_dimensions["Z"].width = 20
    ws["Z99"].apply_number_format("0.00")
    return _saved_bytes(wb)


@pytest.mark.parametrize(
    "rewrite",
    [
        lambda ws: setattr(ws["Z99"], "value", 42),
        lambda ws: setattr(ws["Z100"], "value", "probe"),
        lambda ws: ws.write_rows([[1, 2.5, True, None]], at_row=120),
        lambda ws: setattr(ws.row_dimensions[120], "height", 30),
        lambda ws: setattr(ws.column_dimensions["Z"], "width", 20),
        lambda ws: ws.clear_range("D120:D120"),
        lambda ws: ws["Z99"].apply_number_format("0.00"),
    ],
    ids=["number", "string", "write_rows", "row height", "column width", "clear empty", "format"],
)
def test_writing_what_a_sheet_already_holds_keeps_its_bytes(rewrite):
    """A setter given the value a cell already has changes nothing, so it marks nothing."""
    source = _filled()
    wb = Workbook.open(io.BytesIO(source))
    ws = wb[SHEET]
    rewrite(ws)
    assert _dirty_parts(wb) == []
    report = assert_preserved(read_pkg(source), read_pkg(_saved_bytes(wb)))
    assert not report.reserialised
    assert read_pkg(_saved_bytes(wb)) == read_pkg(source)

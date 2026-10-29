"""Write paths mark the parts they change (#42), so saving them skips a second parse."""

from __future__ import annotations

import inspect
import io

import pytest
from lxml import etree

import xlsxedit
from xlsxedit import Workbook, Worksheet
from xlsxedit.opc import part as part_module
from xlsxedit.opc.constants import CT, SML_NS
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


ALL_FIXTURES = {"Book1": BOOK1, **INSPECT_FIXTURES}
_API_TYPES = tuple(
    obj for obj in vars(xlsxedit).values() if isinstance(obj, type) and obj.__name__ in xlsxedit.__all__
)


def _read_properties(obj, seen: dict[int, object]) -> None:
    """Read every public property of ``obj``, and of the API objects and lists they
    return, except the way back to a ``Workbook`` or ``Worksheet``."""
    if id(obj) in seen:
        return
    seen[id(obj)] = obj  # kept alive, so that its id is not reused
    if isinstance(obj, (list, tuple)):
        for item in obj:
            _read_properties(item, seen)
        return
    if not isinstance(obj, _API_TYPES) or isinstance(obj, (Workbook, Worksheet)):
        return
    for name, _ in inspect.getmembers(type(obj), lambda member: isinstance(member, property)):
        if not name.startswith("_"):
            _read_properties(getattr(obj, name), seen)


def _read_everything(wb: Workbook) -> None:
    """Every public read API, including ``ws[address]`` on cells the sheet does not have."""
    seen: dict[int, object] = {}
    wb.sheetnames, wb.orphan_partnames, wb.shared_strings
    _read_properties(wb.properties, seen)
    first = next((c.value for ws in wb for c in ws.iter_cells() if c.value is not None), "x")
    assert wb.find(first) is not None or isinstance(first, float)
    wb.findall(first), wb.find("absent from every fixture")
    for ws in wb.worksheets:
        assert wb[ws.name] is ws
        ws.values(), ws.values(errors_as_nan=True, max_rows=3), list(ws.iter_rows())
        ws.find(first), ws.findall(first), ws.merged_ranges
        for prop in ("cells", "tables", "data_validations", "conditional_formatting"):
            _read_properties(getattr(ws, prop), seen)
        for prop in ("drawing_objects", "images", "charts"):
            _read_properties(getattr(ws, prop), seen)
        addresses = {"A1", "B2", "XFD1", "A1048576", "XFD1048576"}
        for cell in ws.iter_cells():
            addresses |= {cell.address, cell.offset(cols=1).address, cell.offset(rows=1).address}
        for address in sorted(addresses):
            cell = ws[address]
            _read_properties(cell, seen)
            ws.column_dimensions[address.rstrip("0123456789")].width
            ws.row_dimensions[int(address.lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZ"))].height


@pytest.mark.parametrize("large", [False, True], ids=["eager", "deferred"])
@pytest.mark.parametrize("name", ALL_FIXTURES)
def test_reading_marks_no_part_dirty(name: str, large: bool):
    """Reading through every public read API changes no part, so the save that
    follows writes every member as the source holds it (#44)."""
    source = read_pkg(ALL_FIXTURES[name])
    wb = Workbook.open(ALL_FIXTURES[name], large=large)
    _read_everything(wb)
    assert _dirty_parts(wb) == []
    assert read_pkg(_saved_bytes(wb)) == source


@pytest.mark.parametrize("large", [False, True], ids=["eager", "deferred"])
@pytest.mark.parametrize("name", ALL_FIXTURES)
def test_the_pandas_reader_changes_no_member(name: str, large: bool):
    pd = pytest.importorskip("pandas")
    from xlsxedit import pandas_io

    pandas_io.register()
    source = ALL_FIXTURES[name].read_bytes()
    with pd.ExcelFile(io.BytesIO(source), engine="xlsxedit", engine_kwargs={"large": large}) as xl:
        pd.read_excel(xl, sheet_name=None, header=None)
        assert _dirty_parts(xl.book) == []
        assert read_pkg(_saved_bytes(xl.book)) == read_pkg(source)


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


def test_replacing_in_an_inline_string_marks_the_sheet_only_on_a_match():
    wb = Workbook.open(BOOK1)
    ws = wb[SHEET]
    c_elm = ws._get_or_create_cell_element("Z99")
    c_elm.set("t", "inlineStr")
    c_elm.append(etree.fromstring(f'<is xmlns="{SML_NS}"><t>hello name</t></is>'))
    assert ws["Z99"].replace("absent", "x") == 0
    assert not ws._part.is_dirty
    assert ws["Z99"].replace("name", "world") == 1
    assert ws._part.is_dirty
    assert Workbook.open(io.BytesIO(_saved_bytes(wb)))[SHEET]["Z99"].value == "hello world"


def _with_a_validation() -> Workbook:
    wb = Workbook.create()
    wb[SHEET].add_data_validation("B2:B10", type="list", formula1='"a,b"')
    return Workbook.open(io.BytesIO(_saved_bytes(wb)))


def test_data_validation_writes_mark_the_sheet_only_on_a_change():
    wb = _with_a_validation()
    ws = wb[SHEET]
    (dv,) = ws.data_validations
    dv.sqref = "B2:B10"
    assert _dirty_parts(wb) == []
    dv.sqref = "C2:C10"
    assert ws._part.is_dirty

    wb = _with_a_validation()
    ws = wb[SHEET]
    with pytest.raises(Exception):
        ws.add_data_validation("B5", type="list", formula1='"c"')  # overlaps, refused
    assert _dirty_parts(wb) == []
    ws.add_data_validation("D2:D10", type="list", formula1='"c"')
    assert ws._part.is_dirty

    wb = _with_a_validation()
    ws = wb[SHEET]
    ws.remove_data_validation(ws.data_validations[0])
    assert ws._part.is_dirty

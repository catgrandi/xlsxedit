"""The preservation harness itself: canonical form, package and worksheet diffs."""

from __future__ import annotations

import io
import os
import zipfile

import pytest
from lxml import etree

from xlsxedit import Workbook
from tests.conftest import BOOK1, INSPECT_FIXTURES
from tests.preservation import (
    CONTENT_TYPES,
    CellState,
    assert_preserved,
    canon,
    child_order,
    content_types,
    read_pkg,
    worksheet_diff,
)

SHEET = "xl/worksheets/sheet1.xml"

EXCEL = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    b'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" '
    b'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
    b'<sheetData><row r="1"><c r="A1" t="inlineStr">'
    b'<is><t xml:space="preserve"> a </t></is></c></row></sheetData></worksheet>'
)


def _replace(blob: bytes, old: bytes, new: bytes) -> bytes:
    assert blob.count(old) == 1, f"{old!r} occurs {blob.count(old)} times"
    return blob.replace(old, new)


def _simple_formula_sheet() -> bytes:
    return read_pkg(INSPECT_FIXTURES["SimpleFormula"])[SHEET]


def test_canon_forgives_lxml_reserialisation():
    lxml_bytes = etree.tostring(
        etree.fromstring(EXCEL), xml_declaration=True, encoding="UTF-8", standalone=True
    )
    assert lxml_bytes.startswith(b"<?xml version='1.0'")
    assert lxml_bytes.index(b"xmlns:x14ac") < lxml_bytes.index(b"mc:Ignorable")
    assert canon(SHEET, lxml_bytes) == canon(SHEET, EXCEL)


_X14AC = b' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"'


@pytest.mark.parametrize(
    "edits",
    [
        [(_X14AC, b"")],
        [(_X14AC, b""), (b"<sheetData>", b"<sheetData" + _X14AC + b">")],
        [(b'xml:space="preserve"> a </t>', b'xml:space="preserve">a</t>')],
        [(b"<sheetData>", b"<!-- note --><sheetData>")],
    ],
    ids=[
        "ignorable-declaration-dropped",
        "declaration-moved-to-child",
        "preserved-space",
        "comment",
    ],
)
def test_canon_keeps_what_excel_reads(edits: list[tuple[bytes, bytes]]):
    """Inclusive C14N 1.0; ``etree.canonicalize`` (C14N 2.0) would forgive the first two."""
    edited = EXCEL
    for old, new in edits:
        edited = _replace(edited, old, new)
    assert canon(SHEET, edited) != canon(SHEET, EXCEL)


def test_canon_classifies_members_by_extension():
    single = EXCEL.replace(
        b'"1.0" encoding="UTF-8" standalone="yes"', b"'1.0' encoding='UTF-8' standalone='yes'"
    )
    for name in ("_rels/.rels", "xl/drawings/vmlDrawing1.vml", "xl/TABLES/TABLE1.XML"):
        assert canon(name, single) == canon(name, EXCEL), name
    assert canon("xl/media/image1.png", single) == single
    assert canon(SHEET, b"<worksheet>") == b"<worksheet>"


def test_rewritten_opaque_part_counts_as_reserialised_not_changed():
    """The library keeps table parts as bytes until touched, then writes them with lxml."""
    before = read_pkg(INSPECT_FIXTURES["ChartsAndTables"])
    table = "xl/tables/table1.xml"
    after = {**before, table: etree.tostring(etree.fromstring(before[table]), xml_declaration=True)}
    assert after[table] != before[table]
    assert assert_preserved(before, after).reserialised == {table}


def test_read_pkg_accepts_every_package_source(tmp_path):
    expected = read_pkg(BOOK1)
    with zipfile.ZipFile(BOOK1) as z:
        z.extractall(tmp_path)
    (tmp_path / ".DS_Store").write_bytes(b"\0")
    (tmp_path / "__pycache__").mkdir()
    (tmp_path / "__pycache__" / "x.pyc").write_bytes(b"\0")
    assert read_pkg(BOOK1.read_bytes()) == expected
    assert read_pkg(io.BytesIO(BOOK1.read_bytes())) == expected
    assert read_pkg(os.fspath(tmp_path)) == expected
    assert read_pkg(Workbook.open(BOOK1)).keys() == expected.keys()


def test_read_pkg_skips_directories_and_rejects_duplicate_members():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("xl/", b"")
        z.writestr("xl/a.xml", b"<a/>")
    assert read_pkg(buf.getvalue()) == {"xl/a.xml": b"<a/>"}
    with zipfile.ZipFile(buf, "a") as z, pytest.warns(UserWarning, match="Duplicate name"):
        z.writestr("xl/a.xml", b"<b/>")
    with pytest.raises(ValueError, match="duplicate ZIP members: xl/a.xml"):
        read_pkg(buf.getvalue())


def test_content_types_resolve_override_then_default():
    types = content_types(
        {
            CONTENT_TYPES: (
                b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                b'<Default Extension="XML" ContentType="application/xml"/>'
                b'<Default Extension="png" ContentType="image/png"/>'
                b'<Override PartName="/XL/Workbook.xml" ContentType="wb"/></Types>'
            ),
            "xl/workbook.xml": b"",
            "xl/styles.xml": b"",
            "xl/media/image1.PNG": b"",
            "xl/printerSettings1.bin": b"",
        }
    )
    assert types == {
        "xl/workbook.xml": "wb",
        "xl/styles.xml": "application/xml",
        "xl/media/image1.PNG": "image/png",
        "xl/printerSettings1.bin": None,
    }


def _edited_book1() -> tuple[dict[str, bytes], dict[str, bytes]]:
    before = read_pkg(BOOK1)
    after = dict(before)
    after["xl/worksheets/sheet2.xml"] = _replace(
        after["xl/worksheets/sheet2.xml"],
        b"<sheetData>",
        b'<sheetData><row r="40"><c r="B40"><v>7</v></c></row>',
    )
    after["customXml/item1.xml"] = b"<root/>"
    del after["docProps/app.xml"]
    return before, after


def test_assert_preserved_reports_exact_member_sets():
    before, after = _edited_book1()
    report = assert_preserved(
        before,
        after,
        expected_changed={"xl/worksheets/sheet2.xml"},
        expected_added={"customXml/item1.xml"},
        expected_removed={"docProps/app.xml"},
    )
    assert report.changed == {"xl/worksheets/sheet2.xml"}
    assert report.added == {"customXml/item1.xml"}
    assert report.removed == {"docProps/app.xml"}
    assert report.reserialised == set()


def test_assert_preserved_names_every_surprise():
    before, after = _edited_book1()
    with pytest.raises(AssertionError) as excinfo:
        assert_preserved(before, after, expected_changed={"xl/worksheets/sheet9.xml"})
    message = str(excinfo.value)
    assert "changed but not expected: xl/worksheets/sheet2.xml" in message
    assert "added but not expected: customXml/item1.xml" in message
    assert "removed but not expected: docProps/app.xml" in message
    assert "expected changed but was not: xl/worksheets/sheet9.xml (in neither package)" in message
    assert "B40 CellState(t=None, s=None, formula=None, value='7')" in message


def test_changes_worksheet_diff_does_not_model_are_still_described():
    before = read_pkg(BOOK1)
    after = {**before, SHEET: _replace(before[SHEET], b"\n<worksheet", b"\n<!-- x --><worksheet")}
    with pytest.raises(AssertionError, match="first difference at canonical offset"):
        assert_preserved(before, after)


def test_content_types_are_compared_as_an_effective_map():
    src = INSPECT_FIXTURES["images"]
    report = assert_preserved(src, Workbook.open(src))
    assert CONTENT_TYPES in report.reserialised  # jpeg/jpg Defaults became Overrides

    before = read_pkg(BOOK1)
    after = dict(before)
    after[CONTENT_TYPES] = _replace(
        after[CONTENT_TYPES],
        b"spreadsheetml.sheet.main+xml",
        b"spreadsheetml.template.main+xml",
    )
    assert_preserved(before, after, expected_changed={CONTENT_TYPES})
    with pytest.raises(
        AssertionError, match=r"xl/workbook.xml: \S+sheet.main\+xml -> \S+template.main\+xml"
    ):
        assert_preserved(before, after)


def test_added_members_need_a_content_type():
    before = read_pkg(BOOK1)
    after = {**before, "xl/embeddings/blob.weird": b"\x00"}
    with pytest.raises(
        AssertionError, match="added member has no content type: xl/embeddings/blob.weird"
    ):
        assert_preserved(before, after, expected_added={"xl/embeddings/blob.weird"})


def test_changed_xml_must_stay_well_formed():
    before = read_pkg(BOOK1)
    after = {**before, SHEET: b"<worksheet>"}
    with pytest.raises(AssertionError, match="not well-formed XML: xl/worksheets/sheet1.xml"):
        assert_preserved(before, after, expected_changed={SHEET})


def test_worksheet_diff_cells():
    before = _simple_formula_sheet()
    after = _replace(before, b'<c r="A3"><v>19</v>', b'<c r="A3"><v>20</v>')
    after = _replace(after, b'<c r="B8" t="s"><v>3</v></c>', b"")
    after = _replace(
        after,
        b"<v>2</v></c></row>",
        b'<v>2</v></c><c r="D2" t="inlineStr"><is><t>new</t></is></c></row>',
    )
    after = _replace(after, b'ref="C4:C7"', b'ref="C4:C8"')
    diff = worksheet_diff(before, after)
    assert diff.added == {"D2": CellState("inlineStr", None, None, "new")}
    assert diff.removed == {"B8": CellState("s", None, None, "3")}
    master = CellState(None, None, "A4*B4", "689430")
    assert diff.changed == {
        "A3": (CellState(None, None, None, "19"), CellState(None, None, None, "20")),
        "C4": (master, master),
    }
    assert 'ref="C4:C7"' in diff.summary() and 'ref="C4:C8"' in diff.summary()
    assert not (diff.rows or diff.children or diff.root or diff.order)


def test_worksheet_diff_rows_children_root_and_order():
    before = _simple_formula_sheet()
    after = _replace(before, b'<row r="2" spans="1:3"', b'<row r="2" spans="1:4"')
    after = _replace(after, b'<pageMargins left="0.7"', b'<pageMargins left="0.5"')
    after = _replace(
        after, b"{B19D5468-4E5F-2C42-860C-9883D90D3DF2}", b"{00000000-0000-0000-0000-000000000000}"
    )
    after = _replace(after, b'<dimension ref="A2:C8"/>', b"")
    after = _replace(after, b"<pageMargins", b'<dimension ref="A2:C8"/><pageMargins')
    after = _replace(after, b"<worksheet ", b'<worksheet xmlns:foo="urn:foo" ')
    diff = worksheet_diff(before, after)
    assert not (diff.added or diff.removed or diff.changed)
    row = {"r": "2", "spans": "1:3", "x14ac:dyDescent": "0.2"}
    assert diff.rows == {2: (row, {**row, "spans": "1:4"})}
    assert list(diff.children) == [("pageMargins", 0)]  # the new root namespace is not repeated
    assert diff.root[0]["xr:uid"] != diff.root[1]["xr:uid"]
    assert "xmlns:foo" not in diff.root[0] and diff.root[1]["xmlns:foo"] == "urn:foo"
    assert diff.order == (
        ("dimension", "sheetViews", "sheetFormatPr", "sheetData", "pageMargins"),
        ("sheetViews", "sheetFormatPr", "sheetData", "dimension", "pageMargins"),
    )


def test_worksheet_diff_places_cells_without_r_and_rejects_duplicates():
    empty = (
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b"<sheetData/></worksheet>"
    )
    implicit = empty.replace(
        b"<sheetData/>", b"<sheetData><row><c><v>1</v></c><c><v>2</v></c></row></sheetData>"
    )
    diff = worksheet_diff(empty, implicit)
    assert list(diff.added) == ["A1", "B1"]
    assert diff.rows == {1: (None, {})}
    assert not worksheet_diff(implicit, implicit)
    assert worksheet_diff(implicit, implicit).summary() == "(no differences)"
    with pytest.raises(ValueError, match="two rows numbered 1"):
        worksheet_diff(
            implicit, implicit.replace(b"</sheetData>", b'<row r="1"><c r="C1"/></row></sheetData>')
        )
    with pytest.raises(ValueError, match="two cells at A1"):
        worksheet_diff(
            implicit, implicit.replace(b"</sheetData>", b'<row r="2"><c r="A1"/></row></sheetData>')
        )


def test_child_order():
    assert child_order(_simple_formula_sheet()) == (
        "dimension",
        "sheetViews",
        "sheetFormatPr",
        "sheetData",
        "pageMargins",
    )

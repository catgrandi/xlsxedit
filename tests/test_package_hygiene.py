"""Package hygiene (#8): sheet-removal cleanup, part names, content types, byte-stable parts."""

from __future__ import annotations

import io
import zipfile

import pytest
from lxml import etree

from xlsxedit import Workbook
from xlsxedit.opc.constants import CT, SML_NS
from xlsxedit.opc.package import OpcPackage
from xlsxedit.opc.packuri import PackURI, encode_partname, is_partname
from xlsxedit.opc.part import Part
from xlsxedit.opc.pkgwriter import PackageWriter
from xlsxedit.opc.serialize import XML_DECLARATION
from tests.conftest import BOOK1, INSPECT_FIXTURES
from tests.preservation import CONTENT_TYPES, assert_preserved, content_types, read_pkg
from tests.schema_validate import assert_valid_package

CHARTS = INSPECT_FIXTURES["ChartsAndTables"]
WORKBOOK = "xl/workbook.xml"
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_DEFINED_NAME = f"{{{SML_NS}}}definedName"
_WORKBOOK_VIEW = f"{{{SML_NS}}}bookViews/{{{SML_NS}}}workbookView"

# what removing "bar chart", the first ChartsAndTables sheet, must take with it
BAR_CHART_PARTS = frozenset(
    {
        "xl/worksheets/sheet1.xml",
        "xl/worksheets/_rels/sheet1.xml.rels",
        "xl/drawings/drawing1.xml",
        "xl/drawings/_rels/drawing1.xml.rels",
        "xl/charts/chart1.xml",
        "xl/charts/_rels/chart1.xml.rels",
        "xl/charts/style1.xml",
        "xl/charts/colors1.xml",
    }
)


def _edit(pkg: dict[str, bytes], member: str, old: bytes, new: bytes) -> dict[str, bytes]:
    assert pkg[member].count(old) == 1, f"{old!r} not found once in {member}"
    pkg[member] = pkg[member].replace(old, new)
    return pkg


def _zip(pkg: dict[str, bytes]) -> io.BytesIO:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for name, data in pkg.items():
            z.writestr(name, data)
    buf.seek(0)
    return buf


def _saved(wb: Workbook, **kwargs) -> dict[str, bytes]:
    buf = io.BytesIO()
    wb.save(buf, **kwargs)
    return read_pkg(buf.getvalue())


def _charts_with_chain_and_local_names() -> dict[str, bytes]:
    """ChartsAndTables plus a formula and calcChain entry on "bar chart", and a
    ``_xlnm.Print_Area`` scoped to each of the three sheets."""
    pkg = read_pkg(CHARTS)
    _edit(
        pkg,
        "xl/worksheets/sheet1.xml",
        b'<c r="B6"><v>0</v></c>',
        b'<c r="B6"><f>SUM(B3:B5)</f><v>11008</v></c>',
    )
    pkg["xl/calcChain.xml"] = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        b'<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b'<c r="B6" i="1"/></calcChain>'
    )
    _edit(
        pkg,
        WORKBOOK_RELS,
        b"</Relationships>",
        f'<Relationship Id="rId99" Type="{_RT}/calcChain" Target="calcChain.xml"/>'.encode()
        + b"</Relationships>",
    )
    _edit(
        pkg,
        CONTENT_TYPES,
        b"</Types>",
        f'<Override PartName="/xl/calcChain.xml" ContentType="{CT.CALC_CHAIN}"/>'.encode()
        + b"</Types>",
    )
    local_names = "".join(
        f'<definedName name="_xlnm.Print_Area" localSheetId="{i}">{ref}</definedName>'
        for i, ref in enumerate(("'bar chart'!$A$1:$D$6", "'line chart'!$A$1:$B$12", "Table!$A$1:$C$11"))
    )
    _edit(pkg, WORKBOOK, b"<definedNames>", b"<definedNames>" + local_names.encode())
    return pkg


def _local_names(pkg: dict[str, bytes]) -> list[tuple[str | None, str]]:
    root = etree.fromstring(pkg[WORKBOOK])
    return [(elm.get("localSheetId"), elm.text) for elm in root.iter(_DEFINED_NAME)]


def test_remove_worksheet_drops_what_only_that_sheet_reached():
    before = _charts_with_chain_and_local_names()
    wb = Workbook.open(_zip(before))
    wb.remove_worksheet("bar chart")
    after = _saved(wb)

    assert_preserved(
        before,
        after,
        expected_changed={WORKBOOK, WORKBOOK_RELS},
        expected_removed=BAR_CHART_PARTS | {"xl/calcChain.xml"},
    )
    assert _local_names(after)[:2] == [
        ("0", "'line chart'!$A$1:$B$12"),
        ("1", "Table!$A$1:$C$11"),
    ]
    assert all(local is None for local, _ in _local_names(after)[2:])
    assert etree.fromstring(after[WORKBOOK]).find(_WORKBOOK_VIEW).get("activeTab") == "1"
    assert b"/xl/worksheets/sheet1.xml" not in after[CONTENT_TYPES]
    assert_valid_package(_zip(after))


def test_remove_worksheet_keep_unreachable_leaves_the_parts_in_the_package():
    before = _charts_with_chain_and_local_names()
    wb = Workbook.open(_zip(before))
    wb.remove_worksheet("bar chart", keep_unreachable=True)
    after = _saved(wb)

    kept = BAR_CHART_PARTS - {"xl/worksheets/sheet1.xml", "xl/worksheets/_rels/sheet1.xml.rels"}
    assert kept <= after.keys()
    assert {"xl/worksheets/sheet1.xml", "xl/calcChain.xml"}.isdisjoint(after)
    reopened = Workbook.open(_zip(after))
    assert set(reopened.orphan_partnames) >= {"/xl/drawings/drawing1.xml", "/xl/charts/chart1.xml"}


def test_remove_worksheet_drops_the_only_definednames_block():
    pkg = _edit(
        read_pkg(INSPECT_FIXTURES["TestBook"]),
        WORKBOOK,
        b"</sheets>",
        b'</sheets><definedNames><definedName name="x" localSheetId="0">Sheet1!$A$1'
        b"</definedName></definedNames>",
    )
    wb = Workbook.open(_zip(pkg))
    wb.remove_worksheet("Sheet1")
    assert etree.fromstring(_saved(wb)[WORKBOOK]).find(f"{{{SML_NS}}}definedNames") is None


@pytest.mark.parametrize(
    ("removed", "active_tab"),
    [("bar chart", "1"), ("line chart", "1"), ("Table", "1")],
    ids=["before-active", "before-active-adjacent", "active-and-last"],
)
def test_remove_worksheet_keeps_the_active_tab_in_range(removed: str, active_tab: str):
    wb = Workbook.open(CHARTS)
    wb.remove_worksheet(removed)
    assert etree.fromstring(_saved(wb)[WORKBOOK]).find(_WORKBOOK_VIEW).get("activeTab") == active_tab


def test_remove_worksheet_keeps_parts_another_sheet_shares():
    """Copy a sheet whose drawing shares an image, then remove the copy: the
    package returns to its exact source bytes, image included."""
    src = INSPECT_FIXTURES["images"]
    before = read_pkg(src)
    wb = Workbook.open(src)
    wb.copy_worksheet(wb.sheetnames[0], "copy")
    wb.remove_worksheet("copy")
    report = assert_preserved(before, _saved(wb))
    assert not report.reserialised


def test_copy_worksheet_drops_the_calc_chain():
    src = INSPECT_FIXTURES["SimpleFormula"]
    before = read_pkg(src)
    wb = Workbook.open(src)
    wb.copy_worksheet("Sheet1", "Copy")
    assert_preserved(
        before,
        _saved(wb),
        expected_changed={WORKBOOK, WORKBOOK_RELS},
        expected_added={"xl/worksheets/sheet2.xml"},
        expected_removed={"xl/calcChain.xml"},
    )


def test_prune_unreachable_returns_the_dropped_parts_and_keeps_orphans():
    pkg = read_pkg(CHARTS)
    pkg["custom/cargo.bin"] = b"cargo"
    package = OpcPackage.open(_zip(pkg))
    sheet1 = package.main_document_part().rels["rId1"].target_part
    sheet1.rels._rels.clear()
    assert sorted(package.prune_unreachable()) == [
        "/xl/charts/chart1.xml",
        "/xl/charts/colors1.xml",
        "/xl/charts/style1.xml",
        "/xl/drawings/drawing1.xml",
    ]
    assert package.prune_unreachable() == []
    assert package.orphan_partnames == ("/custom/cargo.bin",)


def test_next_partname_skips_orphan_names_case_insensitively():
    pkg = read_pkg(BOOK1)
    pkg["xl/worksheets/sheet3.xml"] = b"orphan three"
    pkg["xl/worksheets/SHEET4.xml"] = b"orphan four"
    wb = Workbook.open(_zip(pkg))
    wb.add_worksheet("New")
    assert wb["New"]._part.partname == "/xl/worksheets/sheet5.xml"

    after = _saved(wb, include_orphans=True)
    assert after["xl/worksheets/sheet3.xml"] == b"orphan three"
    assert after["xl/worksheets/SHEET4.xml"] == b"orphan four"
    assert Workbook.open(_zip(after)).sheetnames == ["Sheet1", "test", "New"]


def test_package_writer_refuses_duplicate_part_names(tmp_path):
    parts = [
        Part(PackURI("/xl/media/image1.png"), CT.PNG, b"a"),
        Part(PackURI("/xl/media/IMAGE1.png"), CT.PNG, b"b"),
    ]
    out = tmp_path / "out.xlsx"
    with pytest.raises(ValueError, match="duplicate part names: /xl/media/image1.png"):
        PackageWriter.write(out, OpcPackage().rels, parts)
    assert not out.exists()


def test_regenerated_content_types_keep_the_source_defaults():
    src = INSPECT_FIXTURES["images"]
    before = read_pkg(src)
    wb = Workbook.open(src)
    wb.add_worksheet("New")
    after = _saved(wb)
    report = assert_preserved(
        before,
        after,
        expected_changed={WORKBOOK, WORKBOOK_RELS},
        expected_added={"xl/worksheets/sheet3.xml"},
    )
    assert report.reserialised == {CONTENT_TYPES}
    root = etree.fromstring(after[CONTENT_TYPES])
    defaults = [(e.get("Extension"), e.get("ContentType")) for e in root.iterchildren("{*}Default")]
    assert defaults == [
        ("jpeg", "image/jpeg"),
        ("jpg", "image/jpeg"),
        ("rels", CT.OPC_RELATIONSHIPS),
        ("xml", CT.XML),
    ]
    overridden = {e.get("PartName") for e in root.iterchildren("{*}Override")}
    assert {"/xl/media/image1.jpeg", "/xl/media/image2.jpg"}.isdisjoint(overridden)
    assert "/xl/worksheets/sheet3.xml" in overridden


def test_untouched_parts_keep_their_bytes_after_a_reverted_edit():
    before = read_pkg(BOOK1)
    wb = Workbook.open(BOOK1)
    cell = wb["Sheet1"]["A1"]._element
    cell.set("probe", "1")
    del cell.attrib["probe"]
    report = assert_preserved(before, _saved(wb))
    assert not report.reserialised


@pytest.mark.parametrize("large", [False, True], ids=["eager", "deferred"])
def test_read_only_access_keeps_every_byte(large: bool):
    before = read_pkg(CHARTS)
    wb = Workbook.open(CHARTS, large=large)
    assert [cell.value for ws in wb.worksheets for cell in ws.cells]
    report = assert_preserved(before, _saved(wb))
    assert not report.reserialised


def test_rewritten_parts_use_excels_xml_declaration():
    wb = Workbook.open(BOOK1)
    wb["Sheet1"]["Z99"].value = 1
    after = _saved(wb)
    assert after["xl/worksheets/sheet1.xml"].startswith(XML_DECLARATION)


def test_mark_dirty_rewrites_an_unchanged_part():
    before = read_pkg(BOOK1)
    wb = Workbook.open(BOOK1)
    wb._workbook_part.styles_part.mark_dirty()
    after = _saved(wb)
    assert assert_preserved(before, after).reserialised == {"xl/styles.xml"}
    assert after["xl/styles.xml"].startswith(XML_DECLARATION)


def test_part_without_a_declared_content_type_loads_and_saves():
    pkg = _edit(
        read_pkg(INSPECT_FIXTURES["images"]),
        CONTENT_TYPES,
        b'<Default Extension="jpeg" ContentType="image/jpeg"/>',
        b"",
    )
    wb = Workbook.open(_zip(pkg))
    after = _saved(wb)
    assert content_types(after)["xl/media/image1.jpeg"] == "application/octet-stream"
    assert after["xl/media/image1.jpeg"] == pkg["xl/media/image1.jpeg"]


@pytest.mark.parametrize(
    ("name", "legal"),
    [
        ("/xl/workbook.xml", True),
        ("/xl/media/image%201.png", True),
        ("/[trash]/0000.dat", False),
        ("/xl/a b.xml", False),
        ("/xl//a.xml", False),
        ("/xl/a./b.xml", False),
        ("/xl/100%.xml", False),
        ("xl/workbook.xml", False),
    ],
)
def test_is_partname(name: str, legal: bool):
    assert is_partname(name) is legal


@pytest.mark.parametrize(
    ("name", "encoded"),
    [
        ("/[trash]/0000.dat", "/%5Btrash%5D/0000.dat"),
        ("/xl/a b%20c.xml", "/xl/a%20b%20c.xml"),
        ("/xl/100%.xml", "/xl/100%25.xml"),
        ("/xl/café.xml", "/xl/caf%C3%A9.xml"),
        ("/xl/a./b.xml", None),
        ("/xl//a.xml", None),
    ],
)
def test_encode_partname(name: str, encoded: str | None):
    assert encode_partname(name) == encoded


def test_orphans_with_illegal_names_are_encoded_or_skipped():
    pkg = read_pkg(BOOK1)
    pkg["[trash]/0000.dat"] = b"trash"
    pkg["bad./cargo.bin"] = b"unencodable"
    wb = Workbook.open(_zip(pkg))
    assert wb.orphan_partnames == ("/%5Btrash%5D/0000.dat",)

    after = _saved(wb, include_orphans=True)
    assert after["%5Btrash%5D/0000.dat"] == b"trash"
    assert {"[trash]/0000.dat", "bad./cargo.bin"}.isdisjoint(after)
    assert content_types(after)["%5Btrash%5D/0000.dat"] == "application/octet-stream"
    assert Workbook.open(_zip(after)).orphan_partnames == ("/%5Btrash%5D/0000.dat",)


def test_orphans_that_would_share_a_part_name_are_skipped():
    pkg = read_pkg(BOOK1)
    pkg["%5Btrash%5D/0001.dat"] = b"legal name"
    pkg["[trash]/0001.dat"] = b"encodes to the legal name above"
    pkg["XL/WORKBOOK.XML"] = b"case variant of a related part"
    pkg["custom/a.bin"] = b"first"
    pkg["custom/A.bin"] = b"case variant of an earlier orphan"
    wb = Workbook.open(_zip(pkg))
    assert wb.orphan_partnames == ("/%5Btrash%5D/0001.dat", "/custom/a.bin")
    after = _saved(wb, include_orphans=True)
    assert after["%5Btrash%5D/0001.dat"] == b"legal name"
    assert after["custom/a.bin"] == b"first"

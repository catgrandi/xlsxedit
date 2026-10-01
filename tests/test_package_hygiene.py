"""Package hygiene (#8): sheet-removal cleanup, part names, content types, byte-stable parts."""

from __future__ import annotations

import io
import zipfile

import pytest
from lxml import etree

from xlsxedit import Workbook
from xlsxedit.opc.constants import CT
from xlsxedit.opc.package import OpcPackage
from xlsxedit.opc.packuri import PackURI, encode_partname, is_partname
from xlsxedit.opc.part import Part
from xlsxedit.opc.pkgwriter import PackageWriter
from xlsxedit.opc.serialize import XML_DECLARATION
from tests.conftest import BOOK1, INSPECT_FIXTURES
from tests.preservation import CONTENT_TYPES, assert_preserved, content_types, read_pkg

CHARTS = INSPECT_FIXTURES["ChartsAndTables"]
WORKBOOK = "xl/workbook.xml"
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"


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

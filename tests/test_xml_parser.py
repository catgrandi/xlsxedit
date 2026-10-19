"""XML parser hardening: no entity expansion, no network access, no DTDs."""

from __future__ import annotations

import io
import socket
import threading
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from xlsxedit import DTDForbiddenError, Workbook, XlsxeditError
from xlsxedit.oxml import parse_template_xml, parse_xml
from xlsxedit.oxml.parser import _PARSER, _TEMPLATE_PARSER, refuse_dtd_in_prolog

SST = "xl/sharedStrings.xml"
SHEET = "xl/worksheets/sheet1.xml"
PARSERS = {"parts": _PARSER, "templates": _TEMPLATE_PARSER}


@pytest.fixture
def secret(tmp_path: Path) -> Path:
    path = tmp_path / "secret.txt"
    path.write_text("TOPSECRET", encoding="utf-8")
    return path


def _hostile(member: str, doctype: str, text: str = "PLACEHOLDER") -> io.BytesIO:
    """A workbook whose ``member`` declares ``doctype``; A1's string becomes ``text``."""
    wb = Workbook.create()
    wb["Sheet1"]["A1"].value = "PLACEHOLDER"
    buf = io.BytesIO()
    wb.save(buf)
    out = io.BytesIO()
    with zipfile.ZipFile(buf) as src, zipfile.ZipFile(out, "w") as dst:
        for info in src.infolist():
            data = src.read(info)
            if info.filename == member:
                declaration, end, rest = data.partition(b"?>")
                rest = rest.replace(b"PLACEHOLDER", text.encode())
                data = declaration + end + doctype.encode() + rest
            dst.writestr(info, data)
    out.seek(0)
    return out


def test_xxe_payload_leaks_through_a_permissive_parser(secret: Path):
    xml = f'<!DOCTYPE t [<!ENTITY xxe SYSTEM "{secret.as_uri()}">]><t>&xxe;</t>'
    root = etree.fromstring(xml.encode(), etree.XMLParser(resolve_entities=True))
    assert "TOPSECRET" in root.text


@pytest.mark.parametrize("name", PARSERS)
def test_parser_leaves_entities_unexpanded(name: str, secret: Path):
    xml = (
        f'<!DOCTYPE t [<!ENTITY xxe SYSTEM "{secret.as_uri()}"><!ENTITY int "INJECTED">]>'
        "<t>a&xxe;b&int;c</t>"
    )
    root = etree.fromstring(xml.encode(), PARSERS[name])
    assert etree.tostring(root) == b"<t>a&xxe;b&int;c</t>"


def test_entity_bomb_is_never_expanded():
    # Expanded, &l6; would be 10**6 copies of "lol".
    entities = '<!ENTITY l0 "lol">' + "".join(
        f'<!ENTITY l{i} "{f"&l{i - 1};" * 10}">' for i in range(1, 7)
    )
    xml = f"<!DOCTYPE t [{entities}]><t>&l6;</t>".encode()
    try:
        root = etree.fromstring(xml, _PARSER)
    except etree.XMLSyntaxError as exc:
        assert "amplification" in str(exc)  # libxml2 2.11+ refuses the bomb outright
    else:
        assert etree.tostring(root) == b"<t>&l6;</t>"


def test_parser_still_accepts_deeply_nested_xml():
    depth = 300  # libxml2 refuses more than 256 levels unless huge_tree is set
    root = parse_xml("<a>" * depth + "</a>" * depth)
    assert sum(1 for _ in root.iter()) == depth


@pytest.mark.parametrize("parse", [parse_xml, parse_template_xml])
def test_any_dtd_is_refused(parse):
    with pytest.raises(DTDForbiddenError, match=r"<!DOCTYPE a>"):
        parse("<!DOCTYPE a><a/>")


def test_dtd_forbidden_error_is_a_value_error():
    assert issubclass(DTDForbiddenError, XlsxeditError)
    assert issubclass(DTDForbiddenError, ValueError)


def test_workbook_with_an_xxe_payload_is_refused(secret: Path):
    doctype = f'<!DOCTYPE sst [<!ENTITY xxe SYSTEM "{secret.as_uri()}">]>'
    with pytest.raises(DTDForbiddenError) as excinfo:
        Workbook.open(_hostile(SST, doctype, "&xxe;"))
    assert "TOPSECRET" not in str(excinfo.value)


def test_workbook_with_remote_entities_makes_no_connection():
    connections: list[tuple] = []
    stop = threading.Event()
    with socket.create_server(("127.0.0.1", 0)) as server:
        server.settimeout(0.05)
        url = f"http://127.0.0.1:{server.getsockname()[1]}"

        def accept() -> None:
            while not stop.is_set():
                try:
                    conn, peer = server.accept()
                except OSError:
                    continue
                connections.append(peer)
                conn.close()

        listener = threading.Thread(target=accept, daemon=True)
        listener.start()
        doctype = (
            f'<!DOCTYPE sst SYSTEM "{url}/sst.dtd" [<!ENTITY % remote SYSTEM "{url}/r.dtd">'
            f' %remote; <!ENTITY xxe SYSTEM "{url}/secret">]>'
        )
        try:
            with pytest.raises(DTDForbiddenError):
                Workbook.open(_hostile(SST, doctype, "&xxe;"))
        finally:
            stop.set()
            listener.join()
    assert connections == []


def test_dtd_in_a_worksheet_is_refused_once_the_sheet_is_parsed():
    with pytest.raises(DTDForbiddenError):
        Workbook.open(_hostile(SHEET, "<!DOCTYPE worksheet>"))["Sheet1"]["A1"].value


CHARTS = Path(__file__).resolve().parents[1] / "fixtures" / "ChartsAndTables.xlsx"


def _with_member(source, member: str, change) -> io.BytesIO:
    """A copy of the package ``source`` whose ``member`` holds ``change(bytes)``."""
    out = io.BytesIO()
    with zipfile.ZipFile(source) as src, zipfile.ZipFile(out, "w") as dst:
        for info in src.infolist():
            data = src.read(info)
            dst.writestr(info, change(data) if info.filename == member else data)
    out.seek(0)
    return out


def _add_doctype(doctype: str):
    def change(data: bytes) -> bytes:
        declaration, end, rest = data.partition(b"?>")
        if not end:
            return doctype.encode() + data
        return declaration + end + doctype.encode() + rest

    return change


@pytest.mark.parametrize(
    ("member", "large"),
    [
        ("xl/tables/table1.xml", False),
        ("xl/drawings/drawing1.xml", False),
        ("xl/charts/chart1.xml", False),
        ("xl/worksheets/sheet1.xml", True),
    ],
)
def test_dtd_in_a_part_kept_as_bytes_is_refused_at_open(member: str, large: bool):
    hostile = _with_member(CHARTS, member, _add_doctype("<!DOCTYPE root>"))
    with pytest.raises(DTDForbiddenError, match=member):
        Workbook.open(hostile, large=large)


def test_dtd_after_a_bom_comments_and_instructions_is_refused_at_open():
    prolog = '\ufeff<?xml version="1.0"?>\n<!-- a <!DOCTYPE in a comment --><?pi x?>\n<!DOCTYPE c:chartSpace>'

    def change(data: bytes) -> bytes:
        return prolog.encode("utf-8") + data.partition(b"?>")[2]

    with pytest.raises(DTDForbiddenError, match=r"<!DOCTYPE c:chartSpace>"):
        Workbook.open(_with_member(CHARTS, "xl/charts/chart1.xml", change))


@pytest.mark.parametrize("codec", ["utf-16", "utf-16-le", "utf-16-be"])
def test_dtd_in_a_utf16_part_is_refused_at_open(codec: str):
    def change(data: bytes) -> bytes:
        text = data.decode("utf-8").replace('encoding="UTF-8"', 'encoding="UTF-16"')
        declaration, end, rest = text.partition("?>")
        return (declaration + end + "<!DOCTYPE table>" + rest).encode(codec)

    with pytest.raises(DTDForbiddenError, match="xl/tables/table1.xml"):
        Workbook.open(_with_member(CHARTS, "xl/tables/table1.xml", change))


def test_utf16_part_without_a_dtd_still_opens():
    def change(data: bytes) -> bytes:
        return data.decode("utf-8").replace('encoding="UTF-8"', 'encoding="UTF-16"').encode("utf-16")

    wb = Workbook.open(_with_member(CHARTS, "xl/tables/table1.xml", change))
    assert [t.name for ws in wb.worksheets for t in ws.tables]


@pytest.mark.parametrize(
    "xml",
    [
        b"<a/>",
        b"\xef\xbb\xbf<a/>",
        b"<!-- comment -->\n<a><![CDATA[<!DOCTYPE a>]]></a>",
        b"<?xml version='1.0'?><a><!-- <!DOCTYPE a> --></a>",
        b"<a>" + b"<!DOCTYPE a>" + b"</a>",
        b"",
        b"   ",
    ],
)
def test_prolog_without_a_dtd_passes(xml: bytes):
    refuse_dtd_in_prolog(xml, "/part.xml")


@pytest.mark.parametrize(
    "xml",
    [
        "<!DOCTYPE a><a/>",
        "\ufeff<!DOCTYPE a><a/>",
        "<?xml version='1.0'?>\n\n<!--\n" + "x" * 100_000 + "\n--><?pi?> <!DOCTYPE a [<!ENTITY e 'x'>]><a/>",
    ],
    ids=["bare", "bom", "after-a-long-comment"],
)
@pytest.mark.parametrize("codec", ["utf-8", "utf-16", "utf-16-le", "utf-16-be", "utf-32"])
def test_prolog_with_a_dtd_is_refused(xml: str, codec: str):
    if codec in ("utf-16", "utf-32"):  # these codecs write their own byte-order mark
        xml = xml.removeprefix("\ufeff")
    with pytest.raises(DTDForbiddenError, match=r"/part\.xml declares a DTD \(<!DOCTYPE a>\)"):
        refuse_dtd_in_prolog(xml.encode(codec), "/part.xml")


def test_dtd_in_a_part_xlsxedit_never_parses_is_kept():
    """A related part of a type xlsxedit only carries as bytes may keep its DTD."""
    item = b'<?xml version="1.0"?><!DOCTYPE data><data/>'

    def add_rel(data: bytes) -> bytes:
        rel = (
            b'<Relationship Id="rIdCustom" Target="../customXml/item1.xml" Type='
            b'"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"/>'
        )
        return data.replace(b"</Relationships>", rel + b"</Relationships>")

    source = _with_member(CHARTS, "xl/_rels/workbook.xml.rels", add_rel)
    out = io.BytesIO()
    with zipfile.ZipFile(source) as src, zipfile.ZipFile(out, "w") as dst:
        for info in src.infolist():
            dst.writestr(info, src.read(info))
        dst.writestr("customXml/item1.xml", item)
    out.seek(0)
    wb = Workbook.open(out)
    saved = io.BytesIO()
    wb.save(saved)
    with zipfile.ZipFile(saved) as z:
        assert z.read("customXml/item1.xml") == item

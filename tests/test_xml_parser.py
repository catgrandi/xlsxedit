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
from xlsxedit.oxml.parser import _PARSER, _TEMPLATE_PARSER

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

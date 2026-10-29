"""``xlsxedit._ooxml_ext`` records each extension URI with its payload and host.

An entry marked ``verified`` must occur in a workbook in ``fixtures/`` with
the payload and host it records; an unverified one must occur in none, so a
new fixture that holds it fails here until the entry is marked verified.
"""

from __future__ import annotations

from collections import defaultdict
from functools import lru_cache
from pathlib import Path

from lxml import etree

from tests.conftest import FIXTURES
from tests.preservation import read_pkg
from xlsxedit._ooxml_ext import EXTENSIONS, extension
from xlsxedit._ooxml_ns import PREFIXES, SCHEMA_FILES

SCHEMAS = Path(__file__).resolve().parents[1] / "schemas"
_XSD = "{http://www.w3.org/2001/XMLSchema}"
_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True)


@lru_cache(maxsize=1)
def _fixture_extensions() -> dict[str, frozenset[tuple[str, str]]]:
    """``{uri as written: {(host, payload)}}`` for every ``extLst/ext`` in the fixtures."""
    found: dict[str, set[tuple[str, str]]] = defaultdict(set)
    for path in sorted(FIXTURES.rglob("*.xls[xm]")):
        for name, blob in read_pkg(path).items():
            if not name.endswith(".xml"):
                continue
            for ext in etree.fromstring(blob, _PARSER).iter("{*}ext"):
                ext_list = ext.getparent()
                if ext.get("uri") is None or etree.QName(ext_list).localname != "extLst":
                    continue
                host = ext_list.getparent().tag
                for payload in ext.iterchildren(etree.Element):
                    found[ext.get("uri")].add((host, payload.tag))
    return {uri: frozenset(pairs) for uri, pairs in found.items()}


@lru_cache(maxsize=None)
def _global_elements(filename: str) -> frozenset[str]:
    root = etree.parse(str(SCHEMAS / filename)).getroot()
    return frozenset(node.get("name") for node in root.iterchildren(f"{_XSD}element"))


def test_every_fixture_uri_is_recorded_as_excel_spells_it():
    missing = sorted(set(_fixture_extensions()) - set(EXTENSIONS))
    assert not missing, f"add to src/xlsxedit/_ooxml_ext.py: {missing}"


def test_verified_entries_occur_in_the_fixtures():
    found = _fixture_extensions()
    for uri, ext in EXTENSIONS.items():
        if ext.verified:
            assert (ext.host, f"{{{ext.namespace}}}{ext.element}") in found.get(uri, ()), uri


def test_unverified_entries_occur_in_no_fixture():
    held = {uri.upper() for uri in _fixture_extensions()}
    stale = [uri for uri, ext in EXTENSIONS.items() if not ext.verified and uri.upper() in held]
    assert not stale, f"a fixture holds these; mark them verified: {stale}"


def test_entries_name_known_namespaces_and_schema_elements():
    for uri, ext in EXTENSIONS.items():
        assert ext.namespace in PREFIXES, uri
        assert etree.QName(ext.host).namespace in PREFIXES, uri
        filename = SCHEMA_FILES.get(ext.namespace)
        if filename is not None:
            assert ext.element in _global_elements(filename), uri


def test_uris_differ_beyond_case():
    assert len({uri.upper() for uri in EXTENSIONS}) == len(EXTENSIONS)


def test_lookup_ignores_case():
    dv = EXTENSIONS["{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"]
    assert extension("{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}") is dv
    assert extension("{cce6a557-97bc-4b89-adb6-d9c93caab3df}") is dv
    assert extension("{00000000-0000-0000-0000-000000000000}") is None

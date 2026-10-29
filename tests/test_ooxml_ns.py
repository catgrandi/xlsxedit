"""``xlsxedit._ooxml_ns`` and the schema notes are generated and must not drift.

The prefixes are checked against every workbook in ``fixtures/``: Excel wrote
them, so a namespace they bind to another prefix is recorded wrongly.
"""

from __future__ import annotations

import ast
import importlib.util
from collections import defaultdict
from functools import lru_cache
from pathlib import Path

import pytest
from lxml import etree

from tests.conftest import FIXTURES
from tests.preservation import read_pkg
from tests.schema_validate import extension_namespaces
from xlsxedit import _ooxml_ns
from xlsxedit._ooxml_ns import PREFIXES, SCHEMA_FILES
from xlsxedit.opc import constants

ROOT = Path(__file__).resolve().parents[1]
GENERATOR = ROOT / "scripts" / "gen_ooxml_order.py"

# The custom XML item that scripts/excel_fixtures put into Preservation.xlsm.
_NOT_EXCEL = frozenset({"urn:xlsxedit:fixtures"})
_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True)


@pytest.fixture(scope="module")
def generator():
    spec = importlib.util.spec_from_file_location("gen_ooxml_order", GENERATOR)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_generated_module_and_notes_match_schemas(generator):
    assert generator.NS_OUTPUT.read_text(encoding="utf-8") == generator.render_namespaces()
    notes = generator.NOTES_OUTPUT.read_text(encoding="utf-8")
    assert notes == generator.render_notes(), "run: python scripts/gen_ooxml_order.py"


def test_check_covers_every_generated_file(generator, monkeypatch):
    outputs = generator.OUTPUTS
    paths = [path for path, _ in outputs]
    assert paths == [generator.OUTPUT, generator.NS_OUTPUT, generator.NOTES_OUTPUT]
    for stale in paths:
        committed = stale.read_bytes()
        patched = tuple(
            (path, (lambda: "stale") if path == stale else render) for path, render in outputs
        )
        monkeypatch.setattr(generator, "OUTPUTS", patched)
        assert generator.main(["--check"]) == 1, stale
        assert stale.read_bytes() == committed


def test_module_reads_no_file():
    """The package ships without ``schemas/``, so the namespaces must be literals."""
    tree = ast.parse(Path(_ooxml_ns.__file__).read_text(encoding="utf-8"))
    imports = [node for node in tree.body if isinstance(node, (ast.Import, ast.ImportFrom))]
    assert [ast.unparse(node) for node in imports] == ["from __future__ import annotations"]


def test_one_constant_per_namespace():
    values = [value for name, value in vars(_ooxml_ns).items() if name.endswith("_NS")]
    assert len(values) == len(set(values))
    assert set(values) == set(PREFIXES)


def test_vendored_namespaces_come_from_their_schemas():
    for namespace, filename in SCHEMA_FILES.items():
        root = etree.parse(str(ROOT / "schemas" / filename)).getroot()
        assert root.get("targetNamespace") == namespace
    assert extension_namespaces() <= SCHEMA_FILES.keys()


def test_opc_constants_keep_their_names():
    assert constants.SML_NS == _ooxml_ns.SML_NS
    assert constants.REL_NS == _ooxml_ns.REL_NS
    assert constants.CT_NS == _ooxml_ns.CT_NS
    assert constants.OFFICE_REL_NS == _ooxml_ns.OFFICE_REL_NS
    assert constants.NSMAP == {
        "ct": _ooxml_ns.CT_NS,
        "pr": _ooxml_ns.REL_NS,
        "main": _ooxml_ns.SML_NS,
        "r": _ooxml_ns.OFFICE_REL_NS,
    }


@lru_cache(maxsize=1)
def _fixture_bindings() -> dict[str, frozenset[str | None]]:
    """Every namespace an XML member of a fixture declares, with the prefixes bound to it."""
    found: dict[str, set[str | None]] = defaultdict(set)
    for path in sorted(FIXTURES.rglob("*.xls[xm]")):
        for name, blob in read_pkg(path).items():
            if not name.endswith((".xml", ".rels", ".vml")):
                continue
            for elm in etree.fromstring(blob, _PARSER).iter(etree.Element):
                for prefix, namespace in elm.nsmap.items():
                    found[namespace].add(prefix)
    return {namespace: frozenset(prefixes) for namespace, prefixes in found.items()}


def test_every_namespace_excel_writes_is_named():
    unknown = sorted(set(_fixture_bindings()) - set(PREFIXES) - _NOT_EXCEL)
    assert not unknown, f"add to OTHER_NAMESPACES in scripts/gen_ooxml_order.py: {unknown}"


def test_prefixes_are_the_ones_excel_declares():
    for namespace, prefixes in _fixture_bindings().items():
        if namespace in _NOT_EXCEL:
            continue
        expected = {PREFIXES[namespace]} - {None}
        assert prefixes - {None} <= expected, (namespace, sorted(prefixes - {None}))

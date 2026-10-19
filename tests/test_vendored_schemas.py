"""The vendored schemas match ``schemas/README.md`` and ``scripts/fetch_ms_xlsx_schemas.py``.

Nothing here uses the network: the fetch script's pure functions run on the
committed files and on literal pages.
"""

from __future__ import annotations

import hashlib
import importlib.util
import re
from pathlib import Path

import pytest
from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
FETCHER = ROOT / "scripts" / "fetch_ms_xlsx_schemas.py"
README = ROOT / "schemas" / "README.md"

_XSD = "{http://www.w3.org/2001/XMLSchema}"
_HASH_ROW = re.compile(r"^\| `([^`]+\.xsd)` \|(?:[^\n]*\|)? `([0-9a-f]{64})` \|$", re.M)


@pytest.fixture(scope="module")
def fetcher():
    spec = importlib.util.spec_from_file_location("fetch_ms_xlsx_schemas", FETCHER)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_readme_records_the_hash_of_every_vendored_schema(fetcher):
    vendored = sorted((ROOT / "schemas" / "transitional").glob("*.xsd"))
    vendored += [fetcher.VENDORED / source.filename for source in fetcher.SOURCES]
    recorded = dict(_HASH_ROW.findall(README.read_text(encoding="utf-8")))
    assert set(recorded) == {path.name for path in vendored}
    for path in vendored:
        assert hashlib.sha256(path.read_bytes()).hexdigest() == recorded[path.name], path.name


def test_aliases_and_driver_match_the_vendored_schemas(fetcher):
    files = fetcher.generated_files(fetcher.committed_schemas())
    for path, text in files.items():
        assert path.read_bytes().decode("utf-8") == text, "run: python scripts/fetch_ms_xlsx_schemas.py"
    vendored = {path for path in files if path.parent == fetcher.VENDORED}
    assert set(fetcher.VENDORED.glob("*.xsd")) == vendored


def test_every_schema_location_exists():
    for path in (ROOT / "schemas").rglob("*.xsd"):
        root = etree.parse(str(path)).getroot()
        for node in root.iterchildren(f"{_XSD}import", f"{_XSD}include"):
            target = path.parent / node.get("schemaLocation")
            assert target.is_file(), f"{path.relative_to(ROOT)} -> {node.get('schemaLocation')}"


def test_an_import_of_an_unlisted_namespace_is_refused(fetcher):
    schemas = fetcher.committed_schemas()
    name = fetcher.SOURCES[0].filename
    schemas[name] = schemas[name].replace(
        "</xsd:schema>", '<xsd:import namespace="urn:new" schemaLocation="new.xsd"/></xsd:schema>'
    )
    with pytest.raises(LookupError, match="urn:new"):
        fetcher.generated_files(schemas)
    schemas[name] = schemas[name].replace('namespace="urn:new" ', "")
    with pytest.raises(ValueError, match="without a namespace"):
        fetcher.generated_files(schemas)


def test_extract_schema_decodes_the_pre_block(fetcher):
    page = (
        "<html><body><pre>&nbsp;&lt;xsd:schema a=&quot;1&quot;&gt;  \r\n"
        "&nbsp;&nbsp; &lt;xsd:element/&gt;\r\n"
        "&nbsp;&lt;/xsd:schema&gt;\n\n</pre></body></html>"
    )
    assert fetcher.extract_schema(page) == '<xsd:schema a="1">\n  <xsd:element/>\n</xsd:schema>\n'


def test_extract_schema_ignores_blank_lines_when_removing_indentation(fetcher):
    lines = ["&nbsp;" * 4 + " &lt;a&gt;", "&nbsp;", "&nbsp;" * 6 + " &lt;b/&gt;", "&nbsp;" * 4 + " &lt;/a&gt;"]
    expected = "<a>\n\n  <b/>\n</a>\n"
    for newline in ("\n", "\r\n"):
        assert fetcher.extract_schema(f"<pre>{newline.join(lines)}</pre>") == expected


@pytest.mark.parametrize(
    "page, message",
    [
        ("<p>no schema</p>", "found 0"),
        ("<pre>&lt;a/&gt;</pre><pre>&lt;b/&gt;</pre>", "found 2"),
        ("<pre>&lt;a&gt;<span>x</span>&lt;/a&gt;</pre>", "markup"),
    ],
)
def test_extract_schema_refuses_an_unexpected_page(fetcher, page, message):
    with pytest.raises(ValueError, match=message):
        fetcher.extract_schema(page)


def test_mirror_comparison_ignores_only_whitespace(fetcher):
    ours = fetcher._normalised_lines('<a x="1">\n  <b/>\n</a>\n')
    assert fetcher._normalised_lines('\ufeff\xa0<a  x="1"><b/></a>') == ours
    assert fetcher._normalised_lines('<a x="2"><b/></a>') != ours

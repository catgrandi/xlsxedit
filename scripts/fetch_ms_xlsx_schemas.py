#!/usr/bin/env python3
"""Vendor Microsoft's extension schemas from [MS-XLSX] Appendix A (maintainer tool).

Without arguments, download the Appendix A page of every namespace in
SOURCES, extract its schema, and write it under schemas/microsoft/ms-xlsx/,
together with the alias files that resolve the names Microsoft's schemas
import and the driver schemas/microsoft/extensions.xsd. Then print the
provenance rows for schemas/README.md.

With --check, download the pages and exit with status 1 when a committed
file differs from what they produce. With --verify, compare each committed
schema with its copy in the OpenXLIFF mirror at MIRROR_COMMIT, ignoring
whitespace, print every difference, and exit with status 1 if there is one.
"""

from __future__ import annotations

import argparse
import difflib
import hashlib
import html
import re
import sys
import textwrap
import urllib.request
from pathlib import Path
from typing import NamedTuple

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
MICROSOFT = ROOT / "schemas" / "microsoft"
VENDORED = MICROSOFT / "ms-xlsx"
DRIVER = MICROSOFT / "extensions.xsd"
SHIM = MICROSOFT / "sml-shim.xsd"

PAGE_BASE = "https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/"
MIRROR_COMMIT = "f3f1c1e5e7758b7a0e9287ac707c6e6f7aa6fcd5"
MIRROR_BASE = (
    "https://raw.githubusercontent.com/maxprograms-com/OpenXLIFF/"
    f"{MIRROR_COMMIT}/catalog/openxml/MS-XLSX-AppendixA/"
)

_XSD_NS = "http://www.w3.org/2001/XMLSchema"
_SML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_MS_PREFIX = "http://schemas.microsoft.com/office/"


class Source(NamedTuple):
    section: str
    namespace: str
    page: str
    mirror: str

    @property
    def filename(self) -> str:
        return self.namespace.removeprefix(_MS_PREFIX).replace("/", "-") + ".xsd"

    @property
    def url(self) -> str:
        return PAGE_BASE + self.page


# x14, x15, xm and x14ac, and every Microsoft namespace they import.
SOURCES: tuple[Source, ...] = (
    Source(
        "5.1",
        "http://schemas.microsoft.com/office/excel/2006/main",
        "b0ddba06-ac73-4d3d-a913-8fbe4a56cf28",
        "5_1_excel_2006_main.xsd",
    ),
    Source(
        "5.3",
        "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
        "e42bbfd7-2a3d-4308-a4f3-30313fc506b9",
        "5_3_spreadsheetml_2010_11_main.xsd",
    ),
    Source(
        "5.4",
        "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
        "0a377581-c743-4ace-bfcd-22f359e70165",
        "5_4_spreadsheetml_2009_9_main.xsd",
    ),
    Source(
        "5.5",
        "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
        "7148c81d-4d6e-4204-be07-a7c09f4277c7",
        "5_5_spreadsheetml_2009_9_ac.xsd",
    ),
    Source(
        "5.10",
        "http://schemas.microsoft.com/office/spreadsheetml/2014/11/main",
        "4fe24a21-6f69-4680-882f-f86b16a69119",
        "5_10_spreadsheetml_2014_11_main.xsd",
    ),
    Source(
        "5.12",
        "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10",
        "8a818a7f-92d0-4390-9855-1a484d8f9347",
        "5_12_spreadsheetml_2016_revision10.xsd",
    ),
    Source(
        "5.14",
        "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6",
        "e47e052c-40d0-43cd-add4-73a220966b69",
        "5_14_spreadsheetml_2016_revision6.xsd",
    ),
    Source(
        "5.15",
        "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
        "5ba37a83-255b-4e2d-ae7b-b4d14d44ad3a",
        "5_15_spreadsheetml_2014_revision.xsd",
    ),
    Source(
        "5.16",
        "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
        "7b729443-d0cd-434e-8d8c-b228f701a55b",
        "5_16_spreadsheetml_2015_revision2.xsd",
    ),
)

# Where an imported name resolves when its namespace has no vendored file,
# relative to VENDORED.
_BASE_TARGETS = {
    _SML_NS: "../sml-shim.xsd",
    _R_NS: "../../transitional/shared-relationshipReference.xsd",
}

_PRE = re.compile(r"<pre\b[^>]*>(.*?)</pre>", re.S)
_UPDATED = re.compile(r'<meta name="updated_at" content="([^"]*)"')
_PARSER = etree.XMLParser(resolve_entities=False, no_network=True)


class Page(NamedTuple):
    schema: str
    updated: str


def _fetch(url: str) -> str:
    request = urllib.request.Request(url, headers={"User-Agent": "xlsxedit-schema-fetch"})
    with urllib.request.urlopen(request, timeout=60) as response:
        return response.read().decode("utf-8")


def extract_schema(page: str) -> str:
    """Return the schema in an Appendix A page.

    The page holds the schema as escaped text in its only ``<pre>`` block,
    indented with no-break spaces. The result decodes the character
    references, turns each no-break space into a space, removes the common
    indentation and trailing whitespace, and ends every line with LF.
    """
    blocks = _PRE.findall(page)
    if len(blocks) != 1:
        raise ValueError(f"expected one <pre> block, found {len(blocks)}")
    if "<" in blocks[0]:
        raise ValueError("the <pre> block contains markup, not only escaped text")
    lines = html.unescape(blocks[0]).replace("\xa0", " ").splitlines()
    text = textwrap.dedent("\n".join(line.rstrip() for line in lines))
    return text.strip("\n") + "\n"


def parse_page(page: str, source: Source) -> Page:
    schema = extract_schema(page)
    root = etree.fromstring(schema.encode("utf-8"), _PARSER)
    if root.tag != f"{{{_XSD_NS}}}schema" or root.get("targetNamespace") != source.namespace:
        raise ValueError(f"section {source.section} is not the schema for {source.namespace}")
    updated = _UPDATED.search(page)
    if updated is None:
        raise ValueError(f"section {source.section} has no updated_at date")
    return Page(schema, updated.group(1)[:10])


def _imports(schema: str) -> list[tuple[str, str]]:
    root = etree.fromstring(schema.encode("utf-8"), _PARSER)
    found = []
    for node in root.iterchildren(f"{{{_XSD_NS}}}import", f"{{{_XSD_NS}}}include"):
        if etree.QName(node).localname == "include":
            namespace = root.get("targetNamespace")
        elif (namespace := node.get("namespace")) is None:
            raise ValueError("an xsd:import without a namespace has no alias target")
        location = node.get("schemaLocation")
        if location is None:
            raise ValueError(f"an import of {namespace} has no schemaLocation")
        found.append((location, namespace))
    return found


def _alias(location: str, namespace: str, target: str) -> str:
    return textwrap.dedent(
        f"""\
        <?xml version="1.0" encoding="UTF-8"?>
        <!-- Generated by scripts/fetch_ms_xlsx_schemas.py. Microsoft's schemas import
             {location}; it stands for {target}. -->
        <xsd:schema xmlns:xsd="{_XSD_NS}"
          targetNamespace="{namespace}">
          <xsd:include schemaLocation="{target}"/>
        </xsd:schema>
        """
    )


def _driver() -> str:
    imports = [f'  <xsd:import namespace="{_SML_NS}" schemaLocation="{SHIM.name}"/>']
    imports += [
        f'  <xsd:import namespace="{source.namespace}"\n'
        f'    schemaLocation="{VENDORED.name}/{source.filename}"/>'
        for source in SOURCES
    ]
    return "\n".join(
        [
            '<?xml version="1.0" encoding="UTF-8"?>',
            "<!-- Generated by scripts/fetch_ms_xlsx_schemas.py. Compiling this file",
            "     loads SpreadsheetML through sml-shim.xsd and every namespace vendored",
            "     in ms-xlsx/. -->",
            f'<xsd:schema xmlns:xsd="{_XSD_NS}">',
            *imports,
            "</xsd:schema>",
            "",
        ]
    )


def generated_files(schemas: dict[str, str]) -> dict[Path, str]:
    """Map every path the script writes to its text, given the vendored schemas.

    ``schemas`` maps each ``Source.filename`` to its extracted text.
    """
    targets = {source.namespace: source.filename for source in SOURCES} | _BASE_TARGETS
    files = {VENDORED / name: text for name, text in schemas.items()}
    aliases: dict[str, str] = {}
    for name, text in sorted(schemas.items()):
        for location, namespace in _imports(text):
            if namespace not in targets:
                raise LookupError(f"{name} imports {namespace}, which SOURCES does not list")
            if aliases.setdefault(location, namespace) != namespace:
                raise ValueError(f"{location} is imported for two namespaces")
    for location, namespace in sorted(aliases.items()):
        path = VENDORED / location
        if path in files:
            raise ValueError(f"the alias {location} would overwrite a vendored schema")
        files[path] = _alias(location, namespace, targets[namespace])
    files[DRIVER] = _driver()
    return files


def committed_schemas() -> dict[str, str]:
    return {
        source.filename: (VENDORED / source.filename).read_text(encoding="utf-8")
        for source in SOURCES
    }


def _stale(files: dict[Path, str]) -> list[Path]:
    return sorted(set(VENDORED.glob("*.xsd")) - set(files))


def _sha256(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _normalised_lines(text: str) -> list[str]:
    """One tag or text run per line, with whitespace runs collapsed."""
    text = text.lstrip("\ufeff").replace("\xa0", " ")
    text = re.sub(r"\s+", " ", re.sub(r">\s+<", "><", text.strip()))
    return re.split(r"(?<=>)(?=<)", text)


def verify() -> int:
    differences = 0
    for source in SOURCES:
        ours = (VENDORED / source.filename).read_text(encoding="utf-8")
        mirror = _fetch(MIRROR_BASE + source.mirror)
        diff = list(
            difflib.unified_diff(
                _normalised_lines(mirror),
                _normalised_lines(ours),
                f"OpenXLIFF {source.mirror}",
                f"ms-xlsx/{source.filename}",
                lineterm="",
            )
        )
        if diff:
            differences += 1
            print("\n".join(diff))
        else:
            print(f"same: ms-xlsx/{source.filename} and OpenXLIFF {source.mirror}")
    print(f"{differences} of {len(SOURCES)} files differ from the mirror at {MIRROR_COMMIT}")
    return 1 if differences else 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument(
        "--check",
        action="store_true",
        help="report committed files that differ from the pages with exit status 1",
    )
    mode.add_argument(
        "--verify",
        action="store_true",
        help="compare the committed schemas with the OpenXLIFF mirror, ignoring whitespace",
    )
    args = parser.parse_args(argv)
    if args.verify:
        return verify()
    pages = {source: parse_page(_fetch(source.url), source) for source in SOURCES}
    files = generated_files({source.filename: page.schema for source, page in pages.items()})
    changed = [
        path
        for path, text in files.items()
        if not path.is_file() or path.read_text(encoding="utf-8") != text
    ]
    stale = _stale(files)
    if args.check:
        for path in changed + stale:
            print(f"{path.relative_to(ROOT).as_posix()} differs from the pages", file=sys.stderr)
        return 1 if changed or stale else 0
    VENDORED.mkdir(parents=True, exist_ok=True)
    for path in changed:
        path.write_text(files[path], encoding="utf-8", newline="\n")
        print(f"wrote {path.relative_to(ROOT).as_posix()}")
    for path in stale:
        path.unlink()
        print(f"removed {path.relative_to(ROOT).as_posix()}")
    print("\n| File | Page | Updated | SHA-256 |\n| --- | --- | --- | --- |")
    for source, page in pages.items():
        print(
            f"| `{source.filename}` | [{source.section}]({source.url}) "
            f"| {page.updated} | `{_sha256(page.schema)}` |"
        )
    return 0


if __name__ == "__main__":
    sys.exit(main())

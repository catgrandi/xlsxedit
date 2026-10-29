#!/usr/bin/env python3
"""Generate the OOXML reference files from the vendored schemas (maintainer tool).

Writes src/xlsxedit/_ooxml_order.py, src/xlsxedit/_ooxml_ns.py and
docs/ooxml-schema-notes.md. Without arguments, rewrite each file that differs
from what the schemas produce. With --check, write nothing and exit with
status 1 when any file differs.
"""

from __future__ import annotations

import argparse
import json
import sys
from collections.abc import Callable
from functools import lru_cache
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
SCHEMAS = ROOT / "schemas"
OUTPUT = ROOT / "src" / "xlsxedit" / "_ooxml_order.py"
NS_OUTPUT = ROOT / "src" / "xlsxedit" / "_ooxml_ns.py"
NOTES_OUTPUT = ROOT / "docs" / "ooxml-schema-notes.md"

_XSD_NS = "http://www.w3.org/2001/XMLSchema"
_SML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_X14_SCHEMA = "microsoft/ms-xlsx/spreadsheetml-2009-9-main.xsd"

# Folders under schemas/ whose files define namespaces, in the order that
# decides which file is named for a namespace several files declare.
SCHEMA_FOLDERS = ("transitional", "microsoft/ms-xlsx", "microsoft")

# Constant name and prefix of every namespace a vendored schema declares or
# imports. The prefix is the one Excel declares for the namespace.
SCHEMA_NAMESPACES: dict[str, tuple[str, str]] = {
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main": ("SML", "x"),
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships": ("OFFICE_REL", "r"),
    "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes": ("SHARED_TYPES", "s"),
    "http://schemas.openxmlformats.org/drawingml/2006/main": ("A", "a"),
    "http://schemas.openxmlformats.org/drawingml/2006/chart": ("CHART", "c"),
    "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing": ("CDR", "cdr"),
    "http://schemas.openxmlformats.org/drawingml/2006/diagram": ("DGM", "dgm"),
    "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas": ("LC", "lc"),
    "http://schemas.openxmlformats.org/drawingml/2006/picture": ("PIC", "pic"),
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing": ("XDR", "xdr"),
    "http://schemas.microsoft.com/office/excel/2006/main": ("XM", "xm"),
    "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main": ("X14", "x14"),
    "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac": ("X14AC", "x14ac"),
    "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main": ("X15", "x15"),
    "http://schemas.microsoft.com/office/spreadsheetml/2014/11/main": ("X16", "x16"),
    "http://schemas.microsoft.com/office/spreadsheetml/2014/revision": ("XR", "xr"),
    "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2": ("XR2", "xr2"),
    "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6": ("XR6", "xr6"),
    "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10": ("XR10", "xr10"),
}

# Namespaces that Excel writes and no vendored schema declares: constant
# name, prefix (None where Excel only declares the namespace as the default),
# and the specification that defines the namespace.
OTHER_NAMESPACES: dict[str, tuple[str, str | None, str]] = {
    "http://schemas.openxmlformats.org/package/2006/content-types": (
        "CT", None, "ECMA-376 Part 2 (OPC)"
    ),
    "http://schemas.openxmlformats.org/package/2006/relationships": (
        "REL", None, "ECMA-376 Part 2 (OPC)"
    ),
    "http://schemas.openxmlformats.org/package/2006/metadata/core-properties": (
        "CP", "cp", "ECMA-376 Part 2 (OPC)"
    ),
    "http://purl.org/dc/elements/1.1/": ("DC", "dc", "Dublin Core, for ECMA-376 Part 2"),
    "http://purl.org/dc/terms/": ("DCTERMS", "dcterms", "Dublin Core, for ECMA-376 Part 2"),
    "http://purl.org/dc/dcmitype/": ("DCMITYPE", "dcmitype", "Dublin Core, for ECMA-376 Part 2"),
    "http://www.w3.org/2001/XMLSchema-instance": ("XSI", "xsi", "W3C XML Schema"),
    "http://www.w3.org/XML/1998/namespace": ("XML", "xml", "W3C Namespaces in XML"),
    "http://schemas.openxmlformats.org/markup-compatibility/2006": (
        "MC", "mc", "ECMA-376 Part 3"
    ),
    "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties": (
        "EXTENDED_PROPERTIES", None, "ECMA-376 Part 1"
    ),
    "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes": (
        "VT", "vt", "ECMA-376 Part 1"
    ),
    "http://schemas.openxmlformats.org/officeDocument/2006/customXml": (
        "DS", "ds", "ECMA-376 Part 1"
    ),
    "urn:schemas-microsoft-com:vml": ("V", "v", "ECMA-376 Part 4 (VML)"),
    "urn:schemas-microsoft-com:office:office": ("O", "o", "ECMA-376 Part 4 (VML)"),
    "urn:schemas-microsoft-com:office:excel": ("VML_EXCEL", "x", "ECMA-376 Part 4 (VML)"),
    "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac": (
        "X15AC", "x15ac", "[MS-XLSX] 5.9"
    ),
    "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main": (
        "X16R2", "x16r2", "[MS-XLSX] 5.11"
    ),
    "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3": (
        "XR3", "xr3", "[MS-XLSX] 5.17"
    ),
    "http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout": (
        "XPDL", "xpdl", "[MS-XLSX] 5.19"
    ),
    "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2": (
        "XLRD2", "xlrd2", "[MS-XLSX] 5.20"
    ),
    "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures": (
        "XCALCF", "xcalcf", "[MS-XLSX] 5.22"
    ),
    "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments": (
        "THREADED_COMMENTS", None, "[MS-XLSX] 5.23"
    ),
    "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray": (
        "XDA", "xda", "[MS-XLSX] 5.25"
    ),
    "http://schemas.microsoft.com/office/spreadsheetml/2024/workbookCompatibilityVersion": (
        "XLWCV", "xlwcv", "[MS-XLSX] 5.42"
    ),
    "http://schemas.microsoft.com/office/drawing/2010/slicer": ("SLE", "sle", "[MS-XLSX] 5.2"),
    "http://schemas.microsoft.com/office/drawing/2010/main": ("A14", "a14", "[MS-ODRAWXML] 5.1"),
    "http://schemas.microsoft.com/office/drawing/2007/8/2/chart": (
        "C14", "c14", "[MS-ODRAWXML] 5.10"
    ),
    "http://schemas.microsoft.com/office/drawing/2012/chart": ("C15", "c15", "[MS-ODRAWXML] 5.13"),
    "http://schemas.microsoft.com/office/drawing/2012/chartStyle": (
        "CS", "cs", "[MS-ODRAWXML] 5.15"
    ),
    "http://schemas.microsoft.com/office/thememl/2012/main": (
        "THM15", "thm15", "[MS-ODRAWXML] 5.17"
    ),
    "http://schemas.microsoft.com/office/drawing/2014/chart": ("C16", "c16", "[MS-ODRAWXML] 5.20"),
    "http://schemas.microsoft.com/office/drawing/2014/chartex": (
        "CHART_EX", "cx", "[MS-ODRAWXML] 5.22"
    ),
    "http://schemas.microsoft.com/office/drawing/2014/main": ("A16", "a16", "[MS-ODRAWXML] 5.23"),
    "http://schemas.microsoft.com/office/drawing/2017/03/chart": (
        "C16R3", "c16r3", "[MS-ODRAWXML] 5.31"
    ),
    "http://schemas.microsoft.com/office/drawing/2015/06/chart": (
        "C16R2", "c16r2", "[MS-ODRAWXML] 5.42"
    ),
}

# (schema file under schemas/, type name) for every table the module exports.
ORDER_TYPES: tuple[tuple[str, str], ...] = (
    ("transitional/sml.xsd", "CT_Worksheet"),
    ("transitional/sml.xsd", "CT_Stylesheet"),
    ("transitional/sml.xsd", "CT_Xf"),
    ("transitional/sml.xsd", "CT_Workbook"),
    ("transitional/sml.xsd", "CT_Table"),
    ("transitional/sml.xsd", "CT_TableColumn"),
    ("transitional/sml.xsd", "CT_Row"),
    ("transitional/sml.xsd", "CT_Cell"),
    ("transitional/sml.xsd", "CT_CfRule"),
    ("transitional/sml.xsd", "CT_ConditionalFormatting"),
    ("transitional/sml.xsd", "CT_DataValidation"),
    ("transitional/dml-spreadsheetDrawing.xsd", "CT_TwoCellAnchor"),
    ("transitional/dml-chart.xsd", "CT_Chart"),
    ("transitional/dml-chart.xsd", "CT_ChartSpace"),
)
EXTENSION_ORDER_TYPES: tuple[tuple[str, str], ...] = (
    (_X14_SCHEMA, "CT_ConditionalFormattings"),
    (_X14_SCHEMA, "CT_ConditionalFormatting"),
    (_X14_SCHEMA, "CT_CfRule"),
    (_X14_SCHEMA, "CT_ColorScale"),
    (_X14_SCHEMA, "CT_DataBar"),
    (_X14_SCHEMA, "CT_IconSet"),
    (_X14_SCHEMA, "CT_Cfvo"),
    (_X14_SCHEMA, "CT_DataValidations"),
    (_X14_SCHEMA, "CT_DataValidation"),
    (_X14_SCHEMA, "CT_DataValidationFormula"),
    (_X14_SCHEMA, "CT_SparklineGroups"),
    (_X14_SCHEMA, "CT_SparklineGroup"),
    (_X14_SCHEMA, "CT_Sparklines"),
    (_X14_SCHEMA, "CT_Sparkline"),
)
ATTRIBUTE_ORDER_TYPES: tuple[tuple[str, str], ...] = (
    ("transitional/sml.xsd", "CT_DataValidation"),
)
ATTRIBUTE_DEFAULT_TYPES: tuple[tuple[str, str], ...] = (
    ("transitional/sml.xsd", "CT_Table"),
    ("transitional/sml.xsd", "CT_TableColumn"),
    ("transitional/sml.xsd", "CT_DataValidation"),
)
ENUM_TYPES: tuple[tuple[str, str], ...] = (
    ("transitional/sml.xsd", "ST_TableType"),
    ("transitional/sml.xsd", "ST_TotalsRowFunction"),
    ("transitional/sml.xsd", "ST_DataValidationType"),
    ("transitional/sml.xsd", "ST_DataValidationOperator"),
    ("transitional/sml.xsd", "ST_DataValidationErrorStyle"),
)

_HEADER = '''\
"""Child-element and attribute order, attribute defaults, and enumerations from the OOXML schemas.

Generated by ``scripts/gen_ooxml_order.py`` from the ECMA-376 transitional
schemas in ``schemas/transitional`` and Microsoft's [MS-XLSX] schemas in
``schemas/microsoft/ms-xlsx``. Do not edit; change the generator and rerun it.

``ORDER`` maps an ECMA-376 complex type to the local names of its child
elements in schema sequence. The alternatives of an ``xsd:choice`` keep their
schema order; at most one of them occurs. ``ATTRIBUTE_ORDER`` maps a complex
type to the names of its attributes in the order the schema declares them.

``EXTENSION_ORDER`` does the same as ``ORDER`` for complex types of Microsoft's
extension schemas, with the type and every child named in Clark notation
(``{namespace}local``), because one type's children come from several
namespaces: an x14 ``cfRule`` starts with ``xm:f`` elements.
"""

from __future__ import annotations
'''


def _xsd(localname: str) -> str:
    return f"{{{_XSD_NS}}}{localname}"


@lru_cache(maxsize=None)
def _schema(filename: str) -> etree._Element:
    return etree.parse(str(SCHEMAS / filename)).getroot()


def _definition(filename: str, kind: str, name: str) -> etree._Element:
    for node in _schema(filename).iterchildren(_xsd(kind)):
        if node.get("name") == name:
            return node
    raise LookupError(f"{filename} defines no xsd:{kind} named {name!r}")


def _local(qualified: str) -> str:
    return qualified.rpartition(":")[2]


def _clark(namespace: str | None, localname: str) -> str:
    return f"{{{namespace}}}{localname}" if namespace else localname


def _resolve(node: etree._Element, qualified: str) -> str:
    prefix, _, localname = qualified.rpartition(":")
    return _clark(node.nsmap.get(prefix or None), localname)


def _target_namespace(filename: str) -> str | None:
    return _schema(filename).get("targetNamespace")


def _local_element(filename: str, node: etree._Element) -> str:
    form = node.get("form") or _schema(filename).get("elementFormDefault", "unqualified")
    namespace = _target_namespace(filename) if form == "qualified" else None
    return _clark(namespace, node.get("name"))


def _element_names(filename: str, particle: etree._Element) -> list[str]:
    """Clark names of the child elements a content model allows, in schema order."""
    names: list[str] = []
    for node in particle.iterchildren(etree.Element):
        kind = etree.QName(node).localname
        if kind in ("annotation", "attribute", "attributeGroup", "anyAttribute"):
            continue
        if kind == "element":
            ref = node.get("ref")
            names.append(_resolve(node, ref) if ref else _local_element(filename, node))
            continue
        if kind not in ("sequence", "choice", "group"):
            raise ValueError(f"unsupported xsd:{kind} in a content model")
        if node.get("maxOccurs", "1") != "1":
            raise ValueError(f"a repeating xsd:{kind} has no single child order")
        if kind == "group":
            group = etree.QName(_resolve(node, node.get("ref")))
            if group.namespace != _target_namespace(filename):
                raise ValueError(f"group {group} is not defined in {filename}")
            node = _definition(filename, "group", group.localname)
        names.extend(_element_names(filename, node))
    return names


def _qualified_child_order(filename: str, type_name: str) -> tuple[str, ...]:
    names = _element_names(filename, _definition(filename, "complexType", type_name))
    if len(set(names)) != len(names):
        raise ValueError(f"{type_name} names a child element twice")
    return tuple(names)


def child_order(filename: str, type_name: str) -> tuple[str, ...]:
    """Local names of an ECMA-376 type's children, which share its namespace."""
    namespace = _target_namespace(filename)
    names = tuple(etree.QName(name) for name in _qualified_child_order(filename, type_name))
    if any(name.namespace != namespace for name in names):
        raise ValueError(f"{type_name} has a child outside {namespace}; list it as an extension")
    return tuple(name.localname for name in names)


def extension_child_order(filename: str, type_name: str) -> tuple[str, tuple[str, ...]]:
    """The Clark name of a type and of each of its children."""
    key = _clark(_target_namespace(filename), type_name)
    return key, _qualified_child_order(filename, type_name)


def attribute_order(filename: str, type_name: str) -> tuple[str, ...]:
    complex_type = _definition(filename, "complexType", type_name)
    for indirect in ("attributeGroup", "complexContent", "simpleContent"):
        if complex_type.find(_xsd(indirect)) is not None:
            raise ValueError(f"{type_name} declares attributes through an xsd:{indirect}")
    return tuple(
        attribute.get("name") or _local(attribute.get("ref"))
        for attribute in complex_type.iterchildren(_xsd("attribute"))
    )


def attribute_defaults(filename: str, type_name: str) -> dict[str, str]:
    complex_type = _definition(filename, "complexType", type_name)
    return {
        attribute.get("name") or _local(attribute.get("ref")): attribute.get("default")
        for attribute in complex_type.iterchildren(_xsd("attribute"))
        if attribute.get("default") is not None
    }


def enumeration(filename: str, type_name: str) -> tuple[str, ...]:
    simple_type = _definition(filename, "simpleType", type_name)
    return tuple(node.get("value") for node in simple_type.iter(_xsd("enumeration")))


def elements_of_type(filename: str, type_name: str) -> tuple[str, ...]:
    """Names of the elements ``filename`` declares with the type ``type_name``."""
    namespace = _target_namespace(filename)
    names = {
        node.get("name")
        for node in _schema(filename).iter(_xsd("element"))
        if node.get("type") and _resolve(node, node.get("type")) == _clark(namespace, type_name)
    }
    return tuple(sorted(names))


def _is_alias(root: etree._Element) -> bool:
    """Whether a schema document only includes another, as the ms-xlsx aliases do."""
    kinds = {etree.QName(node).localname for node in root.iterchildren(etree.Element)}
    return kinds <= {"include", "annotation"}


def schema_namespaces() -> dict[str, str]:
    """Each namespace a vendored schema declares or imports, with the file that defines it."""
    defined: dict[str, str] = {}
    imported: set[str] = set()
    for folder in SCHEMA_FOLDERS:
        for path in sorted((SCHEMAS / folder).glob("*.xsd")):
            filename = path.relative_to(SCHEMAS).as_posix()
            root = _schema(filename)
            imported.update(node.get("namespace") for node in root.iterchildren(_xsd("import")))
            namespace = root.get("targetNamespace")
            if namespace is not None and not _is_alias(root):
                defined.setdefault(namespace, filename)
    missing = imported - defined.keys() - {None}
    if missing:
        raise ValueError(f"imported but not vendored: {', '.join(sorted(missing))}")
    return defined


def namespace_rows() -> list[tuple[str, str, str | None, str]]:
    """``(name, namespace, prefix, source)`` of every namespace, vendored ones first.

    The source of a vendored namespace is its schema file under ``schemas/``.
    """
    vendored = schema_namespaces()
    unnamed = vendored.keys() - SCHEMA_NAMESPACES.keys()
    if unnamed:
        raise ValueError(f"add to SCHEMA_NAMESPACES: {', '.join(sorted(unnamed))}")
    unvendored = SCHEMA_NAMESPACES.keys() - vendored.keys()
    if unvendored:
        raise ValueError(f"no vendored schema declares {', '.join(sorted(unvendored))}")
    both = vendored.keys() & OTHER_NAMESPACES.keys()
    if both:
        raise ValueError(f"remove from OTHER_NAMESPACES: {', '.join(sorted(both))}")
    rows = [(name, ns, prefix, vendored[ns]) for ns, (name, prefix) in SCHEMA_NAMESPACES.items()]
    rows += [(name, ns, prefix, spec) for ns, (name, prefix, spec) in OTHER_NAMESPACES.items()]
    names = [row[0] for row in rows]
    if len(set(names)) != len(names):
        raise ValueError("two namespaces share a constant name")
    return rows


def _tuple_block(key: str, values: tuple[str, ...]) -> list[str]:
    return [f"    {json.dumps(key)}: (", *(f"        {json.dumps(v)}," for v in values), "    ),"]


def render() -> str:
    """Return the source of the generated module."""
    lines = [_HEADER, "ORDER: dict[str, tuple[str, ...]] = {"]
    for filename, type_name in ORDER_TYPES:
        lines += _tuple_block(type_name, child_order(filename, type_name))
    lines += ["}", "", "EXTENSION_ORDER: dict[str, tuple[str, ...]] = {"]
    for filename, type_name in EXTENSION_ORDER_TYPES:
        lines += _tuple_block(*extension_child_order(filename, type_name))
    lines += ["}", "", "ATTRIBUTE_ORDER: dict[str, tuple[str, ...]] = {"]
    for filename, type_name in ATTRIBUTE_ORDER_TYPES:
        lines += _tuple_block(type_name, attribute_order(filename, type_name))
    lines += ["}", "", "ATTRIBUTE_DEFAULTS: dict[str, dict[str, str]] = {"]
    for filename, type_name in ATTRIBUTE_DEFAULT_TYPES:
        lines.append(f"    {json.dumps(type_name)}: {{")
        for name, default in attribute_defaults(filename, type_name).items():
            lines.append(f"        {json.dumps(name)}: {json.dumps(default)},")
        lines.append("    },")
    lines += ["}", "", "ENUMS: dict[str, tuple[str, ...]] = {"]
    for filename, type_name in ENUM_TYPES:
        lines += _tuple_block(type_name, enumeration(filename, type_name))
    lines.append("}")
    return "\n".join(lines) + "\n"


_NS_HEADER = '''\
"""Namespace URIs of the OOXML vocabularies and the prefixes Excel declares for them.

Generated by ``scripts/gen_ooxml_order.py``. Do not edit; change the generator
and rerun it.

The first constants are every namespace that a schema in ``schemas/transitional``
or ``schemas/microsoft/ms-xlsx`` declares or imports, read from the schemas.
The rest are namespaces Excel writes that no vendored schema declares, such as
the package and markup-compatibility namespaces and Microsoft extensions without
a vendored schema; the generator lists them with the specification of each.
``docs/ooxml-schema-notes.md`` tabulates both.

``PREFIXES`` maps each namespace to the prefix Excel declares for it, or to
``None`` for a namespace Excel only declares as the default namespace. Two
namespaces can share a prefix in different parts: Excel writes ``x`` for
SpreadsheetML and, in VML drawings, for ``urn:schemas-microsoft-com:office:excel``.
``SCHEMA_FILES`` maps each vendored namespace to the schema file under
``schemas/`` that defines it.
"""

from __future__ import annotations
'''


def render_namespaces() -> str:
    """Return the source of the generated namespace module."""
    rows = namespace_rows()
    lines = [_NS_HEADER, "# Declared or imported by a vendored schema"]
    lines += [f"{name}_NS = {json.dumps(ns)}" for name, ns, _, _ in rows if ns in SCHEMA_NAMESPACES]
    lines += ["", "# Written by Excel, without a vendored schema"]
    lines += [f"{name}_NS = {json.dumps(ns)}" for name, ns, _, _ in rows if ns in OTHER_NAMESPACES]
    lines += ["", "PREFIXES: dict[str, str | None] = {"]
    for _, ns, prefix, _ in rows:
        lines.append(f"    {json.dumps(ns)}: {'None' if prefix is None else json.dumps(prefix)},")
    lines += ["}", "", "SCHEMA_FILES: dict[str, str] = {"]
    for _, ns, _, source in rows:
        if ns in SCHEMA_NAMESPACES:
            lines.append(f"    {json.dumps(ns)}: {json.dumps(source)},")
    lines.append("}")
    return "\n".join(lines) + "\n"


def _import_from_repository() -> None:
    """Make ``tests`` and the in-tree ``xlsxedit`` importable for the notes."""
    for path in (ROOT / "src", ROOT):
        if str(path) not in sys.path:
            sys.path.insert(0, str(path))


def _qualified(clark: str) -> str:
    """``prefix:local`` for a Clark name, with the prefix Excel declares; SpreadsheetML unprefixed."""
    qname = etree.QName(clark)
    if qname.namespace == _SML_NS:
        return qname.localname
    prefixes = {ns: prefix for _, ns, prefix, _ in namespace_rows()}
    if prefixes.get(qname.namespace) is None:
        raise ValueError(f"no prefix for the namespace of {clark}")
    return f"{prefixes[qname.namespace]}:{qname.localname}"


def _code_list(names: tuple[str, ...]) -> str:
    return ", ".join(f"`{name}`" for name in names)


_NOTES_INTRO = """\
# OOXML schema notes

<!-- Generated by scripts/gen_ooxml_order.py; do not edit by hand. -->

What the vendored schemas and Excel's own files say about the XML that
xlsxedit reads and writes: namespaces and their prefixes, the extension URIs
Excel uses, child-element order, attribute defaults and enumerations, and the
places where Excel departs from the schemas. Read it before writing or
checking OOXML by hand. [excel-xlsx-structure.md](excel-xlsx-structure.md)
describes the package in prose.

`python scripts/gen_ooxml_order.py` writes this page from the schemas in
[`schemas/`](../schemas/README.md), the extension registry in
`src/xlsxedit/_ooxml_ext.py` and the chart allow-list in
`tests/schema_validate.py`. `tests/test_ooxml_ns.py` fails when the page no
longer matches them.

## Transitional, not strict

ECMA-376 defines each vocabulary twice. Excel writes the transitional form,
whose namespaces start with `http://schemas.openxmlformats.org/` (SpreadsheetML
is `http://schemas.openxmlformats.org/spreadsheetml/2006/main`). The strict
form uses namespaces under `http://purl.oclc.org/ooxml/` instead (SpreadsheetML
is `http://purl.oclc.org/ooxml/spreadsheetml/main`), and Excel writes it only
for a workbook saved as *Strict Open XML Spreadsheet*. The vendored schemas,
the constants in `src/xlsxedit/_ooxml_ns.py` and everything on this page are
transitional; nothing in xlsxedit names a strict namespace.

## Namespaces

`src/xlsxedit/_ooxml_ns.py` has one constant for each namespace below, and
`PREFIXES` maps each namespace to the prefix Excel declares for it. *default*
marks a namespace Excel only declares as the default namespace. A namespace
defined by a file under `schemas/` has a vendored schema; the others name the
specification that defines them.

| Constant | Prefix | Namespace | Defined by |
| --- | --- | --- | --- |
"""

_NOTES_EXTENSIONS = """\
## Extension lists

Excel stores what ECMA-376 lacks in extension lists. An `extLst` holds `ext`
elements; each `ext` names its content with a GUID in its `uri` attribute and
holds one payload element in a Microsoft namespace, such as
`x14:conditionalFormattings`. Excel skips an `ext` whose URI it does not
know, so such content survives only while every tool that saves the file
keeps it.

- Compare URIs without regard to case. Excel writes most GUIDs in mixed case
  (`{78C0D931-6437-407d-A8EE-F0AAD7539E65}`), and the specifications print
  some of the same GUIDs in capitals. `xlsxedit._ooxml_ext.extension(uri)`
  looks a URI up that way.
- The `xm:f`, `xm:ref` and `xm:sqref` elements of a payload hold cell
  references. An application can adjust them when rows or columns move, even
  inside an `ext` it does not know ([MS-XLSX] 2.2.1).
- The base schemas accept any content in an `ext`. Validate a payload against
  `schemas/microsoft/extensions.xsd`, as `tests/schema_validate.py` does for
  every payload in a vendored namespace.

`src/xlsxedit/_ooxml_ext.py` records each URI below. *Host* is the element
whose `extLst` holds the `ext`. *In fixtures* says whether a workbook in
`fixtures/` holds it; `tests/test_ooxml_ext.py` checks that both ways, so a
new fixture that holds an unrecorded URI fails the suite.

| URI | Payload | Host | Feature | Specification | In fixtures |
| --- | --- | --- | --- | --- | --- |
"""

_NOTES_MARKUP_COMPATIBILITY = """\
## Markup compatibility

Excel marks extension content outside `extLst` with markup compatibility
(ECMA-376 Part 3). A part's root lists namespaces in `mc:Ignorable`, such as
`mc:Ignorable="x14ac xr xr2 xr3"`, and attributes in those namespaces, such as
`x14ac:dyDescent` and `xr:uid`, sit on base elements. `mc:AlternateContent`
offers one or more `mc:Choice` elements, each requiring namespaces through its
`Requires` prefixes, and an `mc:Fallback`.

To see a part the way a consumer that understands a given set of namespaces
does, for example to validate it against the base schemas:

1. Replace each `mc:AlternateContent` with the content of its first
   `mc:Choice` whose `Requires` prefixes all name understood namespaces, else
   with the content of its `mc:Fallback`, else remove it.
2. Remove every `mc:*` and `xml:*` attribute.
3. Remove each element and attribute in a namespace that an enclosing
   `mc:Ignorable` lists, unless the namespace is understood. To validate,
   also remove such an attribute from an element whose own namespace is not
   understood: the base schemas have no place for `x14ac:dyDescent` on a
   `row` even when x14ac is understood.

Step 3 has one exception: keep each extension payload, the child of an
`extLst/ext`, and every element inside a payload whose namespace is not
understood. A payload's namespace is often ignorable as well, but the base
schemas require each `ext` to keep its payload; validate the payload on its
own instead. `apply_markup_compatibility` in `tests/schema_validate.py`
implements these steps.

When writing a part, keep every namespace declaration whose prefix an
`mc:Ignorable` or `Requires` attribute names, even when no element or
attribute uses it. Those attributes name prefixes, not namespaces, so
dropping such a declaration leaves a prefix that resolves to nothing.
"""

_NOTES_ORDER = """\
## Child-element order

The schemas fix the order of an element's children; Excel reports a part with
children out of order as corrupt. Each list gives the children in schema
order, from `ORDER` in `src/xlsxedit/_ooxml_order.py`. The alternatives of an
`xsd:choice` keep their schema order, and at most one of them occurs.
"""

_NOTES_EXTENSION_ORDER = """\
### Extension types

From `EXTENSION_ORDER`. Children of one extension type come from several
namespaces: an `x14:cfRule` starts with `xm:f` elements.
"""

_NOTES_DEFAULTS = """\
## Attribute defaults and enumerations

An attribute left out takes its schema default. From `ATTRIBUTE_DEFAULTS` and
`ENUMS` in `src/xlsxedit/_ooxml_order.py`.
"""

_NOTES_DEVIATIONS = """\
## Where Excel departs from the schemas

- Excel writes `xml:space="preserve"` on a string's `t` element when its text
  starts or ends with whitespace. `t` has the simple type `ST_Xstring`, which
  allows no attributes, so a validator must drop `xml:*` attributes first
  (step 2 above). Keep writing it on such text, as Excel does.
- Excel writes the x14 `workbookPr` extension URI as
  `{79F54976-1DA5-4618-B147-4CDE4B953A38}`, which [MS-XLSX] 2.2.4.11 prints as
  `{79F54976-1DA5-4618-B147-ACDE4B953A38}`. The registry records Excel's
  spelling.

Chart parts that Excel writes are not schema-valid. `CHART_ALLOWLIST` in
`tests/schema_validate.py` names each known violation and puts it into schema
order before validation:
"""

def _order_section(
    level: str, filename: str, type_name: str, children: tuple[str, ...]
) -> list[str]:
    """A heading naming the type and its elements, then its children in order.

    Extension types carry their prefix, so ``x14:CT_CfRule`` and ``CT_CfRule``
    get headings of their own.
    """
    namespace = _target_namespace(filename)
    elements = tuple(_qualified(_clark(namespace, e)) for e in elements_of_type(filename, type_name))
    name = _qualified(_clark(namespace, type_name))
    heading = f"{level} `{name}`" + (f" ({_code_list(elements)})" if elements else "")
    return [heading, "", *(f"1. `{child}`" for child in children), ""]


def render_notes() -> str:
    """Return the text of docs/ooxml-schema-notes.md."""
    _import_from_repository()
    from tests.schema_validate import CHART_ALLOWLIST
    from xlsxedit._ooxml_ext import EXTENSIONS

    lines = [_NOTES_INTRO.rstrip("\n")]
    for name, ns, prefix, source in namespace_rows():
        shown = "*default*" if prefix is None else f"`{prefix}`"
        defined = f"[`{source}`](../schemas/{source})" if ns in SCHEMA_NAMESPACES else source
        lines.append(f"| `{name}_NS` | {shown} | `{ns}` | {defined} |")
    lines += ["", _NOTES_EXTENSIONS.rstrip("\n")]
    for uri, ext in EXTENSIONS.items():
        payload = _qualified(f"{{{ext.namespace}}}{ext.element}")
        lines.append(
            f"| `{uri}` | `{payload}` | `{_qualified(ext.host)}` | {ext.feature} "
            f"| {ext.spec or 'not listed'} | {'yes' if ext.verified else 'no'} |"
        )
    lines += ["", _NOTES_MARKUP_COMPATIBILITY.rstrip("\n"), "", _NOTES_ORDER]
    for filename, type_name in ORDER_TYPES:
        lines += _order_section("###", filename, type_name, child_order(filename, type_name))
    lines.append(_NOTES_EXTENSION_ORDER)
    for filename, type_name in EXTENSION_ORDER_TYPES:
        _, children = extension_child_order(filename, type_name)
        qualified = tuple(_qualified(child) for child in children)
        lines += _order_section("####", filename, type_name, qualified)
    lines.append(_NOTES_DEFAULTS)
    for filename, type_name in ATTRIBUTE_DEFAULT_TYPES:
        lines += [f"### `{type_name}` defaults", "", "| Attribute | Default |", "| --- | --- |"]
        defaults = attribute_defaults(filename, type_name)
        lines += [f"| `{name}` | `{value}` |" for name, value in defaults.items()]
        lines.append("")
    for filename, type_name in ENUM_TYPES:
        lines += [f"### `{type_name}`", "", _code_list(enumeration(filename, type_name)), ""]
    lines.append(_NOTES_DEVIATIONS)
    lines += [f"- {reason}." for reason, _ in CHART_ALLOWLIST]
    return "\n".join(lines) + "\n"


# Each generated file with the function that renders it, in writing order:
# the notes import the namespace module through the extension registry.
OUTPUTS: tuple[tuple[Path, Callable[[], str]], ...] = (
    (OUTPUT, render),
    (NS_OUTPUT, render_namespaces),
    (NOTES_OUTPUT, render_notes),
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--check",
        action="store_true",
        help="report stale files with exit status 1 instead of rewriting them",
    )
    args = parser.parse_args(argv)
    stale = 0
    for path, render_file in OUTPUTS:
        source = render_file()
        current = path.read_text(encoding="utf-8") if path.is_file() else None
        if current == source:
            continue
        if args.check:
            print(f"{path.relative_to(ROOT)} is stale; run scripts/gen_ooxml_order.py", file=sys.stderr)
            stale += 1
            continue
        path.write_text(source, encoding="utf-8", newline="\n")
        print(f"wrote {path.relative_to(ROOT)}")
    return 1 if stale else 0


if __name__ == "__main__":
    sys.exit(main())

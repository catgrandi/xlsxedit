"""Validate saved packages against the vendored ECMA-376 transitional XSDs.

``assert_valid_package`` applies Markup Compatibility preprocessing
(ECMA-376 Part 3) to every part it has a schema for, then validates it
against ``schemas/transitional/sml.xsd`` and the DrawingML schemas that file
imports:

- SpreadsheetML, spreadsheet drawing, and theme parts must be valid.
- Chart parts only warn with ``SchemaWarning``, because charts that Excel
  writes are not schema-valid. ``CHART_ALLOWLIST`` names the known Excel
  violations, which stay silent.

Parts without a schema in the vendored subset (relationships, content types,
document properties, VML, Microsoft extension parts) are skipped.
"""

from __future__ import annotations

import io
import re
import warnings
import zipfile
from copy import deepcopy
from functools import lru_cache
from pathlib import Path, PurePosixPath
from typing import IO, TYPE_CHECKING

from lxml import etree

if TYPE_CHECKING:
    from os import PathLike

    from xlsxedit import Workbook

SCHEMA = Path(__file__).resolve().parents[1] / "schemas" / "transitional" / "sml.xsd"

MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XML_NS = "http://www.w3.org/XML/1998/namespace"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

_SPREADSHEETML_CT = re.compile(
    r"application/vnd\.openxmlformats-officedocument\.spreadsheetml\.[\w.]+\+xml"
)
_MACRO_ENABLED_MAIN_CT = frozenset(
    {
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
        "application/vnd.ms-excel.template.macroEnabled.main+xml",
        "application/vnd.ms-excel.addin.macroEnabled.main+xml",
    }
)
_DRAWINGML_KIND = {
    "application/vnd.openxmlformats-officedocument.drawing+xml": "drawing",
    "application/vnd.openxmlformats-officedocument.theme+xml": "theme",
    "application/vnd.openxmlformats-officedocument.themeOverride+xml": "theme",
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml": "chart",
    "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml": "chart",
}

# Schema errors that Excel's own chart parts produce, with where they come
# from. A chart error matching none of these patterns is warned about.
CHART_ALLOWLIST: tuple[tuple[re.Pattern[str], str], ...] = (
    (
        re.compile(
            r"\{http://schemas\.openxmlformats\.org/drawingml/2006/chart\}"
            r"showDLblsOverMax': This element is not expected"
        ),
        "Excel writes c:extLst before c:showDLblsOverMax in CT_Chart "
        "(fixtures/ChartsAndTables.xlsx, xl/charts/chart1.xml and chart2.xml); "
        "src/xlsxedit/templates/default-bar-chart.xml copies that chart",
    ),
)

_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True)


class SchemaWarning(UserWarning):
    """A chart part fails schema validation in a way the allow-list does not explain."""


@lru_cache(maxsize=1)
def _schema() -> etree.XMLSchema:
    return etree.XMLSchema(etree.parse(str(SCHEMA)))


def apply_markup_compatibility(root: etree._Element) -> etree._Element:
    """Reduce ``root`` in place to what a consumer of the base schema sees.

    A validator understands no extension namespace, so:

    - ``mc:AlternateContent`` is replaced by the content of its
      ``mc:Fallback``, or removed when there is none.
    - ``mc:*`` and ``xml:*`` attributes, and attributes in a namespace that an
      enclosing ``mc:Ignorable`` lists, are removed.
    - Elements in an ignorable namespace are removed, except inside an
      extension (``extLst/ext``). The schema accepts any content there, and
      emptying an ``ext`` would fail it for a missing child.
    """
    _resolve_alternate_content(root)
    _strip_ignorable(root, frozenset(), inside_ext=False)
    return root


def _resolve_alternate_content(root: etree._Element) -> None:
    tag = f"{{{MC_NS}}}AlternateContent"
    while (alternate := next(root.iter(tag), None)) is not None:
        parent = alternate.getparent()
        index = parent.index(alternate)
        fallback = alternate.find(f"{{{MC_NS}}}Fallback")
        parent.remove(alternate)
        if fallback is not None:
            parent[index:index] = list(fallback)


def _declared_ignorable(elm: etree._Element) -> frozenset[str]:
    prefixes = (elm.get(f"{{{MC_NS}}}Ignorable") or "").split()
    return frozenset(elm.nsmap[p] for p in prefixes if p in elm.nsmap)


def _strip_ignorable(
    elm: etree._Element, inherited: frozenset[str], *, inside_ext: bool
) -> None:
    ignorable = inherited | _declared_ignorable(elm)
    for name in list(elm.attrib):
        namespace = etree.QName(name).namespace
        if namespace in (MC_NS, XML_NS) or namespace in ignorable:
            del elm.attrib[name]
    parent_is_ext_list = etree.QName(elm).localname == "extLst"
    for child in list(elm):
        if not isinstance(child.tag, str):
            continue
        qname = etree.QName(child)
        if qname.namespace in ignorable and not inside_ext:
            elm.remove(child)
            continue
        child_is_ext = parent_is_ext_list and qname.localname == "ext"
        _strip_ignorable(child, ignorable, inside_ext=inside_ext or child_is_ext)


def schema_errors(xml: bytes | etree._Element) -> list[str]:
    """Return the schema errors of one part after Markup Compatibility preprocessing.

    An element argument is copied, not modified.
    """
    if isinstance(xml, bytes):
        root = etree.fromstring(xml, _PARSER)
    else:
        root = deepcopy(xml)
    apply_markup_compatibility(root)
    schema = _schema()
    if schema.validate(root):
        return []
    return [f"line {error.line}: {error.message}" for error in schema.error_log]


def _is_known_chart_violation(error: str) -> bool:
    return any(pattern.search(error) for pattern, _reason in CHART_ALLOWLIST)


def _part_kind(content_type: str | None) -> str | None:
    if content_type is None:
        return None
    if content_type in _MACRO_ENABLED_MAIN_CT or _SPREADSHEETML_CT.fullmatch(content_type):
        return "spreadsheetml"
    return _DRAWINGML_KIND.get(content_type)


def _content_types(package: zipfile.ZipFile) -> tuple[dict[str, str], dict[str, str]]:
    root = etree.fromstring(package.read("[Content_Types].xml"), _PARSER)
    defaults = {
        elm.get("Extension").lower(): elm.get("ContentType")
        for elm in root.iter(f"{{{_CT_NS}}}Default")
    }
    overrides = {
        elm.get("PartName").lower(): elm.get("ContentType")
        for elm in root.iter(f"{{{_CT_NS}}}Override")
    }
    return defaults, overrides


def _as_zip_source(
    source: str | PathLike[str] | bytes | IO[bytes] | Workbook,
) -> str | PathLike[str] | IO[bytes]:
    if isinstance(source, bytes):
        return io.BytesIO(source)
    if hasattr(source, "save"):
        buffer = io.BytesIO()
        source.save(buffer)
        buffer.seek(0)
        return buffer
    return source


def assert_valid_package(
    source: str | PathLike[str] | bytes | IO[bytes] | Workbook,
) -> dict[str, str]:
    """Assert that every part of a package with a vendored schema is valid.

    ``source`` is a path, the package bytes, a binary file-like, or an object
    with a ``save(stream)`` method such as ``xlsxedit.Workbook``. Returns the
    validated part names mapped to their kind: ``"spreadsheetml"``,
    ``"drawing"``, ``"theme"``, or ``"chart"``.
    """
    checked: dict[str, str] = {}
    failures: list[str] = []
    with zipfile.ZipFile(_as_zip_source(source)) as package:
        defaults, overrides = _content_types(package)
        for name in package.namelist():
            content_type = overrides.get(f"/{name}".lower())
            if content_type is None:
                content_type = defaults.get(PurePosixPath(name).suffix[1:].lower())
            kind = _part_kind(content_type)
            if kind is None:
                continue
            checked[name] = kind
            errors = schema_errors(package.read(name))
            if kind != "chart":
                failures.extend(f"{name} {error}" for error in errors)
                continue
            for error in errors:
                if not _is_known_chart_violation(error):
                    warnings.warn(f"{name} {error}", SchemaWarning, stacklevel=2)
    if failures:
        raise AssertionError("schema-invalid package parts:\n" + "\n".join(failures))
    return checked

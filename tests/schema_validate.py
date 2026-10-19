"""Validate saved packages against the vendored ECMA-376 and Microsoft schemas.

``assert_valid_package`` applies Markup Compatibility preprocessing
(ECMA-376 Part 3) to every part it has a schema for, then validates it
against ``schemas/transitional/sml.xsd`` and the DrawingML schemas that file
imports:

- SpreadsheetML, spreadsheet drawing, and theme parts must be valid.
- Chart parts only warn with ``SchemaWarning``, because charts that Excel
  writes are not schema-valid. ``CHART_ALLOWLIST`` names each known Excel
  violation with a rewrite that puts it into schema order before
  validation, so a known violation neither warns nor hides a later error.

The base schemas accept any content inside an extension (``extLst/ext``).
Each extension payload in a namespace that ``schemas/microsoft/ms-xlsx``
vendors, such as x14 conditional formatting, data validation, and
sparklines, is validated as well, against ``schemas/microsoft/extensions.xsd``.
Payloads in other namespaces are skipped.

Parts without a schema in the vendored subset (relationships, content types,
document properties, VML, Microsoft extension parts) are skipped.
"""

from __future__ import annotations

import io
import re
import warnings
import zipfile
from collections.abc import Callable
from copy import deepcopy
from functools import lru_cache
from pathlib import Path, PurePosixPath
from typing import IO, TYPE_CHECKING

from lxml import etree

if TYPE_CHECKING:
    from os import PathLike

    from xlsxedit import Workbook

_SCHEMAS = Path(__file__).resolve().parents[1] / "schemas"
SCHEMA = _SCHEMAS / "transitional" / "sml.xsd"
EXTENSION_SCHEMA = _SCHEMAS / "microsoft" / "extensions.xsd"

MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XML_NS = "http://www.w3.org/XML/1998/namespace"
_XSD_NS = "http://www.w3.org/2001/XMLSchema"
_SML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"

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


def _show_dlbls_over_max_before_ext_lst(chart_space: etree._Element) -> None:
    chart = chart_space.find(f"{{{_C_NS}}}chart")
    if chart is None:
        return
    ext_lst = chart.find(f"{{{_C_NS}}}extLst")
    show = chart.find(f"{{{_C_NS}}}showDLblsOverMax")
    if ext_lst is not None and show is not None and chart.index(show) > chart.index(ext_lst):
        ext_lst.addprevious(show)


# Schema violations that Excel writes in its own chart parts, each with where
# it comes from and the rewrite that puts it into schema order. Rewriting
# instead of filtering error messages matters: libxml2 reports only the first
# content-model error in an element, so a filtered message would hide any
# later error in the same element.
CHART_ALLOWLIST: tuple[tuple[str, Callable[[etree._Element], None]], ...] = (
    (
        "Excel writes c:extLst before c:showDLblsOverMax in CT_Chart "
        "(fixtures/ChartsAndTables.xlsx, xl/charts/chart1.xml and chart2.xml); "
        "src/xlsxedit/templates/default-bar-chart.xml copies that chart",
        _show_dlbls_over_max_before_ext_lst,
    ),
)

_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True)


class SchemaWarning(UserWarning):
    """A chart part fails schema validation after the allow-list rewrites."""


@lru_cache(maxsize=1)
def _schema() -> etree.XMLSchema:
    return etree.XMLSchema(etree.parse(str(SCHEMA)))


@lru_cache(maxsize=1)
def _extension_schema() -> etree.XMLSchema:
    return etree.XMLSchema(etree.parse(str(EXTENSION_SCHEMA)))


@lru_cache(maxsize=1)
def extension_namespaces() -> frozenset[str]:
    """The Microsoft namespaces that ``EXTENSION_SCHEMA`` validates."""
    driver = etree.parse(str(EXTENSION_SCHEMA)).getroot()
    return frozenset(
        node.get("namespace") for node in driver.iterchildren(f"{{{_XSD_NS}}}import")
    ) - {_SML_NS}


def apply_markup_compatibility(
    root: etree._Element, understood: frozenset[str] = frozenset()
) -> etree._Element:
    """Reduce ``root`` in place to what a consumer of the given namespaces sees.

    ``understood`` names the extension namespaces the consumer validates.
    Content in a namespace that no ``mc:Ignorable`` lists is kept. So:

    - ``mc:AlternateContent`` is replaced by the content of its first
      ``mc:Choice`` whose ``Requires`` namespaces are all understood, else of
      its ``mc:Fallback``, or removed when there is neither.
    - ``mc:*`` and ``xml:*`` attributes are removed.
    - An element in a namespace that an enclosing ``mc:Ignorable`` lists is
      removed unless the namespace is understood, the element is an
      extension payload (a child of ``extLst/ext``), or it lies inside a
      payload whose namespace is not understood. The base schemas accept any
      payload, removing one would leave its ``ext`` without the child it
      requires, and libxml2 still validates any element inside it that the
      schemas declare.
    - An attribute in an ignorable namespace is removed unless both its
      namespace and its element's namespace are understood: an extension
      attribute on a base element has no place in the base type's schema.
    """
    _resolve_alternate_content(root, understood)
    _strip_ignorable(root, frozenset(), understood)
    return root


def _resolve_alternate_content(root: etree._Element, understood: frozenset[str]) -> None:
    tag = f"{{{MC_NS}}}AlternateContent"
    while (alternate := next(root.iter(tag), None)) is not None:
        parent = alternate.getparent()
        index = parent.index(alternate)
        chosen = alternate.find(f"{{{MC_NS}}}Fallback")
        for choice in alternate.iterchildren(f"{{{MC_NS}}}Choice"):
            required = (choice.get("Requires") or "").split()
            if required and all(choice.nsmap.get(p) in understood for p in required):
                chosen = choice
                break
        parent.remove(alternate)
        if chosen is not None:
            parent[index:index] = list(chosen)


def _declared_ignorable(elm: etree._Element) -> frozenset[str]:
    prefixes = (elm.get(f"{{{MC_NS}}}Ignorable") or "").split()
    return frozenset(elm.nsmap[p] for p in prefixes if p in elm.nsmap)


def _is_ext(elm: etree._Element) -> bool:
    """Whether ``elm`` is an ``ext`` in an ``extLst``, whose children are payloads."""
    if etree.QName(elm).localname != "ext":
        return False
    ext_list = elm.getparent()
    return ext_list is not None and etree.QName(ext_list).localname == "extLst"


def _strip_ignorable(
    elm: etree._Element,
    inherited: frozenset[str],
    understood: frozenset[str],
    *,
    keep_elements: bool = False,
) -> None:
    ignorable = inherited | _declared_ignorable(elm)
    element_understood = etree.QName(elm).namespace in understood
    for name in list(elm.attrib):
        namespace = etree.QName(name).namespace
        if namespace in (MC_NS, XML_NS):
            del elm.attrib[name]
        elif namespace in ignorable and not (element_understood and namespace in understood):
            del elm.attrib[name]
    holds_payloads = _is_ext(elm)
    for child in list(elm):
        if not isinstance(child.tag, str):
            continue
        namespace = etree.QName(child).namespace
        unknown = namespace not in understood
        if unknown and namespace in ignorable and not (keep_elements or holds_payloads):
            elm.remove(child)
            continue
        _strip_ignorable(
            child,
            ignorable,
            understood,
            keep_elements=keep_elements or (holds_payloads and unknown),
        )


def _payloads(root: etree._Element, namespaces: frozenset[str]) -> list[etree._Element]:
    """The outermost extension payloads in ``root`` in the given namespaces.

    A payload nested inside another is validated with the outer one, laxly:
    it is checked when the schemas declare its element.
    """
    found: list[etree._Element] = []
    for ext in root.iter("{*}ext"):
        if not _is_ext(ext):
            continue
        for payload in ext.iterchildren(etree.Element):
            if etree.QName(payload).namespace not in namespaces:
                continue
            if not any(ancestor in found for ancestor in payload.iterancestors()):
                found.append(payload)
    return found


def extension_payloads(xml: bytes | etree._Element) -> list[str]:
    """Return the Clark names of the payloads in one part that are validated.

    These are the outermost extension payloads whose namespace
    ``extension_namespaces()`` lists, after Markup Compatibility
    preprocessing. An element argument is copied, not modified.
    """
    root = etree.fromstring(xml, _PARSER) if isinstance(xml, bytes) else deepcopy(xml)
    namespaces = extension_namespaces()
    apply_markup_compatibility(root, namespaces)
    return [payload.tag for payload in _payloads(root, namespaces)]


def _extension_errors(root: etree._Element) -> list[str]:
    namespaces = extension_namespaces()
    if next(root.iter(*(f"{{{ns}}}*" for ns in namespaces)), None) is None:
        return []
    root = apply_markup_compatibility(deepcopy(root), namespaces)
    schema = _extension_schema()
    errors: list[str] = []
    for payload in _payloads(root, namespaces):
        if not schema.validate(payload):
            errors.extend(f"line {error.line}: {error.message}" for error in schema.error_log)
    return errors


def _validate(root: etree._Element) -> list[str]:
    extension_errors = _extension_errors(root)
    apply_markup_compatibility(root)
    schema = _schema()
    if schema.validate(root):
        return extension_errors
    return [f"line {error.line}: {error.message}" for error in schema.error_log] + extension_errors


def schema_errors(xml: bytes | etree._Element) -> list[str]:
    """Return the schema errors of one part after Markup Compatibility preprocessing.

    The errors of the part come first, then those of its extension payloads.
    An element argument is copied, not modified.
    """
    if isinstance(xml, bytes):
        return _validate(etree.fromstring(xml, _PARSER))
    return _validate(deepcopy(xml))


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
            root = etree.fromstring(package.read(name), _PARSER)
            if kind != "chart":
                failures.extend(f"{name} {error}" for error in _validate(root))
                continue
            for _reason, rewrite in CHART_ALLOWLIST:
                rewrite(root)
            for error in _validate(root):
                warnings.warn(f"{name} {error}", SchemaWarning, stacklevel=2)
    if failures:
        raise AssertionError("schema-invalid package parts:\n" + "\n".join(failures))
    return checked

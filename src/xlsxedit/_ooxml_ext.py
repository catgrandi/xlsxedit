"""Extension URIs: what the ``ext`` elements of an ``extLst`` hold.

Excel stores features newer than ECMA-376 in extension lists. Each ``ext``
names its payload with a GUID in its ``uri`` attribute. ``EXTENSIONS`` maps
each known URI, spelled as Excel writes it, to the payload element, the
element whose ``extLst`` holds the ``ext``, and the specification section
that lists it. Look a URI up with :func:`extension`, which ignores case: the
specifications print the GUIDs in capitals, Excel mostly in mixed case.

``verified`` says whether a workbook in ``fixtures/`` holds the extension.
``tests/test_ooxml_ext.py`` checks it both ways and fails when a fixture holds
a URI this registry lacks. Run ``scripts/gen_ooxml_order.py`` after a change
here: ``docs/ooxml-schema-notes.md`` lists this registry.
"""

from __future__ import annotations

from typing import NamedTuple

from xlsxedit._ooxml_ns import (
    A14_NS,
    A16_NS,
    A_NS,
    C16_NS,
    C16R3_NS,
    CHART_NS,
    SML_NS,
    THM15_NS,
    X14_NS,
    X15_NS,
    XCALCF_NS,
    XDA_NS,
    XDR_NS,
    XLWCV_NS,
    XPDL_NS,
)


class Extension(NamedTuple):
    """One extension: its payload, where it goes, and where it is specified."""

    feature: str
    namespace: str  # of the payload element
    element: str  # local name of the payload element
    host: str  # Clark name of the element whose extLst holds the ext
    spec: str  # specification section that lists the URI, or "" when none does
    verified: bool  # a workbook in fixtures/ holds it


def _sml(localname: str) -> str:
    return f"{{{SML_NS}}}{localname}"


_WORKSHEET = _sml("worksheet")
_WORKBOOK = _sml("workbook")
_STYLESHEET = _sml("styleSheet")
_SLICER_CACHE = f"{{{X14_NS}}}slicerCacheDefinition"

EXTENSIONS: dict[str, Extension] = {
    "{78C0D931-6437-407d-A8EE-F0AAD7539E65}": Extension(
        "x14 conditional formatting", X14_NS, "conditionalFormattings", _WORKSHEET,
        "[MS-XLSX] 2.2.4.12", True,
    ),
    "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}": Extension(
        "x14 data validation", X14_NS, "dataValidations", _WORKSHEET,
        "[MS-XLSX] 2.2.4.12", True,
    ),
    "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}": Extension(
        "sparklines", X14_NS, "sparklineGroups", _WORKSHEET, "[MS-XLSX] 2.2.4.12", True
    ),
    "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}": Extension(
        "slicers on the sheet", X14_NS, "slicerList", _WORKSHEET, "[MS-XLSX] 2.2.4.12", True
    ),
    "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}": Extension(
        "table slicers on the sheet", X14_NS, "slicerList", _WORKSHEET, "", False
    ),
    "{FC87AEE6-9EDD-4A0A-B7FB-166176984837}": Extension(
        "x14 protected ranges", X14_NS, "protectedRanges", _WORKSHEET,
        "[MS-XLSX] 2.2.4.12", False,
    ),
    "{01252117-D84E-4E92-8308-4BE1C098FCBB}": Extension(
        "x14 ignored errors", X14_NS, "ignoredErrors", _WORKSHEET, "[MS-XLSX] 2.2.4.12", False
    ),
    "{F7C9EE02-42E1-4005-9D12-6889AFFD525C}": Extension(
        "web extensions", X15_NS, "webExtensions", _WORKSHEET, "[MS-XLSX] 2.2.4.12", False
    ),
    "{7E03D99C-DC04-49d9-9315-930204A7B6E9}": Extension(
        "timelines on the sheet", X15_NS, "timelineRefs", _WORKSHEET, "[MS-XLSX] 2.2.4.12", False
    ),
    "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}": Extension(
        "id linking a rule to its x14 twin", X14_NS, "id", _sml("cfRule"),
        "[MS-XLSX] 2.2.4.12", True,
    ),
    "{79F54976-1DA5-4618-B147-4CDE4B953A38}": Extension(
        "x14 workbook properties", X14_NS, "workbookPr", _WORKBOOK, "[MS-XLSX] 2.2.4.11", True
    ),
    "{BBE1A952-AA13-448e-AADC-164F8A28A991}": Extension(
        "slicer caches", X14_NS, "slicerCaches", _WORKBOOK, "[MS-XLSX] 2.2.4.11", True
    ),
    "{140A7094-0E35-4892-8432-C4D2E57EDEB5}": Extension(
        "x15 workbook properties", X15_NS, "workbookPr", _WORKBOOK, "[MS-XLSX] 2.2.4.11", True
    ),
    "{983426D0-5260-488c-9760-48F4B6AC55F4}": Extension(
        "pivot table references", X15_NS, "pivotTableReferences", _WORKBOOK,
        "[MS-XLSX] 2.2.4.11", False,
    ),
    "{A2CB5862-8E78-49c6-8D9D-AF26E26ADB89}": Extension(
        "timeline pivot caches", X15_NS, "timelineCachePivotCaches", _WORKBOOK,
        "[MS-XLSX] 2.2.4.11", False,
    ),
    "{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}": Extension(
        "timeline caches", X15_NS, "timelineCacheRefs", _WORKBOOK, "[MS-XLSX] 2.2.4.11", False
    ),
    "{FCE2AD5D-F65C-4FA6-A056-5C36A1767C68}": Extension(
        "data model", X15_NS, "dataModel", _WORKBOOK, "[MS-XLSX] 2.2.4.11", False
    ),
    "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}": Extension(
        "calculation features", XCALCF_NS, "calcFeatures", _WORKBOOK, "[MS-XLSX] 2.2.4.11", True
    ),
    "{D14903EA-33C4-47F7-8F05-3474C54BE107}": Extension(
        "workbook compatibility version", XLWCV_NS, "version", _WORKBOOK,
        "[MS-XLSX] 2.4.107", True,
    ),
    "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}": Extension(
        "slicer styles", X14_NS, "slicerStyles", _STYLESHEET, "[MS-XLSX] 2.2.4.9", True
    ),
    "{9260A510-F301-46a8-8635-F512D64BE5F5}": Extension(
        "timeline styles", X15_NS, "timelineStyles", _STYLESHEET, "[MS-XLSX] 2.2.4.9", True
    ),
    "{504A1906-F514-4F6F-8877-14C23A59335A}": Extension(
        "x14 table properties", X14_NS, "table", _sml("table"), "[MS-XLSX] 2.2.4.10", False
    ),
    "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}": Extension(
        "dynamic array properties", XDA_NS, "dynamicArrayProperties", _sml("bk"),
        "[MS-XLSX] 2.2.4.4", True,
    ),
    "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}": Extension(
        "x14 pivot table properties", X14_NS, "pivotTableDefinition",
        _sml("pivotTableDefinition"), "[MS-XLSX] 2.2.4.5", True,
    ),
    "{747A6164-185A-40DC-8AA5-F01512510D54}": Extension(
        "pivot table default layout", XPDL_NS, "pivotTableDefinition16",
        _sml("pivotTableDefinition"), "[MS-XLSX] 2.2.4.5", True,
    ),
    "{725AE2AE-9491-48be-B2B4-4EB974FC3084}": Extension(
        "x14 pivot cache properties", X14_NS, "pivotCacheDefinition",
        _sml("pivotCacheDefinition"), "[MS-XLSX] 2.2.4.6", True,
    ),
    "{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}": Extension(
        "table slicer cache", X15_NS, "tableSlicerCache", _SLICER_CACHE,
        "[MS-XLSX] 2.2.4.8", False,
    ),
    "{03082B11-2C62-411c-B77F-237D8FCFBE4C}": Extension(
        "slicer cache pivot tables", X15_NS, "slicerCachePivotTables", _SLICER_CACHE,
        "[MS-XLSX] 2.2.4.8", False,
    ),
    "{470722E0-AACD-4C17-9CDC-17EF765DBC7E}": Extension(
        "slicer cache hides items without data", X15_NS, "slicerCacheHideItemsWithNoData",
        _SLICER_CACHE, "[MS-XLSX] 2.2.4.8", False,
    ),
    "{28A0092B-C50C-407E-A947-70E740481C1C}": Extension(
        "picture keeps its resolution", A14_NS, "useLocalDpi", f"{{{A_NS}}}blip",
        "[MS-ODRAWXML] 2.3.1.13", True,
    ),
    "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}": Extension(
        "drawing object creation id", A16_NS, "creationId", f"{{{XDR_NS}}}cNvPr",
        "[MS-ODRAWXML] 2.25.1.2", True,
    ),
    "{C3380CC4-5D6E-409C-BE32-E72D297353CC}": Extension(
        "chart series unique id", C16_NS, "uniqueId", f"{{{CHART_NS}}}ser",
        "[MS-ODRAWXML] 2.22", True,
    ),
    "{56B9EC1D-385E-4148-901F-78D8002777C0}": Extension(
        "chart data display options", C16R3_NS, "dataDisplayOptions16", f"{{{CHART_NS}}}chart",
        "[MS-ODRAWXML] 2.33.1.1", True,
    ),
    "{05A4C25C-085E-4340-85A3-A5531E510DB2}": Extension(
        "theme family", THM15_NS, "themeFamily", f"{{{A_NS}}}theme",
        "[MS-ODRAWXML] 2.4.1.1", True,
    ),
}

_BY_KEY = {uri.upper(): ext for uri, ext in EXTENSIONS.items()}


def extension(uri: str) -> Extension | None:
    """The registered extension for ``uri``, compared without regard to case, or ``None``."""
    return _BY_KEY.get(uri.upper())

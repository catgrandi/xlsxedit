"""SpreadsheetML / OPC content-type and relationship-type constants."""

from __future__ import annotations


class CT:
    OPC_RELATIONSHIPS = "application/vnd.openxmlformats-package.relationships+xml"
    OPC_CORE_PROPERTIES = "application/vnd.openxmlformats-package.core-properties+xml"
    XML = "application/xml"
    WORKBOOK = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    )
    WORKBOOK_MACRO_ENABLED = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
    WORKBOOK_TEMPLATE = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
    )
    WORKBOOK_MACRO_ENABLED_TEMPLATE = (
        "application/vnd.ms-excel.template.macroEnabled.main+xml"
    )
    WORKSHEET = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
    )
    SHARED_STRINGS = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
    )
    STYLES = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
    THEME = "application/vnd.openxmlformats-officedocument.theme+xml"
    DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
    CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
    TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
    CALC_CHAIN = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"
    )
    JPEG = "image/jpeg"
    PNG = "image/png"


class RT:
    OFFICE_DOCUMENT = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
        "officeDocument"
    )
    CORE_PROPERTIES = (
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/"
        "core-properties"
    )
    WORKSHEET = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    )
    SHARED_STRINGS = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
        "sharedStrings"
    )
    STYLES = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    )
    THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    DRAWING = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
    )
    IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    TABLE = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
    )
    HYPERLINK = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    )
    CALC_CHAIN = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
        "calcChain"
    )


NSMAP = {
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

SML_NS = NSMAP["main"]
REL_NS = NSMAP["pr"]
CT_NS = NSMAP["ct"]
OFFICE_REL_NS = NSMAP["r"]

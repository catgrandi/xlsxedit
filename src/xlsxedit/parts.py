"""Typed SpreadsheetML package parts."""

from __future__ import annotations

from lxml import etree

from xlsxedit.exceptions import MissingPartError
from xlsxedit.opc.constants import CT, RT, SML_NS
from xlsxedit.opc.part import XmlPart
from xlsxedit.oxml.parser import parse_xml

_OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_SHEET = f"{{{SML_NS}}}sheet"
_SHEETS = f"{{{SML_NS}}}sheets"


class WorkbookPart(XmlPart):
    """``xl/workbook.xml``."""

    @property
    def sheet_elements(self):
        sheets = self.element.find(_SHEETS)
        if sheets is None:
            return []
        return list(sheets.findall(_SHEET))

    def worksheet_parts(self) -> list[tuple[str, WorksheetPart]]:
        """Return ``(sheet_name, WorksheetPart)`` in workbook order."""
        result = []
        for sheet_elm in self.sheet_elements:
            name = sheet_elm.get("name")
            r_id = sheet_elm.get(f"{{{_OFFICE_REL_NS}}}id")
            if r_id is None:
                continue
            part = self.rels[r_id].target_part
            result.append((name, part))
        return result

    @property
    def shared_strings_part(self) -> SharedStringsPart | None:
        try:
            return self.rels.part_with_reltype(RT.SHARED_STRINGS)
        except KeyError:
            return None

    @property
    def styles_part(self) -> StylesPart | None:
        try:
            return self.rels.part_with_reltype(RT.STYLES)
        except KeyError:
            return None

    def remove_sheet_element(self, name: str) -> tuple[str, str] | None:
        """Remove ``<sheet>`` by name; return ``(r_id, sheet_id)`` or None."""
        for sheet_elm in self.sheet_elements:
            if sheet_elm.get("name") == name:
                r_id = sheet_elm.get(f"{{{_OFFICE_REL_NS}}}id")
                sheet_id = sheet_elm.get("sheetId")
                sheets = self.element.find(_SHEETS)
                if sheets is not None:
                    sheets.remove(sheet_elm)
                return r_id, sheet_id
        return None

    def next_sheet_id(self) -> int:
        ids = [int(elm.get("sheetId", "0")) for elm in self.sheet_elements]
        return max(ids, default=0) + 1

    def append_sheet_element(self, name: str, sheet_id: int, r_id: str) -> None:
        sheets = self.element.find(_SHEETS)
        if sheets is None:
            sheets = etree.SubElement(self.element, _SHEETS)
        sheet_elm = etree.SubElement(sheets, _SHEET)
        sheet_elm.set("name", name)
        sheet_elm.set("sheetId", str(sheet_id))
        sheet_elm.set(f"{{{_OFFICE_REL_NS}}}id", r_id)


class WorksheetPart(XmlPart):
    """``xl/worksheets/sheetN.xml``."""

    @classmethod
    def load_deferred(cls, partname, content_type, blob, package=None):
        """Keep the source bytes and parse them on first ``element`` access."""
        part = cls(partname, content_type, None, package)
        part._orig_blob = blob
        return part

    @property
    def unparsed_blob(self) -> bytes | None:
        """The source bytes while the part is still deferred, else ``None``.

        Lets a caller scan a sheet opened with ``large=True`` for a marker
        without parsing it; reading :attr:`element` ends the deferral.
        """
        return self._orig_blob if self._element is None else None

    @property
    def element(self) -> etree._Element:
        if self._element is None:
            if self._orig_blob is None:
                raise MissingPartError("worksheet part has no XML content")
            self._element = parse_xml(self._orig_blob)
        return super().element


class SharedStringsPart(XmlPart):
    """``xl/sharedStrings.xml``."""


class StylesPart(XmlPart):
    """``xl/styles.xml`` — live stylesheet element.

    ``Styles`` calls :meth:`mark_dirty` when it mutates the stylesheet; until
    then an untouched ``styles.xml`` round-trips byte-identically.
    """


class LazyXmlPart(XmlPart):
    """An ``XmlPart`` that parses its source bytes on first ``.element`` access.

    Until then the bytes are written back without being parsed.
    """

    @classmethod
    def load(cls, partname, content_type, blob, package=None):
        part = cls(partname, content_type, None, package)
        part._orig_blob = blob
        return part

    @property
    def element(self) -> etree._Element:
        if self._element is None:
            if self._orig_blob is None:
                raise MissingPartError(f"{self.partname} has no XML content")
            self._element = parse_xml(self._orig_blob)
        return super().element


class DrawingPart(LazyXmlPart):
    """``xl/drawings/drawingN.xml`` — the one live ``xdr:wsDr`` tree of the part."""


class ChartPart(LazyXmlPart):
    """``xl/charts/chartN.xml`` — the one live ``c:chartSpace`` tree of the part."""


class CorePropertiesPart(XmlPart):
    """``/docProps/core.xml`` — Dublin Core document metadata."""

    @classmethod
    def default(cls, package) -> CorePropertiesPart:
        from xlsxedit.coreprops import new_core_properties_element
        from xlsxedit.opc.packuri import PackURI

        return cls(
            PackURI("/docProps/core.xml"),
            CT.OPC_CORE_PROPERTIES,
            new_core_properties_element(),
            package,
        )

    @property
    def core_properties(self):
        from xlsxedit.coreprops import CoreProperties

        return CoreProperties(self.element)


def register_part_types() -> None:
    from xlsxedit.opc.part import PartFactory

    PartFactory.part_type_for[CT.WORKBOOK] = WorkbookPart
    PartFactory.part_type_for[CT.WORKBOOK_MACRO_ENABLED] = WorkbookPart
    PartFactory.part_type_for[CT.WORKBOOK_TEMPLATE] = WorkbookPart
    PartFactory.part_type_for[CT.WORKBOOK_MACRO_ENABLED_TEMPLATE] = WorkbookPart
    PartFactory.part_type_for[CT.WORKSHEET] = WorksheetPart
    PartFactory.part_type_for[CT.SHARED_STRINGS] = SharedStringsPart
    PartFactory.part_type_for[CT.STYLES] = StylesPart
    PartFactory.part_type_for[CT.OPC_CORE_PROPERTIES] = CorePropertiesPart
    PartFactory.part_type_for[CT.DRAWING] = DrawingPart
    PartFactory.part_type_for[CT.CHART] = ChartPart

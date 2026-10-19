"""Data validations: a worksheet's ``dataValidation`` rules and their x14 form.

A worksheet holds its validations in one ``dataValidations`` element. Excel
2010 and later write a validation whose formulas refer to another sheet into
the worksheet's ``extLst`` instead, as an ``x14:dataValidation`` with
``xm:sqref`` and ``xm:f`` children. Both kinds are listed.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from xlsxedit._ooxml_order import ATTRIBUTE_DEFAULTS
from xlsxedit.opc.constants import SML_NS
from xlsxedit.range_set import SheetRangeSet

if TYPE_CHECKING:
    from xlsxedit.worksheet import Worksheet

_X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
_XM_NS = "http://schemas.microsoft.com/office/excel/2006/main"
_X14_URI = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"

_DATA_VALIDATIONS = f"{{{SML_NS}}}dataValidations"
_DATA_VALIDATION = f"{{{SML_NS}}}dataValidation"
_EXT_LST = f"{{{SML_NS}}}extLst"
_EXT = f"{{{SML_NS}}}ext"
_X14_DATA_VALIDATIONS = f"{{{_X14_NS}}}dataValidations"
_X14_DATA_VALIDATION = f"{{{_X14_NS}}}dataValidation"
_XM_SQREF = f"{{{_XM_NS}}}sqref"

_DEFAULTS = ATTRIBUTE_DEFAULTS["CT_DataValidation"]


def _xsd_bool(text: str) -> bool:
    return text.strip() in ("1", "true")


class DataValidation:
    """One validation rule and the cells it applies to.

    From ``ws.data_validations``. An attribute the XML leaves out reads as
    its schema default.
    """

    def __init__(self, element: etree._Element, worksheet: Worksheet):
        self._element = element
        self._worksheet = worksheet

    def __repr__(self) -> str:
        return f"<DataValidation {self.type} {self._sqref_text()!r}>"

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, DataValidation):
            return NotImplemented
        return self._element is other._element

    def __hash__(self) -> int:
        return hash(self._element)

    @property
    def is_x14(self) -> bool:
        """Whether the rule is an ``x14:dataValidation`` in the extension list."""
        return self._element.tag == _X14_DATA_VALIDATION

    def _sqref_text(self) -> str:
        if self.is_x14:
            sqref = self._element.find(_XM_SQREF)
            return "" if sqref is None else sqref.text or ""
        return self._element.get("sqref", "")

    @property
    def sqref(self) -> SheetRangeSet:
        """The cells the rule applies to.

        Raises ``InvalidRangeError`` when the stored range list is malformed.
        """
        return SheetRangeSet.parse(self._sqref_text())

    def _attribute(self, name: str) -> str | None:
        return self._element.get(name, _DEFAULTS.get(name))

    @property
    def type(self) -> str:
        """``none``, ``whole``, ``decimal``, ``list``, ``date``, ``time``, ``textLength``
        or ``custom``."""
        return self._attribute("type")

    @property
    def operator(self) -> str:
        """How ``whole`` to ``textLength`` rules compare: ``between``, ``greaterThan``, …"""
        return self._attribute("operator")

    @property
    def error_style(self) -> str:
        """``stop``, ``warning`` or ``information``."""
        return self._attribute("errorStyle")

    @property
    def allow_blank(self) -> bool:
        return _xsd_bool(self._attribute("allowBlank"))

    @property
    def show_input_message(self) -> bool:
        return _xsd_bool(self._attribute("showInputMessage"))

    @property
    def show_error_message(self) -> bool:
        return _xsd_bool(self._attribute("showErrorMessage"))

    @property
    def error_title(self) -> str | None:
        return self._attribute("errorTitle")

    @property
    def error(self) -> str | None:
        return self._attribute("error")

    @property
    def prompt_title(self) -> str | None:
        return self._attribute("promptTitle")

    @property
    def prompt(self) -> str | None:
        return self._attribute("prompt")

    @property
    def formula1(self) -> str | None:
        """The first formula, without a leading ``=``; ``None`` when absent."""
        return self._formula("formula1")

    @property
    def formula2(self) -> str | None:
        """The second formula of a ``between`` or ``notBetween`` rule."""
        return self._formula("formula2")

    def _formula(self, name: str) -> str | None:
        if self.is_x14:
            elm = self._element.find(f"{{{_X14_NS}}}{name}/{{{_XM_NS}}}f")
        else:
            elm = self._element.find(f"{{{SML_NS}}}{name}")
        return None if elm is None else elm.text or ""


def _x14_elements(root: etree._Element):
    for ext_lst in root.iterchildren(_EXT_LST):
        for ext in ext_lst.iterchildren(_EXT):
            if (ext.get("uri") or "").upper() != _X14_URI.upper():
                continue
            for block in ext.iterchildren(_X14_DATA_VALIDATIONS):
                yield from block.iterchildren(_X14_DATA_VALIDATION)


def iter_data_validations(worksheet: Worksheet) -> list[DataValidation]:
    """The sheet's ``dataValidation`` rules in document order, then its x14 rules."""
    root = worksheet._part.element
    rules = [
        elm
        for block in root.iterchildren(_DATA_VALIDATIONS)
        for elm in block.iterchildren(_DATA_VALIDATION)
    ]
    rules += _x14_elements(root)
    return [DataValidation(elm, worksheet) for elm in rules]


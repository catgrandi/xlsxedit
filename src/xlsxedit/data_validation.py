"""Data validations: a worksheet's ``dataValidation`` rules and their x14 form.

A worksheet holds its validations in one ``dataValidations`` element. Excel
2010 and later write a validation whose formulas refer to another sheet into
the worksheet's ``extLst`` instead, as an ``x14:dataValidation`` with
``xm:sqref`` and ``xm:f`` children. Both kinds are listed; only the first can
be changed or removed.

Every argument is checked before the worksheet changes, so a refused edit
leaves it as it was.
"""

from __future__ import annotations

import math
import re
import sys
from typing import TYPE_CHECKING

from lxml import etree

from xlsxedit._ooxml_order import ATTRIBUTE_DEFAULTS, ATTRIBUTE_ORDER, ENUMS, ORDER
from xlsxedit.exceptions import DataValidationError, InvalidRangeError
from xlsxedit.opc.constants import SML_NS
from xlsxedit.range_set import CellRange, SheetRangeSet
from xlsxedit.styles import _XML_ILLEGAL
from xlsxedit.worksheet_order import insert_ordered, insert_worksheet_child

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

_ATTRIBUTES = ATTRIBUTE_ORDER["CT_DataValidation"]
_DEFAULTS = ATTRIBUTE_DEFAULTS["CT_DataValidation"]
_CHILDREN = ORDER["CT_DataValidation"]
TYPES = ENUMS["ST_DataValidationType"]
OPERATORS = ENUMS["ST_DataValidationOperator"]
ERROR_STYLES = ENUMS["ST_DataValidationErrorStyle"]

# Types that compare a cell's value with formula1, and with formula2 for a range.
_COMPARED = frozenset({"whole", "decimal", "date", "time", "textLength"})
_RANGE_OPERATORS = frozenset({"between", "notBetween"})

# Excel's limits, in UTF-16 code units; XlsxWriter enforces the same title, message
# and list limits.
_TITLE_LIMIT = 32
_MESSAGE_LIMIT = 255
_LIST_LIMIT = 255
_FORMULA_LIMIT = 8192

_STRING_LITERAL = re.compile(r'"((?:[^"]|"")*)"')
# Read left to right: a string literal, a quoted or plain sheet name before "!", or a lone "!".
_FORMULA_TOKENS = re.compile(
    r'"(?:[^"]|"")*"'
    r"|'((?:[^']|'')*)'!"
    r"|(?<![\w.\[\]:])([\w.\[\]:]+)!"
    r"|!"
)


def _utf16_len(text: str) -> int:
    return len(text.encode("utf-16-le")) // 2


def _xsd_bool(text: str) -> bool:
    return text.strip() in ("1", "true")


class DataValidation:
    """One validation rule and the cells it applies to.

    From ``ws.data_validations`` or ``ws.add_data_validation(...)``. An
    attribute the XML leaves out reads as its schema default. A rule stored
    in the x14 extension list (``is_x14``) is read-only.
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
        Setting it takes a range string such as ``"B2:B10 D2"``, a
        ``CellRange`` or a ``SheetRangeSet``, written as Excel writes it
        (``A:A`` becomes ``A1:A1048576``). Setting raises, changing nothing,
        ``DataValidationError`` for an x14 or removed rule and for a range
        that overlaps another validation on the sheet.
        """
        return SheetRangeSet.parse(self._sqref_text())

    @sqref.setter
    def sqref(self, value: str | CellRange | SheetRangeSet) -> None:
        worksheet = self._editable()
        ranges = _check_sqref(value)
        _check_overlap(worksheet, ranges, self._element)
        self._element.set("sqref", str(ranges))

    def _editable(self) -> Worksheet:
        if self.is_x14:
            raise DataValidationError("an x14 data validation is read-only")
        if not _is_listed(self._worksheet, self._element):
            raise DataValidationError("the data validation is no longer on its worksheet")
        return self._worksheet

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


def _is_listed(worksheet: Worksheet, element: etree._Element) -> bool:
    block = element.getparent()
    return (
        element.tag == _DATA_VALIDATION
        and block is not None
        and block.tag == _DATA_VALIDATIONS
        and block.getparent() is worksheet._part.element
    )


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


def _check_sqref(value: object) -> SheetRangeSet:
    if isinstance(value, SheetRangeSet):
        ranges = value
    elif isinstance(value, CellRange):
        ranges = SheetRangeSet((value,))
    elif isinstance(value, str):
        ranges = SheetRangeSet.parse(value)
    else:
        raise TypeError(
            f"sqref must be a str, CellRange or SheetRangeSet, got {type(value).__name__}"
        )
    if not ranges:
        raise InvalidRangeError(f"sqref {value!r} names no cells")
    areas = (CellRange.from_bounds(r.min_col, r.min_row, r.max_col, r.max_row) for r in ranges)
    return SheetRangeSet(dict.fromkeys(areas))


def _check_overlap(
    worksheet: Worksheet, ranges: SheetRangeSet, own: etree._Element | None = None
) -> None:
    for other in iter_data_validations(worksheet):
        if other._element is not own and other.sqref.intersects(ranges):
            raise DataValidationError(
                f"sqref {str(ranges)!r} overlaps the validation on {str(other.sqref)!r}; "
                "Excel keeps one validation per cell"
            )


def _check_choice(value: object, name: str, allowed: tuple[str, ...]) -> str:
    if not isinstance(value, str):
        raise TypeError(f"{name} must be a str, got {type(value).__name__}")
    if value not in allowed:
        raise DataValidationError(f"{name} must be one of {', '.join(allowed)}, got {value!r}")
    return value


def _check_flag(value: object, name: str) -> bool:
    if not isinstance(value, bool):
        raise TypeError(f"{name} must be a bool, got {type(value).__name__}")
    return value


def _check_text(value: object, name: str, limit: int) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str):
        raise TypeError(f"{name} must be a str, got {type(value).__name__}")
    if _XML_ILLEGAL.search(value):
        raise DataValidationError(f"{name} {value!r} holds characters XML cannot store")
    if _utf16_len(value) > limit:
        raise DataValidationError(f"{name} is longer than the {limit} characters Excel allows")
    return value


def _check_formula(value: object, name: str, sheet_name: str) -> str | None:
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, (str, int, float)):
        raise TypeError(f"{name} must be a str, int or float, got {type(value).__name__}")
    if isinstance(value, str):
        text = value[1:] if value.startswith("=") else value
    elif abs(value) > sys.float_info.max or (isinstance(value, float) and math.isnan(value)):
        raise DataValidationError(f"{name} {value!r} is not a number Excel can store")
    else:
        text = repr(value) if isinstance(value, float) else str(value)
    if not text.strip():
        raise DataValidationError(f"{name} is empty")
    if _XML_ILLEGAL.search(text):
        raise DataValidationError(f"{name} {text!r} holds characters XML cannot store")
    if _utf16_len(text) > _FORMULA_LIMIT:
        raise DataValidationError(
            f"{name} is longer than the {_FORMULA_LIMIT} characters Excel allows"
        )
    _check_same_sheet(text, name, sheet_name)
    return text


def _check_same_sheet(formula: str, name: str, sheet_name: str) -> None:
    """Refuse a reference to another sheet or workbook, which needs the x14 form."""
    if "!" not in formula:
        return
    for token in _FORMULA_TOKENS.finditer(formula):
        if token.group(0).startswith('"'):
            continue
        quoted, plain = token.groups()
        sheet = plain if quoted is None else quoted.replace("''", "'")
        if sheet is None or sheet.casefold() != sheet_name.casefold():
            raise DataValidationError(
                f"{name} {formula!r} refers to another sheet or workbook; such a validation "
                "needs the x14 form, which xlsxedit cannot write yet, so refer to a defined "
                "name instead"
            )


def _check_list_literal(formula: str) -> None:
    literal = _STRING_LITERAL.fullmatch(formula)
    if literal is not None and _utf16_len(literal.group(1).replace('""', '"')) > _LIST_LIMIT:
        raise DataValidationError(
            f"the list {formula!r} is longer than the {_LIST_LIMIT} characters Excel allows; "
            "list the items in cells and refer to their range instead"
        )


def _check_formulas(
    kind: str, operator: str | None, formula1: str | None, formula2: str | None
) -> None:
    if operator is not None and kind not in _COMPARED:
        raise DataValidationError(
            "operator applies to whole, decimal, date, time and textLength validations, "
            f"not {kind!r}"
        )
    if kind == "none":
        if formula1 is not None or formula2 is not None:
            raise DataValidationError("a 'none' validation takes no formula")
        return
    if formula1 is None:
        raise DataValidationError(f"a {kind!r} validation needs formula1")
    effective = operator or _DEFAULTS["operator"]
    if kind in _COMPARED and effective in _RANGE_OPERATORS:
        if formula2 is None:
            raise DataValidationError(f"operator {effective!r} needs formula2")
    elif formula2 is not None:
        raise DataValidationError("formula2 applies only to the between and notBetween operators")
    if kind == "list":
        _check_list_literal(formula1)


def new_data_validation(
    worksheet: Worksheet,
    sqref: str | CellRange | SheetRangeSet,
    *,
    type: str,
    formula1: str | int | float | None = None,
    formula2: str | int | float | None = None,
    operator: str | None = None,
    allow_blank: bool = True,
    show_input_message: bool = True,
    show_error_message: bool = True,
    error_style: str = "stop",
    error_title: str | None = None,
    error: str | None = None,
    prompt_title: str | None = None,
    prompt: str | None = None,
) -> DataValidation:
    """Append a ``dataValidation`` to the worksheet after checking every argument."""
    ranges = _check_sqref(sqref)
    kind = _check_choice(type, "type", TYPES)
    if operator is not None:
        _check_choice(operator, "operator", OPERATORS)
    formulas = [
        _check_formula(formula1, "formula1", worksheet.name),
        _check_formula(formula2, "formula2", worksheet.name),
    ]
    _check_formulas(kind, operator, *formulas)
    values: dict[str, str | bool | None] = {
        "type": kind,
        "errorStyle": _check_choice(error_style, "error_style", ERROR_STYLES),
        "operator": operator,
        "allowBlank": _check_flag(allow_blank, "allow_blank"),
        "showInputMessage": _check_flag(show_input_message, "show_input_message"),
        "showErrorMessage": _check_flag(show_error_message, "show_error_message"),
        "errorTitle": _check_text(error_title, "error_title", _TITLE_LIMIT),
        "error": _check_text(error, "error", _MESSAGE_LIMIT),
        "promptTitle": _check_text(prompt_title, "prompt_title", _TITLE_LIMIT),
        "prompt": _check_text(prompt, "prompt", _MESSAGE_LIMIT),
        "sqref": str(ranges),
    }
    _check_overlap(worksheet, ranges)

    elm = etree.Element(_DATA_VALIDATION)
    for name in _ATTRIBUTES:
        value = values.get(name)
        if isinstance(value, bool):
            if value != _xsd_bool(_DEFAULTS[name]):
                elm.set(name, "1" if value else "0")
        elif value is not None and value != _DEFAULTS.get(name):
            elm.set(name, value)
    for name, text in zip(("formula1", "formula2"), formulas):
        if text is not None:
            child = etree.Element(f"{{{SML_NS}}}{name}")
            child.text = text
            insert_ordered(elm, child, _CHILDREN)

    root = worksheet._part.element
    block = root.find(_DATA_VALIDATIONS)
    if block is None:
        block = etree.Element(_DATA_VALIDATIONS)
        insert_worksheet_child(root, block)
    block.append(elm)
    block.set("count", str(len(block.findall(_DATA_VALIDATION))))
    return DataValidation(elm, worksheet)


def delete_data_validation(worksheet: Worksheet, validation: DataValidation) -> None:
    """Remove a ``dataValidation``, and its container when it was the last one."""
    if not isinstance(validation, DataValidation):
        raise TypeError(f"expected a DataValidation, got {type(validation).__name__}")
    if validation.is_x14:
        raise DataValidationError("an x14 data validation is read-only")
    elm = validation._element
    if not _is_listed(worksheet, elm):
        raise DataValidationError(f"the data validation is not on worksheet {worksheet.name!r}")
    block = elm.getparent()
    block.remove(elm)
    remaining = len(block.findall(_DATA_VALIDATION))
    if remaining:
        block.set("count", str(remaining))
    else:
        worksheet._part.element.remove(block)

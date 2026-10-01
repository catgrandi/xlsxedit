"""Cell proxy — value get/set preserves the style index."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any
from datetime import date, datetime

from lxml import etree
from lxml.etree import _Element

from xlsxedit.exceptions import FormulaGroupError, InvalidRangeError
from xlsxedit.merge import address_in_range, parse_range
from xlsxedit.opc.constants import SML_NS
from xlsxedit.shared_strings import si_display_text
from xlsxedit.styles import CellStyle, serial_to_datetime, datetime_to_serial
from xlsxedit.hyperlinks import Hyperlink

if TYPE_CHECKING:
    from xlsxedit.worksheet import Worksheet

_V = f"{{{SML_NS}}}v"
_F = f"{{{SML_NS}}}f"
_IS = f"{{{SML_NS}}}is"
_T = f"{{{SML_NS}}}t"
_R = f"{{{SML_NS}}}r"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_RANGE_FORMULA_TYPES = ("array", "dataTable")
_RANGE_FORMULAS = etree.XPath(
    "m:sheetData/m:row/m:c/m:f[@t='array' or @t='dataTable']", namespaces={"m": SML_NS}
)
_SHARED_FORMULAS = etree.XPath(
    "m:sheetData/m:row/m:c/m:f[@t='shared' and @si=$si]", namespaces={"m": SML_NS}
)


def _ref_contains(ref: str | None, address: str) -> bool:
    if not ref:
        return False
    try:
        return address_in_range(address, ref)
    except InvalidRangeError:
        return False


def _spans_several_cells(ref: str | None) -> bool:
    if not ref:
        return False
    try:
        c1, r1, c2, r2 = parse_range(ref)
    except InvalidRangeError:
        return False
    return (c1, r1) != (c2, r2)


def _refuse_range_formula_edit(address: str, f_elm: _Element) -> None:
    """Raise if ``f_elm`` anchors an array or data-table formula over several cells."""
    kind, ref = f_elm.get("t"), f_elm.get("ref")
    if kind in _RANGE_FORMULA_TYPES and _spans_several_cells(ref):
        raise FormulaGroupError(
            f"{address} anchors the {kind} formula over {ref}; "
            "editing part of a multi-cell formula is not supported"
        )


def _inline_t_nodes(is_elm: _Element) -> list[_Element]:
    direct = list(is_elm.findall(_T))
    if direct:
        return direct
    nodes: list[_Element] = []
    for r in is_elm.findall(_R):
        nodes.extend(r.findall(_T))
    return nodes


def _clear_cell_content(c_elm: _Element) -> None:
    for child in list(c_elm):
        if child.tag in (_V, _F, _IS):
            c_elm.remove(child)


def _ensure_v(c_elm: _Element) -> _Element:
    v = c_elm.find(_V)
    if v is None:
        v = etree.SubElement(c_elm, _V)
    return v


def _coerce_numeric_cache(v: str) -> int | float | str:
    """Parse a ``<v>`` text as int/float; return ``v`` unchanged if not numeric."""
    try:
        return int(v)
    except ValueError:
        pass
    try:
        return float(v)
    except ValueError:
        return v


def cell_matches(cell: Cell, value: Any) -> bool:
    """True if ``cell`` is a non-formula cell whose ``value`` equals ``value``."""
    if cell.has_formula:
        return False
    return cell.value == value


class Cell:
    """Thin proxy over a ``<c>`` element."""

    def __init__(self, element: _Element, worksheet: Worksheet):
        self._element = element
        self._worksheet = worksheet

    def __repr__(self) -> str:
        return f"<Cell {self.address!r}>"

    @property
    def address(self) -> str:
        return self._element.get("r", "")

    @property
    def worksheet(self) -> Worksheet:
        return self._worksheet

    def offset(self, cols: int = 0, rows: int = 0) -> Cell:
        """Return the cell ``cols`` columns right and ``rows`` rows down (negative allowed)."""
        from xlsxedit.oxml.address import (
            col_to_index,
            index_to_col,
            join_address,
            split_address,
        )

        col, row = split_address(self.address)
        new_col = col_to_index(col) + cols
        new_row = row + rows
        if new_col < 0 or new_row < 1:
            raise ValueError(
                f"offset ({cols}, {rows}) from {self.address!r} is out of range"
            )
        return self._worksheet[join_address(index_to_col(new_col), new_row)]

    @property
    def style_index(self) -> str | None:
        return self._element.get("s")

    @property
    def has_formula(self) -> bool:
        return self._element.find(_F) is not None

    @property
    def data_type(self) -> str | None:
        return self._element.get("t")

    @property
    def is_string(self) -> bool:
        t = self.data_type
        return t in ("s", "inlineStr", "str") and not self.has_formula

    def _v_text(self) -> str | None:
        v = self._element.find(_V)
        if v is None:
            return None
        return v.text

    @property
    def formula(self) -> str | None:
        """Formula text as Excel stores it, without a leading ``=``.

        ``None`` without a formula, and on a shared-formula follower, whose
        text lives on the group's master. The setter accepts an optional
        leading ``=``, drops the cached result and its ``t`` type, and sets
        ``fullCalcOnLoad`` so Excel computes the new formula on open. A
        follower becomes a standalone formula, and a single-cell array
        formula stays an array formula. ``None`` removes the formula.
        """
        f = self._element.find(_F)
        if f is None:
            return None
        return f.text

    @formula.setter
    def formula(self, expr: str | None) -> None:
        if expr is None:
            f = self._element.find(_F)
            if f is not None:
                self._worksheet._workbook._on_formula_removed(self._worksheet, self._element)
                self._element.remove(f)
            return
        if not isinstance(expr, str):
            raise TypeError(f"Cell.formula accepts str or None; got {type(expr)!r}")
        text = expr[1:] if expr.startswith("=") else expr
        if not text.strip():
            raise ValueError(f"formula for {self.address!r} is empty")
        f = self._element.find(_F)
        self._refuse_formula_group_edit(f)
        keep = f is not None and f.get("t", "normal") in ("normal", "array")
        for child in list(self._element):
            if child.tag in (_V, _IS) or (child.tag == _F and not keep):
                self._element.remove(child)
        self._element.attrib.pop("t", None)
        if not keep:
            f = etree.SubElement(self._element, _F)
            self._element.insert(0, f)
        f.text = text
        self._worksheet._workbook.set_full_calc_on_load()

    @property
    def formula_type(self) -> str | None:
        """``"normal"``, ``"shared"``, ``"array"``, or ``"dataTable"``; ``None`` without ``<f>``.

        Only an array or data-table formula's anchor cell carries ``<f>``; the
        other cells of its range return ``None``.
        """
        f = self._element.find(_F)
        if f is None:
            return None
        return f.get("t", "normal")

    def _refuse_formula_group_edit(self, f: _Element | None) -> None:
        """Raise unless a new formula here leaves every other cell's formula as it was."""
        address = self.address
        sheet = self._worksheet._part.element
        if f is not None:
            _refuse_range_formula_edit(address, f)
            si = f.get("si")
            is_master = f.get("ref") is not None or bool((f.text or "").strip())
            if f.get("t") == "shared" and is_master and si is not None:
                others = [m for m in _SHARED_FORMULAS(sheet, si=si) if m is not f]
                if others:
                    raise FormulaGroupError(
                        f"{address} holds shared formula si={si}, which "
                        f"{len(others)} other cell(s) derive from; give those cells "
                        f"their own formulas first, or set {address}'s value to turn "
                        "them into plain values"
                    )
        for master in _RANGE_FORMULAS(sheet):
            if master is not f and _ref_contains(master.get("ref"), address):
                anchor = master.getparent().get("r", "?")
                raise FormulaGroupError(
                    f"{address} lies in the {master.get('t')} formula anchored at {anchor} "
                    f"over {master.get('ref')}; editing part of a multi-cell formula "
                    "is not supported"
                )

    @property
    def style(self) -> CellStyle:
        return CellStyle(
            self._worksheet._workbook.styles,
            self.style_index,
            self._element,
        )

    @property
    def hyperlink(self) -> Hyperlink:
        return Hyperlink(self)

    def clear(self) -> None:
        """Clear cell content while preserving style index."""
        self.value = None

    def apply_date_format(self) -> None:
        """Apply a built-in date number format to this cell."""
        styles = self._worksheet._workbook.styles
        styles.apply_date_format(self._element)

    def apply_number_format(self, format_code: str) -> None:
        """Apply a custom number format string to this cell."""
        styles = self._worksheet._workbook.styles
        styles.apply_num_format(self._element, format_code)

    def apply_style(
        self,
        *,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        font_size: int | None = None,
        font_name: str | None = None,
        font_color: str | None = None,
        bg_color: str | None = None,
        horizontal_align: str | None = None,
        vertical_align: str | None = None,
    ) -> None:
        """Apply whole-cell font, fill, and alignment via a cloned ``xf``."""
        kwargs = {
            "bold": bold,
            "italic": italic,
            "underline": underline,
            "font_size": font_size,
            "font_name": font_name,
            "font_color": font_color,
            "bg_color": bg_color,
            "horizontal_align": horizontal_align,
            "vertical_align": vertical_align,
        }
        filtered = {k: v for k, v in kwargs.items() if v is not None}
        if not filtered:
            return
        self._worksheet._workbook.styles.apply_cell_style(self._element, **filtered)

    @property
    def value(self) -> str | float | int | bool | datetime | date | None:
        """Display/string value for string cells; numeric raw value otherwise."""
        if self.has_formula:
            v = self._v_text()
            if v is None:
                return None
            t = self.data_type
            if t == "e":
                return v
            if t == "str":
                return v
            if t == "b":
                return v == "1"
            numeric = _coerce_numeric_cache(v)
            if isinstance(numeric, (int, float)) and self.style.is_date:
                return serial_to_datetime(float(numeric))
            return numeric
        t = self.data_type
        if t == "s":
            idx_text = self._v_text()
            if idx_text is None:
                return None
            sst = self._worksheet.shared_strings
            if sst is None:
                return None
            return sst.text(int(idx_text))
        if t == "inlineStr":
            is_elm = self._element.find(_IS)
            if is_elm is None:
                return None
            return "".join((n.text or "") for n in _inline_t_nodes(is_elm))
        if t == "str":
            return self._v_text()
        if t == "b":
            v = self._v_text()
            return v == "1" if v is not None else None
        v = self._v_text()
        if v is None:
            return None
        if t == "e":
            return v
        numeric = _coerce_numeric_cache(v)
        if isinstance(numeric, (int, float)) and self.style.is_date:
            return serial_to_datetime(float(numeric))
        return numeric

    @value.setter
    def value(self, new_value: Any) -> None:
        """Set cell content while preserving ``s`` (style index)."""
        if self.has_formula:
            self._worksheet._workbook._on_formula_removed(self._worksheet, self._element)

        if new_value is None:
            _clear_cell_content(self._element)
            if "t" in self._element.attrib:
                del self._element.attrib["t"]
            return

        if isinstance(new_value, bool):
            _clear_cell_content(self._element)
            self._element.set("t", "b")
            _ensure_v(self._element).text = "1" if new_value else "0"
            return

        if isinstance(new_value, (datetime, date)):
            _clear_cell_content(self._element)
            if "t" in self._element.attrib:
                del self._element.attrib["t"]
            serial = datetime_to_serial(new_value)
            v = _ensure_v(self._element)
            v.text = str(int(serial)) if serial == int(serial) else repr(serial)
            return

        if isinstance(new_value, (int, float)) and not isinstance(new_value, bool):
            _clear_cell_content(self._element)
            if "t" in self._element.attrib:
                del self._element.attrib["t"]
            v = _ensure_v(self._element)
            if isinstance(new_value, float):
                v.text = repr(new_value)
            else:
                v.text = str(new_value)
            return

        if not isinstance(new_value, str):
            raise TypeError(
                f"Cell.value accepts str, int, float, bool, datetime, date, or None; "
                f"got {type(new_value)!r}"
            )

        self._set_string_value(new_value)

    def _set_string_value(self, new_value: str) -> None:
        sst = self._worksheet.shared_strings
        if sst is None:
            raise ValueError("workbook has no shared strings part")

        t = self.data_type
        if t == "s":
            idx_text = self._v_text()
            if idx_text is not None:
                idx = int(idx_text)
                if self._worksheet._workbook._sst_index_ref_count(idx) <= 1:
                    from xlsxedit.shared_strings import _t_nodes_in_si

                    si = sst.get(idx)
                    nodes = _t_nodes_in_si(si)
                    if len(nodes) == 1:
                        nodes[0].text = new_value
                        if " " in new_value or new_value.startswith(" ") or new_value.endswith(
                            " "
                        ):
                            nodes[0].set(_XML_SPACE, "preserve")
                        else:
                            nodes[0].attrib.pop(_XML_SPACE, None)
                        return
                _clear_cell_content(self._element)
        else:
            _clear_cell_content(self._element)
            if "t" in self._element.attrib:
                del self._element.attrib["t"]

        idx = sst.add(new_value)
        self._element.set("t", "s")
        _ensure_v(self._element).text = str(idx)

    def replace(self, old: str, new: str) -> int:
        """Replace substring in this cell's string content. Returns replacement count."""
        if not old or self.has_formula:
            return 0
        t = self.data_type
        if t == "s":
            idx_text = self._v_text()
            if idx_text is None:
                return 0
            sst = self._worksheet.shared_strings
            if sst is None:
                return 0
            return sst.replace_in_index(int(idx_text), old, new)
        if t == "inlineStr":
            is_elm = self._element.find(_IS)
            if is_elm is None:
                return 0
            total = 0
            for t_elm in _inline_t_nodes(is_elm):
                text = t_elm.text or ""
                if old not in text:
                    continue
                count = text.count(old)
                t_elm.text = text.replace(old, new)
                total += count
            return total
        return 0

"""Cell proxy — value get/set preserves the style index."""

from __future__ import annotations

import functools
from typing import TYPE_CHECKING, Any, Callable, TypeVar
from datetime import date, datetime

from lxml import etree
from lxml.etree import _Element

from xlsxedit.exceptions import FormulaGroupError, GridOverflowError, InvalidRangeError
from xlsxedit.merge import address_in_range, parse_range
from xlsxedit.opc.constants import SML_NS
from xlsxedit.oxml.address import (
    col_to_index,
    index_to_col,
    join_address,
    split_address,
)
from xlsxedit.shared_strings import si_display_text
from xlsxedit.styles import CellStyle, serial_to_datetime, datetime_to_serial
from xlsxedit.hyperlinks import Hyperlink

if TYPE_CHECKING:
    from xlsxedit.worksheet import Worksheet

_C = f"{{{SML_NS}}}c"
# Cell._shifts_seen once a shift has pushed the cell off the grid or deleted it
_PUSHED_OFF, _DELETED = -1, -2
_V = f"{{{SML_NS}}}v"
_F = f"{{{SML_NS}}}f"
_IS = f"{{{SML_NS}}}is"
_T = f"{{{SML_NS}}}t"
_R = f"{{{SML_NS}}}r"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_RANGE_FORMULA_TYPES = ("array", "dataTable")
_Method = TypeVar("_Method", bound=Callable[..., Any])
_RANGE_FORMULAS = etree.XPath(
    "m:sheetData/m:row/m:c/m:f[@t='array' or @t='dataTable']", namespaces={"m": SML_NS}
)
_SHARED_FORMULAS = etree.XPath(
    "m:sheetData/m:row/m:c/m:f[@t='shared' and @si=$si]", namespaces={"m": SML_NS}
)


def _marks_sheet(method: _Method) -> _Method:
    """Mark the worksheet's part dirty when ``method`` changes the cell's element.

    A call that leaves the element as it was, such as setting the value the
    cell already holds, marks nothing, so the sheet keeps its source bytes.
    A cell the sheet has no ``<c>`` for is written by :func:`_write_new_cell`.
    """

    @functools.wraps(method)
    def wrapper(self: Cell, *args: Any, **kwargs: Any) -> Any:
        part = self._worksheet._part
        if "_element" not in self.__dict__:
            self._element, undo = self._worksheet._cell_for_write(self._where())
            if undo is not None:
                return _write_new_cell(self, undo, method, args, kwargs)
        if part.is_dirty:
            return method(self, *args, **kwargs)
        before = etree.tostring(self._element, with_tail=False)
        try:
            return method(self, *args, **kwargs)
        finally:
            if etree.tostring(self._element, with_tail=False) != before:
                part.mark_dirty()

    return wrapper  # type: ignore[return-value]


def _write_new_cell(
    cell: Cell, undo: Callable[[], None], method: Callable[..., Any], args: tuple, kwargs: dict
) -> Any:
    """Run the write ``method`` on the element just created for ``cell``.

    ``undo`` takes the element back out, with its row when that is new too.
    It runs when the method raises or leaves the element as it was created,
    so a refused write, or one that writes nothing such as ``None``, leaves
    the sheet as it was.
    """
    element = cell._element
    created = element.items()
    try:
        result = method(cell, *args, **kwargs)
    except BaseException:
        undo()
        del cell._element
        raise
    if len(element) or element.text is not None or element.items() != created:
        cell._worksheet._part.mark_dirty()
    else:
        undo()
        del cell._element
    return result


def _date_or_number(numeric: int | float) -> Any:
    """A number in a date format as a ``datetime``, or as the number itself when no
    ``datetime`` can hold it, as for a serial past 9999-12-31, which Excel shows as ``#``."""
    try:
        return serial_to_datetime(float(numeric))
    except (OverflowError, ValueError):
        return numeric


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


def _shared_string_index(text: str | None, size: int) -> int | None:
    """The index ``text`` names in a table of ``size`` shared strings; ``None`` unless it
    names one in the plain form ``_sst_index_ref_count`` counts."""
    if text is None or not (text.isascii() and text.isdigit()):
        return None
    index = int(text)
    return index if index < size and str(index) == text else None


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
    """Thin proxy over a ``<c>`` element.

    ``ws[address]`` also returns one where the sheet has no ``<c>``, made with
    ``element=None`` and the ``address``. Reading it changes nothing: it reads
    as an empty cell until a write, through it or any other way, creates the
    element, which it then proxies. Until then it moves with the cells around
    it when rows or columns are inserted or cells shifted, as a cell's element
    does, and raises on use once a shift deletes it or pushes it off the grid.
    """

    def __init__(self, element: _Element | None, worksheet: Worksheet, address: str | None = None):
        if element is not None:
            self._element = element
        else:
            self._shifts_seen = len(worksheet._shifts)
        self._address = address
        self._worksheet = worksheet

    def __getattr__(self, name: str) -> Any:
        # Only a cell made without its element gets here for ``_element``.
        if name != "_element":
            raise AttributeError(name)
        element = self._bind()
        return etree.Element(_C, r=self._where()) if element is None else element

    def _where(self) -> str:
        """The address of a cell made without its element, moved by the shifts since.

        Replays the inserts and cell shifts applied to the sheet since the
        ``Cell`` was made (``Worksheet._shifts``). Raises ``GridOverflowError``
        once one has pushed the cell off the grid, and ``InvalidRangeError``
        once one has deleted it.
        """
        shifts = self._worksheet._shifts
        seen = self._shifts_seen
        if 0 <= seen < len(shifts):
            col, row = split_address(self._address)
            place: tuple[int, int] | None = (col_to_index(col) + 1, row)
            for shift in shifts[seen:]:
                place = shift.cell(*place)
                if place is None:
                    seen = _PUSHED_OFF if shift.inserting else _DELETED
                    break
            else:
                self._address = join_address(index_to_col(place[0] - 1), place[1])
                seen = len(shifts)
            self._shifts_seen = seen
        if seen == _PUSHED_OFF:
            raise GridOverflowError(f"an insert moved cell {self._address} past the grid")
        if seen == _DELETED:
            raise InvalidRangeError(f"cell {self._address} was deleted")
        return self._address

    def _bind(self) -> _Element | None:
        """The cell's ``<c>``, looked up while the proxy has none; ``None`` while the sheet has none."""
        element = self.__dict__.get("_element")
        if element is None:
            element = self._worksheet._find_cell_element(self._where())
            if element is not None:
                self._element = element
        return element

    def __repr__(self) -> str:
        element = self.__dict__.get("_element")
        if element is not None:
            return f"<Cell {element.get('r', '')!r}>"
        try:
            return f"<Cell {self._where()!r}>"
        except GridOverflowError:
            return f"<Cell {self._address!r}, moved past the grid>"
        except InvalidRangeError:
            return f"<Cell {self._address!r}, deleted>"

    @property
    def address(self) -> str:
        element = self.__dict__.get("_element")
        return self._where() if element is None else element.get("r", "")

    @property
    def worksheet(self) -> Worksheet:
        return self._worksheet

    def offset(self, cols: int = 0, rows: int = 0) -> Cell:
        """Return the cell ``cols`` columns right and ``rows`` rows down (negative allowed)."""
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
    @_marks_sheet
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
        new_f = etree.Element(_F, dict(f.attrib) if keep else {})
        new_f.text = text
        for child in list(self._element):
            if child.tag in (_V, _IS, _F):
                self._element.remove(child)
        self._element.attrib.pop("t", None)
        self._element.insert(0, new_f)
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
        """Read proxy for the cell's effective style (see ``Styles.effective_xf_index``)."""
        styles = self._worksheet._workbook.styles
        element = self._bind()
        if element is None:
            index = str(styles.effective_xf_index(self._worksheet, self._where()))
        else:
            index = element.get("s")
            if index is None:
                index = str(styles._cell_xf_index(element))
        return CellStyle(styles, index, self)

    @property
    def hyperlink(self) -> Hyperlink:
        return Hyperlink(self)

    def clear(self) -> None:
        """Clear cell content while preserving style index."""
        self.value = None

    @_marks_sheet
    def apply_date_format(self) -> None:
        """Apply a built-in date number format to this cell."""
        styles = self._worksheet._workbook.styles
        styles.apply_date_format(self._element)

    @_marks_sheet
    def apply_number_format(self, format_code: str) -> None:
        """Apply a custom number format string to this cell."""
        styles = self._worksheet._workbook.styles
        styles.apply_num_format(self._element, format_code)

    @_marks_sheet
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
        """Change only the passed font, fill, and alignment properties of the cell.

        Builds on the cell's effective style (see ``cell.style``): font changes
        keep the current font's other properties, and ``False`` turns bold,
        italic or underline off. An identical existing ``xf``, font or fill is
        reused instead of appending a duplicate.
        """
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
        if self._bind() is None:
            return None
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
                return _date_or_number(numeric)
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
            return _date_or_number(numeric)
        return numeric

    @value.setter
    @_marks_sheet
    def value(self, new_value: Any) -> None:
        """Set cell content while preserving ``s`` (style index).

        A value the cell cannot hold raises before anything changes.
        """
        from xlsxedit.row_shift import check_values_writable

        check_values_writable(
            "Cell.value",
            ((new_value,),),
            start_col_idx=0,
            string_columns=None,
            has_shared_strings=self._worksheet.shared_strings is not None,
        )
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
                v.text = repr(float(new_value))  # NumPy 2 scalars repr as np.float64(...)
            else:
                v.text = str(int(new_value))
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
            # An index that names no shared string is replaced like any other value.
            idx = _shared_string_index(self._v_text(), len(sst))
            if idx is not None:
                if self._worksheet._workbook._sst_index_ref_count(idx) <= 1:
                    from xlsxedit.shared_strings import _t_nodes_in_si

                    si = sst.get(idx)
                    nodes = _t_nodes_in_si(si)
                    if len(nodes) == 1:
                        before = etree.tostring(si, with_tail=False)
                        nodes[0].text = new_value
                        if " " in new_value or new_value.startswith(" ") or new_value.endswith(
                            " "
                        ):
                            nodes[0].set(_XML_SPACE, "preserve")
                        else:
                            nodes[0].attrib.pop(_XML_SPACE, None)
                        if etree.tostring(si, with_tail=False) != before:
                            sst.mark_dirty()
                        return
            _clear_cell_content(self._element)
        else:
            _clear_cell_content(self._element)
            if "t" in self._element.attrib:
                del self._element.attrib["t"]

        idx = sst.add(new_value)
        self._element.set("t", "s")
        _ensure_v(self._element).text = str(idx)

    @_marks_sheet
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

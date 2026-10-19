"""Bulk row/cell write helpers for large exports."""

from __future__ import annotations

from datetime import date, datetime
from typing import TYPE_CHECKING, Any

from lxml import etree

from xlsxedit.cell import _F, _clear_cell_content, _ensure_v, _refuse_range_formula_edit
from xlsxedit.opc.constants import SML_NS
from xlsxedit.styles import datetime_to_serial

if TYPE_CHECKING:
    from xlsxedit.shared_strings import SharedStringTable

_V = f"{{{SML_NS}}}v"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def write_cell_value(
    c_elm: etree._Element,
    value: Any,
    sst: SharedStringTable | None,
    sst_cache: dict[str, int],
    *,
    force_string: bool = False,
    worksheet=None,
) -> None:
    """Write a value directly to a ``<c>`` element (bulk path, no Cell proxy).

    A value the cell cannot hold, or, given a ``worksheet``, a cell that
    anchors a multi-cell array or data-table formula, raises before anything
    changes.
    """
    t: str | None = None
    text: str | None = None
    if value is None:
        pass
    elif force_string or isinstance(value, str):
        if sst is None:
            raise ValueError("workbook has no shared strings part")
        t, text = "s", str(value)
    elif isinstance(value, bool):
        t, text = "b", "1" if value else "0"
    elif isinstance(value, (datetime, date)):
        serial = datetime_to_serial(value)
        text = str(int(serial)) if serial == int(serial) else repr(serial)
    elif isinstance(value, (int, float)):
        text = repr(value) if isinstance(value, float) else str(value)
    else:
        raise TypeError(
            f"bulk write accepts str, int, float, bool, datetime, date, or None; got {type(value)!r}"
        )

    f_elm = None if worksheet is None else c_elm.find(_F)
    if f_elm is not None:
        _refuse_range_formula_edit(c_elm.get("r", ""), f_elm)
    if t == "s":
        text = str(_string_index(text, sst, sst_cache))
    if f_elm is not None:
        worksheet._workbook._on_formula_removed(worksheet, c_elm)

    _clear_cell_content(c_elm)
    if t is None:
        c_elm.attrib.pop("t", None)
    else:
        c_elm.set("t", t)
    if text is not None:
        _ensure_v(c_elm).text = text


def _string_index(text: str, sst: SharedStringTable, sst_cache: dict[str, int]) -> int:
    idx = sst_cache.get(text)
    if idx is None:
        idx = sst.add(text)
        sst_cache[text] = idx
    return idx

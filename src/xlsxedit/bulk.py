"""Bulk row/cell write helpers for large exports."""

from __future__ import annotations

from datetime import date, datetime
from typing import TYPE_CHECKING, Any

from lxml import etree

from xlsxedit.cell import _clear_cell_content, _ensure_v
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
    """Write a value directly to a ``<c>`` element (bulk path, no Cell proxy)."""
    if worksheet is not None:
        from xlsxedit.cell import _F

        if c_elm.find(_F) is not None:
            worksheet._workbook._on_formula_removed(worksheet, c_elm)

    if value is None:
        _clear_cell_content(c_elm)
        c_elm.attrib.pop("t", None)
        return

    if force_string:
        _write_string(c_elm, str(value), sst, sst_cache)
        return

    if isinstance(value, bool):
        _clear_cell_content(c_elm)
        c_elm.set("t", "b")
        _ensure_v(c_elm).text = "1" if value else "0"
        return

    if isinstance(value, (datetime, date)):
        _clear_cell_content(c_elm)
        c_elm.attrib.pop("t", None)
        serial = datetime_to_serial(value)
        _ensure_v(c_elm).text = str(int(serial)) if serial == int(serial) else repr(serial)
        return

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        _clear_cell_content(c_elm)
        c_elm.attrib.pop("t", None)
        v = _ensure_v(c_elm)
        v.text = repr(value) if isinstance(value, float) else str(value)
        return

    if isinstance(value, str):
        _write_string(c_elm, value, sst, sst_cache)
        return

    raise TypeError(
        f"bulk write accepts str, int, float, bool, datetime, date, or None; got {type(value)!r}"
    )


def _write_string(
    c_elm: etree._Element,
    text: str,
    sst: SharedStringTable | None,
    sst_cache: dict[str, int],
) -> None:
    if sst is None:
        raise ValueError("workbook has no shared strings part")
    _clear_cell_content(c_elm)
    if text in sst_cache:
        idx = sst_cache[text]
    else:
        idx = sst.add(text)
        sst_cache[text] = idx
    c_elm.set("t", "s")
    _ensure_v(c_elm).text = str(idx)

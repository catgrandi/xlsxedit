"""Style sheet helpers (``styles.xml``) — read and limited write."""

from __future__ import annotations

from copy import deepcopy
from datetime import date, datetime, timedelta
from typing import TYPE_CHECKING, Any

from lxml import etree

from xlsxedit._ooxml_order import ORDER
from xlsxedit.exceptions import InvalidColorError
from xlsxedit.opc.constants import SML_NS
from xlsxedit.worksheet_order import insert_ordered, reposition_ordered

if TYPE_CHECKING:
    from lxml.etree import _Element

_XF = f"{{{SML_NS}}}xf"
_NUM_FMT = f"{{{SML_NS}}}numFmt"
_NUM_FMTS = f"{{{SML_NS}}}numFmts"
_CELL_XFS = f"{{{SML_NS}}}cellXfs"
_FONTS = f"{{{SML_NS}}}fonts"
_FONT = f"{{{SML_NS}}}font"
_FILLS = f"{{{SML_NS}}}fills"
_FILL = f"{{{SML_NS}}}fill"
_DXFS = f"{{{SML_NS}}}dxfs"
_DXF = f"{{{SML_NS}}}dxf"
_ALIGNMENT = f"{{{SML_NS}}}alignment"
_SZ = f"{{{SML_NS}}}sz"
_COLOR = f"{{{SML_NS}}}color"
_NAME = f"{{{SML_NS}}}name"
_B = f"{{{SML_NS}}}b"
_I = f"{{{SML_NS}}}i"
_U = f"{{{SML_NS}}}u"
_FAMILY = f"{{{SML_NS}}}family"
_SCHEME = f"{{{SML_NS}}}scheme"
_PATTERN_FILL = f"{{{SML_NS}}}patternFill"
_FG_COLOR = f"{{{SML_NS}}}fgColor"
_BG_COLOR = f"{{{SML_NS}}}bgColor"

_STYLESHEET_ORDER = ORDER["CT_Stylesheet"]
_XF_ORDER = ORDER["CT_Xf"]

_BUILTIN_DATE_IDS = frozenset({14, 15, 16, 17, 22, 27, 28, 29, 30, 31, 36, 50, 52, 57})
_BUILTIN_PERCENT_IDS = frozenset({9, 10})
_BUILTIN_DATE_NUMFMT_ID = 14
_EXCEL_EPOCH = datetime(1899, 12, 30)
_MIN_CUSTOM_NUMFMT_ID = 164


def normalize_rgb(color: str) -> str:
    """Normalize ``RRGGBB`` or ``AARRGGBB`` to 8-char ``AARRGGBB`` uppercase."""
    c = color.strip().lstrip("#").upper()
    if len(c) == 6:
        return "FF" + c
    if len(c) == 8:
        return c
    raise InvalidColorError(f"invalid color: {color!r}")


def _format_code_indicates_date(code: str) -> bool:
    c = code.lower()
    if "yy" in c or "dd" in c or "mm" in c or "h" in c:
        return True
    return "[$-f800]" in c.lower() or "dddd" in c


class Styles:
    """View of ``styles.xml`` with optional write support."""

    def __init__(self, element: _Element | None, workbook=None):
        self._element = element
        self._workbook = workbook
        self._dirty = False
        self._num_fmts: dict[int, str] = {}
        self._cell_xfs: list[_Element] = []
        self._fonts: list[_Element] = []
        self._fills: list[_Element] = []
        if element is not None:
            num_fmts = element.find(_NUM_FMTS)
            if num_fmts is not None:
                for nf in num_fmts.findall(_NUM_FMT):
                    self._num_fmts[int(nf.get("numFmtId"))] = nf.get("formatCode", "")
            cell_xfs = element.find(_CELL_XFS)
            if cell_xfs is not None:
                self._cell_xfs = list(cell_xfs.findall(_XF))
            fonts = element.find(_FONTS)
            if fonts is not None:
                self._fonts = list(fonts.findall(_FONT))
            fills = element.find(_FILLS)
            if fills is not None:
                self._fills = list(fills.findall(_FILL))

    @property
    def dirty(self) -> bool:
        return self._dirty

    def _mark_dirty(self) -> None:
        self._dirty = True
        if self._workbook is not None:
            part = self._workbook._workbook_part.styles_part
            if part is not None and hasattr(part, "mark_dirty"):
                part.mark_dirty()

    def _xf(self, style_index: int | str | None) -> _Element | None:
        if style_index is None or self._element is None:
            return None
        idx = int(style_index)
        if idx < 0 or idx >= len(self._cell_xfs):
            return None
        return self._cell_xfs[idx]

    def _font(self, font_id: int | str | None) -> _Element | None:
        if font_id is None:
            return None
        idx = int(font_id)
        if idx < 0 or idx >= len(self._fonts):
            return None
        return self._fonts[idx]

    def _fill(self, fill_id: int | str | None) -> _Element | None:
        if fill_id is None:
            return None
        idx = int(fill_id)
        if idx < 0 or idx >= len(self._fills):
            return None
        return self._fills[idx]

    def num_format_id(self, style_index: int | str | None) -> int | None:
        xf = self._xf(style_index)
        if xf is None:
            return None
        return int(xf.get("numFmtId", "0"))

    def format_code(self, style_index: int | str | None) -> str | None:
        nf_id = self.num_format_id(style_index)
        if nf_id is None:
            return None
        if nf_id in self._num_fmts:
            return self._num_fmts[nf_id]
        return None

    def is_date_format(self, style_index: int | str | None) -> bool:
        nf_id = self.num_format_id(style_index)
        if nf_id is None:
            return False
        if nf_id in _BUILTIN_DATE_IDS:
            return True
        code = self.format_code(style_index)
        return bool(code and _format_code_indicates_date(code))

    def is_percent_format(self, style_index: int | str | None) -> bool:
        nf_id = self.num_format_id(style_index)
        if nf_id is None:
            return False
        if nf_id in _BUILTIN_PERCENT_IDS:
            return True
        code = self.format_code(style_index)
        return bool(code and "%" in code)

    def font_bold(self, style_index: int | str | None) -> bool:
        xf = self._xf(style_index)
        if xf is None:
            return False
        font = self._font(xf.get("fontId", "0"))
        return font is not None and font.find(_B) is not None

    def font_italic(self, style_index: int | str | None) -> bool:
        xf = self._xf(style_index)
        if xf is None:
            return False
        font = self._font(xf.get("fontId", "0"))
        return font is not None and font.find(_I) is not None

    def font_underline(self, style_index: int | str | None) -> bool:
        xf = self._xf(style_index)
        if xf is None:
            return False
        font = self._font(xf.get("fontId", "0"))
        return font is not None and font.find(_U) is not None

    def font_size(self, style_index: int | str | None) -> int | None:
        xf = self._xf(style_index)
        if xf is None:
            return None
        font = self._font(xf.get("fontId", "0"))
        if font is None:
            return None
        sz = font.find(_SZ)
        return int(sz.get("val")) if sz is not None else None

    def font_name(self, style_index: int | str | None) -> str | None:
        xf = self._xf(style_index)
        if xf is None:
            return None
        font = self._font(xf.get("fontId", "0"))
        if font is None:
            return None
        name = font.find(_NAME)
        return name.get("val") if name is not None else None

    def font_color_rgb(self, style_index: int | str | None) -> str | None:
        xf = self._xf(style_index)
        if xf is None:
            return None
        font = self._font(xf.get("fontId", "0"))
        if font is None:
            return None
        color = font.find(_COLOR)
        if color is None:
            return None
        return color.get("rgb") or color.get("theme")

    def fill_color_rgb(self, style_index: int | str | None) -> str | None:
        xf = self._xf(style_index)
        if xf is None:
            return None
        fill = self._fill(xf.get("fillId", "0"))
        if fill is None:
            return None
        pf = fill.find(_PATTERN_FILL)
        if pf is None:
            return None
        fg = pf.find(_FG_COLOR)
        return fg.get("rgb") if fg is not None else None

    def horizontal_align(self, style_index: int | str | None) -> str | None:
        xf = self._xf(style_index)
        if xf is None:
            return None
        al = xf.find(_ALIGNMENT)
        return al.get("horizontal") if al is not None else None

    def vertical_align(self, style_index: int | str | None) -> str | None:
        xf = self._xf(style_index)
        if xf is None:
            return None
        al = xf.find(_ALIGNMENT)
        return al.get("vertical") if al is not None else None

    def _cell_xfs_element(self) -> _Element:
        if self._element is None:
            raise ValueError("workbook has no styles part")
        cell_xfs = self._element.find(_CELL_XFS)
        if cell_xfs is None:
            cell_xfs = etree.Element(_CELL_XFS)
            cell_xfs.set("count", "0")
            insert_ordered(self._element, cell_xfs, _STYLESHEET_ORDER)
        return cell_xfs

    def _fonts_element(self) -> _Element:
        if self._element is None:
            raise ValueError("workbook has no styles part")
        fonts = self._element.find(_FONTS)
        if fonts is None:
            fonts = etree.Element(_FONTS)
            fills = self._element.find(_FILLS)
            if fills is not None:
                self._element.insert(list(self._element).index(fills), fonts)
            else:
                self._element.insert(0, fonts)
            fonts.set("count", "0")
        return fonts

    def _fills_element(self) -> _Element:
        if self._element is None:
            raise ValueError("workbook has no styles part")
        fills = self._element.find(_FILLS)
        if fills is None:
            fills = etree.Element(_FILLS)
            fills.set("count", "0")
            insert_ordered(self._element, fills, _STYLESHEET_ORDER)
        return fills

    def _dxfs_element(self) -> _Element:
        if self._element is None:
            raise ValueError("workbook has no styles part")
        dxfs = self._element.find(_DXFS)
        if dxfs is None:
            dxfs = etree.Element(_DXFS)
            dxfs.set("count", "0")
            insert_ordered(self._element, dxfs, _STYLESHEET_ORDER)
        return dxfs

    def _next_custom_num_fmt_id(self) -> int:
        used = set(self._num_fmts)
        for xf in self._cell_xfs:
            used.add(int(xf.get("numFmtId", "0")))
        n = _MIN_CUSTOM_NUMFMT_ID
        while n in used:
            n += 1
        return n

    def _fix_num_fmts_position(self, num_fmts: _Element) -> None:
        """OOXML requires ``numFmts`` first in the stylesheet; Excel rejects wrong order."""
        if reposition_ordered(self._element, num_fmts, _STYLESHEET_ORDER):
            self._mark_dirty()

    def _ensure_num_fmts_block(self) -> _Element:
        if self._element is None:
            raise ValueError("workbook has no styles part")
        num_fmts = self._element.find(_NUM_FMTS)
        if num_fmts is None:
            num_fmts = etree.Element(_NUM_FMTS)
            num_fmts.set("count", "0")
            insert_ordered(self._element, num_fmts, _STYLESHEET_ORDER)
        else:
            self._fix_num_fmts_position(num_fmts)
        return num_fmts

    def ensure_num_format(self, format_code: str) -> int:
        for nf_id, code in self._num_fmts.items():
            if code == format_code:
                return nf_id
        num_fmts = self._ensure_num_fmts_block()

        nf_id = self._next_custom_num_fmt_id()
        nf_elm = etree.SubElement(num_fmts, _NUM_FMT)
        nf_elm.set("numFmtId", str(nf_id))
        nf_elm.set("formatCode", format_code)
        num_fmts.set("count", str(len(num_fmts.findall(_NUM_FMT))))
        self._num_fmts[nf_id] = format_code
        self._mark_dirty()
        return nf_id

    def ensure_font(
        self,
        *,
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        size: int = 12,
        color_rgb: str | None = None,
        name: str | None = None,
    ) -> int:
        if not self._fonts:
            raise ValueError("styles part has no fonts")
        font = deepcopy(self._fonts[0])
        for tag in (_B, _I, _U):
            existing = font.find(tag)
            if existing is not None:
                font.remove(existing)
        if bold:
            font.insert(0, etree.Element(_B))
        if italic:
            font.insert(0, etree.Element(_I))
        if underline:
            font.insert(0, etree.Element(_U))

        sz = font.find(_SZ)
        if sz is None:
            sz = etree.SubElement(font, _SZ)
        sz.set("val", str(size))

        if color_rgb is not None:
            color = font.find(_COLOR)
            if color is None:
                color = etree.SubElement(font, _COLOR)
            color.set("rgb", normalize_rgb(color_rgb))
            color.attrib.pop("theme", None)

        if name is not None:
            name_elm = font.find(_NAME)
            if name_elm is None:
                name_elm = etree.SubElement(font, _NAME)
            name_elm.set("val", name)
            scheme = font.find(_SCHEME)
            if scheme is not None:
                font.remove(scheme)

        fonts_elm = self._fonts_element()
        fonts_elm.append(font)
        font_id = len(self._fonts)
        self._fonts.append(font)
        fonts_elm.set("count", str(len(self._fonts)))
        self._mark_dirty()
        return font_id

    def ensure_fill(self, color_rgb: str) -> int:
        fill = etree.Element(_FILL)
        pf = etree.SubElement(fill, _PATTERN_FILL)
        pf.set("patternType", "solid")
        fg = etree.SubElement(pf, _FG_COLOR)
        fg.set("rgb", normalize_rgb(color_rgb))
        bg = etree.SubElement(pf, _BG_COLOR)
        bg.set("indexed", "64")

        fills_elm = self._fills_element()
        fills_elm.append(fill)
        fill_id = len(self._fills)
        self._fills.append(fill)
        fills_elm.set("count", str(len(self._fills)))
        self._mark_dirty()
        return fill_id

    def ensure_dxf(
        self,
        *,
        font_color: str | None = None,
        bg_color: str | None = None,
        bold: bool = False,
    ) -> int:
        dxf = etree.Element(_DXF)
        if font_color is not None or bold:
            font = etree.SubElement(dxf, _FONT)
            if bold:
                etree.SubElement(font, _B)
            sz = etree.SubElement(font, _SZ)
            sz.set("val", "11")
            if font_color is not None:
                color = etree.SubElement(font, _COLOR)
                color.set("rgb", normalize_rgb(font_color))
        if bg_color is not None:
            fill = etree.SubElement(dxf, _FILL)
            pf = etree.SubElement(fill, _PATTERN_FILL)
            bg = etree.SubElement(pf, _BG_COLOR)
            bg.set("rgb", normalize_rgb(bg_color))

        dxfs_elm = self._dxfs_element()
        dxfs_elm.append(dxf)
        dxf_id = len(dxfs_elm.findall(_DXF)) - 1
        dxfs_elm.set("count", str(len(dxfs_elm.findall(_DXF))))
        self._mark_dirty()
        return dxf_id

    def clone_xf(self, source_index: int | None = None, **overrides: Any) -> int:
        cell_xfs = self._cell_xfs_element()
        if source_index is None:
            source_index = 0
        if source_index < 0 or source_index >= len(self._cell_xfs):
            raise IndexError(f"style index out of range: {source_index}")
        new_xf = deepcopy(self._cell_xfs[source_index])

        if "numFmtId" in overrides:
            new_xf.set("numFmtId", str(overrides["numFmtId"]))
            new_xf.set("applyNumberFormat", "1")
        if "fontId" in overrides:
            new_xf.set("fontId", str(overrides["fontId"]))
            new_xf.set("applyFont", "1")
        if "fillId" in overrides:
            new_xf.set("fillId", str(overrides["fillId"]))
            new_xf.set("applyFill", "1")
        if "alignment" in overrides:
            al = overrides["alignment"]
            al_elm = new_xf.find(_ALIGNMENT)
            if al_elm is None:
                al_elm = etree.Element(_ALIGNMENT)
                insert_ordered(new_xf, al_elm, _XF_ORDER)
            for key, val in al.items():
                if val is not None:
                    al_elm.set(key, val)
            new_xf.set("applyAlignment", "1")

        cell_xfs.append(new_xf)
        new_index = len(self._cell_xfs)
        self._cell_xfs.append(new_xf)
        cell_xfs.set("count", str(len(self._cell_xfs)))
        self._mark_dirty()
        return new_index

    def apply_cell_style(self, cell_element: _Element, **style_kwargs: Any) -> None:
        source = int(cell_element.get("s", "0"))
        overrides = self._style_kwargs_to_overrides(style_kwargs)
        if not overrides:
            return
        new_idx = self.clone_xf(source, **overrides)
        cell_element.set("s", str(new_idx))

    def _style_kwargs_to_overrides(self, style_kwargs: dict[str, Any]) -> dict[str, Any]:
        overrides: dict[str, Any] = {}

        if style_kwargs.get("num_format") is not None:
            overrides["numFmtId"] = self.ensure_num_format(style_kwargs["num_format"])

        font_keys = {
            k: style_kwargs[k]
            for k in ("bold", "italic", "underline", "font_size", "font_color", "font_name")
            if k in style_kwargs and style_kwargs[k] is not None
        }
        if font_keys:
            overrides["fontId"] = self.ensure_font(
                bold=bool(font_keys.get("bold")),
                italic=bool(font_keys.get("italic")),
                underline=bool(font_keys.get("underline")),
                size=int(font_keys.get("font_size", 12)),
                color_rgb=font_keys.get("font_color"),
                name=font_keys.get("font_name"),
            )

        if style_kwargs.get("bg_color") is not None:
            overrides["fillId"] = self.ensure_fill(style_kwargs["bg_color"])

        align = {}
        if style_kwargs.get("horizontal_align") is not None:
            align["horizontal"] = style_kwargs["horizontal_align"]
        if style_kwargs.get("vertical_align") is not None:
            align["vertical"] = style_kwargs["vertical_align"]
        if align:
            overrides["alignment"] = align

        return overrides

    def allocate_cell_style(self, *, base_xf: int = 0, **style_kwargs: Any) -> int:
        """Create a cellXf index without touching a worksheet cell."""
        overrides = self._style_kwargs_to_overrides(style_kwargs)
        if not overrides:
            return base_xf
        return self.clone_xf(base_xf, **overrides)

    def apply_num_format(self, cell_element: _Element, format_code: str) -> None:
        nf_id = self.ensure_num_format(format_code)
        source = int(cell_element.get("s", "0"))
        new_idx = self.clone_xf(source, numFmtId=nf_id)
        cell_element.set("s", str(new_idx))

    def apply_date_format(self, cell_element: _Element) -> None:
        if self.is_date_format(cell_element.get("s")):
            return
        source = int(cell_element.get("s", "0"))
        new_idx = self.clone_xf(source, numFmtId=_BUILTIN_DATE_NUMFMT_ID)
        cell_element.set("s", str(new_idx))

    def set_cell_style_index(self, cell_element: _Element, style_index: int) -> None:
        cell_element.set("s", str(style_index))


def serial_to_datetime(serial: float, *, date1904: bool = False) -> datetime:
    epoch = datetime(1904, 1, 1) if date1904 else _EXCEL_EPOCH
    days = int(serial)
    frac = serial - days
    dt = epoch + timedelta(days=days, seconds=round(frac * 86400))
    return dt


def datetime_to_serial(value: datetime | date, *, date1904: bool = False) -> float:
    epoch = datetime(1904, 1, 1) if date1904 else _EXCEL_EPOCH
    if isinstance(value, date) and not isinstance(value, datetime):
        value = datetime(value.year, value.month, value.day)
    delta = value - epoch
    return delta.days + (delta.seconds + delta.microseconds / 1_000_000) / 86400


class CellStyle:
    """Cell style proxy."""

    def __init__(self, styles: Styles | None, style_index: str | None, cell_element=None):
        self._styles = styles
        self._index = style_index
        self._cell_element = cell_element

    @property
    def style_index(self) -> str | None:
        return self._index

    @property
    def num_format_id(self) -> int | None:
        if self._styles is None:
            return None
        return self._styles.num_format_id(self._index)

    @property
    def num_format(self) -> str | None:
        if self._styles is None:
            return None
        return self._styles.format_code(self._index)

    @num_format.setter
    def num_format(self, format_code: str) -> None:
        if self._styles is None or self._cell_element is None:
            raise ValueError("cannot set number format without styles part")
        self._styles.apply_num_format(self._cell_element, format_code)
        self._index = self._cell_element.get("s")

    @property
    def is_date(self) -> bool:
        if self._styles is None:
            return False
        return self._styles.is_date_format(self._index)

    @property
    def is_percent(self) -> bool:
        if self._styles is None:
            return False
        return self._styles.is_percent_format(self._index)

    @property
    def bold(self) -> bool:
        return self._styles.font_bold(self._index) if self._styles else False

    @property
    def italic(self) -> bool:
        return self._styles.font_italic(self._index) if self._styles else False

    @property
    def underline(self) -> bool:
        return self._styles.font_underline(self._index) if self._styles else False

    @property
    def font_size(self) -> int | None:
        return self._styles.font_size(self._index) if self._styles else None

    @property
    def font_name(self) -> str | None:
        return self._styles.font_name(self._index) if self._styles else None

    @property
    def font_color(self) -> str | None:
        return self._styles.font_color_rgb(self._index) if self._styles else None

    @property
    def bg_color(self) -> str | None:
        return self._styles.fill_color_rgb(self._index) if self._styles else None

    @property
    def horizontal_align(self) -> str | None:
        return self._styles.horizontal_align(self._index) if self._styles else None

    @property
    def vertical_align(self) -> str | None:
        return self._styles.vertical_align(self._index) if self._styles else None

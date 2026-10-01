"""Style sheet helpers (``styles.xml``) — read and limited write."""

from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import TYPE_CHECKING, Any

from lxml import etree

from xlsxedit._ooxml_order import ORDER
from xlsxedit.exceptions import InvalidColorError
from xlsxedit.opc.constants import SML_NS
from xlsxedit.oxml.address import col_to_index, join_address, split_address
from xlsxedit.worksheet_order import insert_ordered, reposition_ordered

if TYPE_CHECKING:
    from lxml.etree import _Element

    from xlsxedit.worksheet import Worksheet

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
_COLS = f"{{{SML_NS}}}cols"
_COL = f"{{{SML_NS}}}col"
_SHEET_DATA = f"{{{SML_NS}}}sheetData"

_STYLESHEET_ORDER = ORDER["CT_Stylesheet"]
_XF_ORDER = ORDER["CT_Xf"]
# CT_Font is an unbounded xsd:choice, so the schema fixes no child order; this
# is the order Excel writes.
_FONT_ORDER = (
    "b",
    "i",
    "strike",
    "condense",
    "extend",
    "outline",
    "shadow",
    "u",
    "vertAlign",
    "sz",
    "color",
    "name",
    "family",
    "charset",
    "scheme",
)

_TRUE = frozenset({"1", "true"})
# CT_BooleanProperty val="0"/"false" and CT_UnderlineProperty val="none" switch the property off.
_OFF = frozenset({"0", "false", "none"})

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


@dataclass(frozen=True, eq=False)
class Color:
    """A colour reference from ``styles.xml`` (``CT_Color``).

    One of ``rgb`` (``AARRGGBB``), ``theme`` (index into the theme palette) or
    ``indexed`` (legacy palette index) is set, or ``auto`` is true. ``tint``
    lightens (positive) or darkens (negative) that base colour. A colour equals
    a string that is its ``rgb`` value, so ``font_color == "FFFF0000"`` works.
    """

    rgb: str | None = None
    theme: int | None = None
    tint: float = 0.0
    indexed: int | None = None
    auto: bool = False

    def _fields(self) -> tuple[Any, ...]:
        return (self.rgb, self.theme, self.tint, self.indexed, self.auto)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, str):
            return self.rgb == other
        if isinstance(other, Color):
            return self._fields() == other._fields()
        return NotImplemented

    def __hash__(self) -> int:
        return hash(self.rgb) if self.rgb is not None else hash(self._fields())

    @classmethod
    def from_element(cls, elm: _Element) -> Color:
        theme = elm.get("theme")
        indexed = elm.get("indexed")
        return cls(
            rgb=elm.get("rgb"),
            theme=int(theme) if theme is not None else None,
            tint=float(elm.get("tint", "0")),
            indexed=int(indexed) if indexed is not None else None,
            auto=elm.get("auto") in _TRUE,
        )


def _canonical(elm: _Element) -> tuple:
    """Hashable form of ``elm`` that ignores attribute order, prefixes and whitespace."""
    return (
        elm.tag,
        tuple(sorted(elm.attrib.items())),
        (elm.text or "").strip(),
        tuple(_canonical(child) for child in elm if isinstance(child.tag, str)),
    )


class _Pool:
    """Entries of one stylesheet collection with a canonical-form index.

    The index is seeded from the entries already in the collection and only
    grows: existing entries keep their position, and a duplicate already in the
    file resolves to its first occurrence.
    """

    def __init__(self, container: _Element | None, tag: str):
        self.tag = tag
        self.load(container)

    def load(self, container: _Element | None) -> None:
        self.container = container
        self.entries: list[_Element] = [] if container is None else container.findall(self.tag)
        self._positions: dict[tuple, int] = {}
        for i, entry in enumerate(self.entries):
            self._positions.setdefault(_canonical(entry), i)

    def add(self, candidate: _Element) -> tuple[int, bool]:
        """Return the position of an entry equal to ``candidate``, appending it if new."""
        key = _canonical(candidate)
        pos = self._positions.get(key)
        # An entry edited in place after indexing no longer matches its key.
        if pos is not None and _canonical(self.entries[pos]) == key:
            return pos, False
        self.container.append(candidate)
        pos = len(self.entries)
        self.entries.append(candidate)
        self._positions[key] = pos
        self.container.set("count", str(len(self.entries)))
        return pos, True


def _flag(font: _Element, tag: str) -> bool:
    elm = font.find(tag)
    return elm is not None and elm.get("val") not in _OFF


def _font_child(font: _Element, tag: str) -> _Element:
    elm = font.find(tag)
    if elm is None:
        elm = etree.Element(tag)
        insert_ordered(font, elm, _FONT_ORDER)
    return elm


def _set_flag(font: _Element, tag: str, value: bool | None) -> None:
    if value is None:
        return
    elm = font.find(tag)
    if not value:
        if elm is not None:
            font.remove(elm)
    elif elm is None:
        insert_ordered(font, etree.Element(tag), _FONT_ORDER)
    elif elm.get("val") in _OFF:
        del elm.attrib["val"]


def _format_number(value: float) -> str:
    number = float(value)
    return str(int(number)) if number.is_integer() else repr(number)


def _parse_number(text: str) -> int | float:
    number = float(text)
    return int(number) if number.is_integer() else number


class Styles:
    """View of ``styles.xml`` with optional write support.

    ``clone_xf``, ``ensure_font``, ``ensure_fill`` and ``ensure_dxf`` return the
    index of an identical existing entry instead of appending a duplicate.
    Existing entries are never changed or renumbered.
    """

    def __init__(self, element: _Element | None, workbook=None):
        self._element = element
        self._workbook = workbook
        self._dirty = False
        self._num_fmts: dict[int, str] = {}

        def pool(container_tag: str, tag: str) -> _Pool:
            container = element.find(container_tag) if element is not None else None
            return _Pool(container, tag)

        self._xf_pool = pool(_CELL_XFS, _XF)
        self._font_pool = pool(_FONTS, _FONT)
        self._fill_pool = pool(_FILLS, _FILL)
        self._dxf_pool = pool(_DXFS, _DXF)
        if element is not None:
            num_fmts = element.find(_NUM_FMTS)
            if num_fmts is not None:
                for nf in num_fmts.findall(_NUM_FMT):
                    self._num_fmts[int(nf.get("numFmtId"))] = nf.get("formatCode", "")

    @property
    def dirty(self) -> bool:
        return self._dirty

    @property
    def _cell_xfs(self) -> list[_Element]:
        return self._xf_pool.entries

    @property
    def _fonts(self) -> list[_Element]:
        return self._font_pool.entries

    @property
    def _fills(self) -> list[_Element]:
        return self._fill_pool.entries

    def _mark_dirty(self) -> None:
        self._dirty = True
        if self._workbook is not None:
            part = self._workbook._workbook_part.styles_part
            if part is not None and hasattr(part, "mark_dirty"):
                part.mark_dirty()

    def _synced(self, pool: _Pool, container: _Element) -> _Pool:
        """``pool``, reloaded if its collection element was replaced since it was read."""
        if pool.container is not container:
            pool.load(container)
        return pool

    def _add(self, pool: _Pool, container: _Element, candidate: _Element) -> int:
        index, added = self._synced(pool, container).add(candidate)
        if added:
            self._mark_dirty()
        return index

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

    def _xf_font(self, style_index: int | str | None) -> _Element | None:
        xf = self._xf(style_index)
        if xf is None:
            return None
        return self._font(xf.get("fontId", "0"))

    def effective_xf_index(self, worksheet: Worksheet, address: str) -> int:
        """Return the ``cellXfs`` index that formats the cell at ``address``.

        The cell's own ``s`` wins, and a cell with a value but no ``s`` is in
        style 0, as Excel shows it. A cell with no ``<c>`` element, or an empty
        one such as ``ws[address]`` creates, takes the row's ``s`` when the row
        sets ``customFormat``, then the ``style`` of the ``<col>`` covering
        its column, then 0.
        """
        col, row_num = split_address(address)
        cell = worksheet._find_cell_element(join_address(col, row_num))
        row = cell.getparent() if cell is not None else worksheet._find_row_element(row_num)
        sheet_data = worksheet._part.element.find(_SHEET_DATA)
        xf = self.resolve_xf_index(sheet_data, row, cell, col_to_index(col) + 1)
        return 0 if xf is None else xf

    def _cell_xf_index(self, cell_element: _Element) -> int:
        """:meth:`effective_xf_index` for a ``<c>`` element already in hand."""
        s = cell_element.get("s")
        if s is not None:
            return int(s)
        if len(cell_element):
            return 0
        row = cell_element.getparent()
        sheet_data = row.getparent() if row is not None else None
        xf = self.resolve_xf_index(sheet_data, row, cell_element)
        return 0 if xf is None else xf

    def resolve_xf_index(
        self,
        sheet_data: _Element | None,
        row: _Element | None,
        cell: _Element | None,
        column: int | None = None,
    ) -> int | None:
        """Return the ``cellXfs`` index a cell takes from itself, its row or its column.

        The element-level form of :meth:`effective_xf_index`. ``sheet_data`` is
        the worksheet's ``<sheetData>``, ``row`` and ``cell`` the ``<row>`` and
        ``<c>`` elements (either may be absent) and ``column`` the 1-based
        column number, read from the cell's ``r`` when omitted. Returns
        ``None`` when nothing sets a style, including for a cell with a value
        but no ``s``.
        """
        if cell is not None:
            s = cell.get("s")
            if s is not None:
                return int(s)
            if len(cell):
                return None
        if row is not None and row.get("customFormat") in _TRUE:
            s = row.get("s")
            if s is not None:
                return int(s)
        if sheet_data is None:
            return None
        if column is None:
            address = cell.get("r") if cell is not None else None
            if not address:
                return None
            column = col_to_index(split_address(address)[0]) + 1
        # CT_Worksheet places every <cols> before <sheetData>.
        for cols in sheet_data.itersiblings(_COLS, preceding=True):
            for col in cols.iterchildren(_COL):
                style = col.get("style")
                if style is None:
                    continue
                if int(col.get("min", "0")) <= column <= int(col.get("max", "0")):
                    return int(style)
        return None

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
        code = self._num_fmts.get(nf_id)
        return bool(code and _format_code_indicates_date(code))

    def is_percent_format(self, style_index: int | str | None) -> bool:
        nf_id = self.num_format_id(style_index)
        if nf_id is None:
            return False
        if nf_id in _BUILTIN_PERCENT_IDS:
            return True
        code = self._num_fmts.get(nf_id)
        return bool(code and "%" in code)

    def font_bold(self, style_index: int | str | None) -> bool:
        font = self._xf_font(style_index)
        return font is not None and _flag(font, _B)

    def font_italic(self, style_index: int | str | None) -> bool:
        font = self._xf_font(style_index)
        return font is not None and _flag(font, _I)

    def font_underline(self, style_index: int | str | None) -> bool:
        font = self._xf_font(style_index)
        return font is not None and _flag(font, _U)

    def font_size(self, style_index: int | str | None) -> int | float | None:
        font = self._xf_font(style_index)
        if font is None:
            return None
        sz = font.find(_SZ)
        return _parse_number(sz.get("val")) if sz is not None else None

    def font_name(self, style_index: int | str | None) -> str | None:
        font = self._xf_font(style_index)
        if font is None:
            return None
        name = font.find(_NAME)
        return name.get("val") if name is not None else None

    def font_color(self, style_index: int | str | None) -> Color | None:
        font = self._xf_font(style_index)
        if font is None:
            return None
        color = font.find(_COLOR)
        return Color.from_element(color) if color is not None else None

    def font_color_rgb(self, style_index: int | str | None) -> str | None:
        color = self.font_color(style_index)
        return color.rgb if color is not None else None

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

    def _collection_element(self, tag: str) -> _Element:
        if self._element is None:
            raise ValueError("workbook has no styles part")
        container = self._element.find(tag)
        if container is None:
            container = etree.Element(tag)
            container.set("count", "0")
            insert_ordered(self._element, container, _STYLESHEET_ORDER)
        return container

    def _cell_xfs_element(self) -> _Element:
        return self._collection_element(_CELL_XFS)

    def _fills_element(self) -> _Element:
        return self._collection_element(_FILLS)

    def _dxfs_element(self) -> _Element:
        return self._collection_element(_DXFS)

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
        base_font: int = 0,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        size: float | None = None,
        color_rgb: str | None = None,
        name: str | None = None,
    ) -> int:
        """Return the id of font ``base_font`` with the given properties changed.

        ``None`` keeps a property as ``base_font`` has it; ``False`` removes
        bold, italic or underline. Setting ``name`` drops the theme font
        ``scheme``, which would otherwise override it. Reuses an identical
        existing font instead of adding one.
        """
        fonts_elm = self._element.find(_FONTS) if self._element is not None else None
        if fonts_elm is None or fonts_elm.find(_FONT) is None:
            raise ValueError("styles part has no fonts")
        entries = self._synced(self._font_pool, fonts_elm).entries
        base = self._font(base_font)
        font = deepcopy(base if base is not None else entries[0])

        _set_flag(font, _B, bold)
        _set_flag(font, _I, italic)
        _set_flag(font, _U, underline)
        if size is not None:
            _font_child(font, _SZ).set("val", _format_number(size))
        if color_rgb is not None:
            color = _font_child(font, _COLOR)
            for attr in ("auto", "indexed", "theme", "tint"):
                color.attrib.pop(attr, None)
            color.set("rgb", normalize_rgb(color_rgb))
        if name is not None:
            _font_child(font, _NAME).set("val", name)
            scheme = font.find(_SCHEME)
            if scheme is not None:
                font.remove(scheme)

        return self._add(self._font_pool, fonts_elm, font)

    def ensure_fill(self, color_rgb: str) -> int:
        """Return the id of a solid fill in ``color_rgb``, reusing an identical one."""
        fill = etree.Element(_FILL)
        pf = etree.SubElement(fill, _PATTERN_FILL)
        pf.set("patternType", "solid")
        fg = etree.SubElement(pf, _FG_COLOR)
        fg.set("rgb", normalize_rgb(color_rgb))
        bg = etree.SubElement(pf, _BG_COLOR)
        bg.set("indexed", "64")
        return self._add(self._fill_pool, self._fills_element(), fill)

    def ensure_dxf(
        self,
        *,
        font_color: str | None = None,
        bg_color: str | None = None,
        bold: bool = False,
    ) -> int:
        """Return the id of a differential format, reusing an identical one."""
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
        return self._add(self._dxf_pool, self._dxfs_element(), dxf)

    def clone_xf(self, source_index: int | None = None, **overrides: Any) -> int:
        """Return the index of xf ``source_index`` with ``overrides`` applied.

        Reuses an identical existing xf (possibly ``source_index`` itself)
        instead of adding one.
        """
        cell_xfs = self._cell_xfs_element()
        entries = self._synced(self._xf_pool, cell_xfs).entries
        if source_index is None:
            source_index = 0
        if source_index < 0 or source_index >= len(entries):
            raise IndexError(f"style index out of range: {source_index}")
        new_xf = deepcopy(entries[source_index])

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

        return self._add(self._xf_pool, cell_xfs, new_xf)

    def apply_cell_style(self, cell_element: _Element, **style_kwargs: Any) -> None:
        """Point the cell at an xf built on its effective style plus ``style_kwargs``."""
        source = self._cell_xf_index(cell_element)
        overrides = self._style_kwargs_to_overrides(style_kwargs, base_xf=source)
        if not overrides:
            return
        new_idx = self.clone_xf(source, **overrides)
        cell_element.set("s", str(new_idx))

    def _style_kwargs_to_overrides(
        self, style_kwargs: dict[str, Any], *, base_xf: int = 0
    ) -> dict[str, Any]:
        overrides: dict[str, Any] = {}

        if style_kwargs.get("num_format") is not None:
            overrides["numFmtId"] = self.ensure_num_format(style_kwargs["num_format"])

        def flag(key: str) -> bool | None:
            value = style_kwargs.get(key)
            return None if value is None else bool(value)

        font_keys = ("bold", "italic", "underline", "font_size", "font_color", "font_name")
        if any(style_kwargs.get(k) is not None for k in font_keys):
            xf = self._xf(base_xf)
            overrides["fontId"] = self.ensure_font(
                base_font=int(xf.get("fontId", "0")) if xf is not None else 0,
                bold=flag("bold"),
                italic=flag("italic"),
                underline=flag("underline"),
                size=style_kwargs.get("font_size"),
                color_rgb=style_kwargs.get("font_color"),
                name=style_kwargs.get("font_name"),
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
        """Return a cellXf index for ``base_xf`` plus ``style_kwargs`` without touching a cell."""
        overrides = self._style_kwargs_to_overrides(style_kwargs, base_xf=base_xf)
        if not overrides:
            return base_xf
        return self.clone_xf(base_xf, **overrides)

    def apply_num_format(self, cell_element: _Element, format_code: str) -> None:
        nf_id = self.ensure_num_format(format_code)
        source = self._cell_xf_index(cell_element)
        new_idx = self.clone_xf(source, numFmtId=nf_id)
        cell_element.set("s", str(new_idx))

    def apply_date_format(self, cell_element: _Element) -> None:
        source = self._cell_xf_index(cell_element)
        if self.is_date_format(source):
            if cell_element.get("s") is None:
                cell_element.set("s", str(source))
            return
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
    def font_size(self) -> int | float | None:
        return self._styles.font_size(self._index) if self._styles else None

    @property
    def font_name(self) -> str | None:
        return self._styles.font_name(self._index) if self._styles else None

    @property
    def font_color(self) -> Color | None:
        return self._styles.font_color(self._index) if self._styles else None

    @property
    def bg_color(self) -> str | None:
        return self._styles.fill_color_rgb(self._index) if self._styles else None

    @property
    def horizontal_align(self) -> str | None:
        return self._styles.horizontal_align(self._index) if self._styles else None

    @property
    def vertical_align(self) -> str | None:
        return self._styles.vertical_align(self._index) if self._styles else None

"""Drawing / picture proxies for worksheet images."""

from __future__ import annotations

import math
import mimetypes
from copy import deepcopy
from pathlib import Path
from typing import TYPE_CHECKING, Iterator

from lxml import etree
from lxml.etree import _Element

from xlsxedit._ooxml_order import ORDER
from xlsxedit.merge import parse_range

from xlsxedit.opc.constants import CT, OFFICE_REL_NS, RT, SML_NS
from xlsxedit.opc.packuri import PackURI
from xlsxedit.opc.part import Part
from xlsxedit.oxml.address import col_to_index, index_to_col, join_address, split_address
from xlsxedit.oxml.parser import parse_xml, serialize_xml
from xlsxedit.parts import ChartPart, DrawingPart
from xlsxedit.worksheet_order import insert_ordered, insert_worksheet_child

if TYPE_CHECKING:
    from xlsxedit.worksheet import Worksheet

XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"

_XDR_WSDR = f"{{{XDR_NS}}}wsDr"
_XDR_TWO_CELL = f"{{{XDR_NS}}}twoCellAnchor"
_XDR_ONE_CELL = f"{{{XDR_NS}}}oneCellAnchor"
_XDR_FROM = f"{{{XDR_NS}}}from"
_XDR_TO = f"{{{XDR_NS}}}to"
_XDR_COL = f"{{{XDR_NS}}}col"
_XDR_ROW = f"{{{XDR_NS}}}row"
_XDR_COLOFF = f"{{{XDR_NS}}}colOff"
_XDR_ROWOFF = f"{{{XDR_NS}}}rowOff"
_XDR_PIC = f"{{{XDR_NS}}}pic"
_XDR_CNVPR = f"{{{XDR_NS}}}cNvPr"
_XDR_BLIP_FILL = f"{{{XDR_NS}}}blipFill"
_XDR_SPPR = f"{{{XDR_NS}}}spPr"
_XDR_GRAPHIC_FRAME = f"{{{XDR_NS}}}graphicFrame"
_XDR_EXT = f"{{{XDR_NS}}}ext"
_A_BLIP = f"{{{A_NS}}}blip"
_A_OFF = f"{{{A_NS}}}off"
_A_EXT = f"{{{A_NS}}}ext"
_A_GRAPHIC = f"{{{A_NS}}}graphic"
_A_GRAPHIC_DATA = f"{{{A_NS}}}graphicData"
_C_CHART = f"{{{CHART_NS}}}chart"
_C_TITLE = f"{{{CHART_NS}}}title"
_C_TX = f"{{{CHART_NS}}}tx"
_C_V = f"{{{CHART_NS}}}v"
_C_F = f"{{{CHART_NS}}}f"
_C_STR_REF = f"{{{CHART_NS}}}strRef"
_C_NUM_REF = f"{{{CHART_NS}}}numRef"
_C_STR_CACHE = f"{{{CHART_NS}}}strCache"
_C_NUM_CACHE = f"{{{CHART_NS}}}numCache"
_C_RICH = f"{{{CHART_NS}}}rich"
_C_PT_COUNT = f"{{{CHART_NS}}}ptCount"
_C_PT = f"{{{CHART_NS}}}pt"
_C_OVERLAY = f"{{{CHART_NS}}}overlay"
_A_BODY_PR = f"{{{A_NS}}}bodyPr"
_A_LST_STYLE = f"{{{A_NS}}}lstStyle"
_A_P = f"{{{A_NS}}}p"
_A_P_PR = f"{{{A_NS}}}pPr"
_A_DEF_R_PR = f"{{{A_NS}}}defRPr"
_A_R = f"{{{A_NS}}}r"
_A_R_PR = f"{{{A_NS}}}rPr"
_A_T = f"{{{A_NS}}}t"
_A_END_PARA_R_PR = f"{{{A_NS}}}endParaRPr"
_C_TX_PR = f"{{{CHART_NS}}}txPr"
_AUTO_FILTER = f"{{{SML_NS}}}autoFilter"
_TABLE_COLUMNS = f"{{{SML_NS}}}tableColumns"
_TABLE_COLUMN = f"{{{SML_NS}}}tableColumn"
_WS_DRAWING = f"{{{SML_NS}}}drawing"
_TABLE_PARTS = f"{{{SML_NS}}}tableParts"
_TABLE_PART = f"{{{SML_NS}}}tablePart"
_A_XFRM = f"{{{A_NS}}}xfrm"
_SHEET_DATA = f"{{{SML_NS}}}sheetData"
_SHEET_FORMAT_PR = f"{{{SML_NS}}}sheetFormatPr"
_COLS = f"{{{SML_NS}}}cols"
_COL = f"{{{SML_NS}}}col"
_ROW = f"{{{SML_NS}}}row"

EMU_PER_INCH = 914400
EMU_PER_PIXEL = 9525  # 96 dpi approximation
_EMU_PER_POINT = 12700
_MAX_COL = 16_384
_MAX_ROW = 1_048_576
# Maximum digit width, in pixels, of Excel's default 11-point body font (Calibri, Aptos).
_MAX_DIGIT_WIDTH_PX = 7


def _px_to_emu(pixels: int) -> int:
    return int(pixels * EMU_PER_PIXEL)


def _emu_to_px(emu: int) -> int:
    return max(0, round(emu / EMU_PER_PIXEL))


def _validate_offset_px(pixels: int) -> int:
    if not isinstance(pixels, int) or isinstance(pixels, bool):
        raise TypeError("offset must be an int")
    if pixels < 0:
        raise ValueError("offset must be non-negative")
    return pixels


def _validate_size_px(pixels: int | float) -> int | float:
    if not isinstance(pixels, (int, float)) or isinstance(pixels, bool):
        raise TypeError("size must be a number of pixels")
    if pixels <= 0:
        raise ValueError("size must be positive")
    return pixels


def _get_from_offset_emu(anchor: _Element) -> tuple[int, int]:
    from_elm = anchor.find(_XDR_FROM)
    if from_elm is None:
        return 0, 0
    return (
        int(from_elm.findtext(_XDR_COLOFF, "0") or "0"),
        int(from_elm.findtext(_XDR_ROWOFF, "0") or "0"),
    )


def _set_corner_offset_emu(corner: _Element, col_off: int, row_off: int) -> None:
    col_off_elm = corner.find(_XDR_COLOFF)
    row_off_elm = corner.find(_XDR_ROWOFF)
    if col_off_elm is not None:
        col_off_elm.text = str(col_off)
    if row_off_elm is not None:
        row_off_elm.text = str(row_off)


def _set_from_offset_emu(anchor: _Element, col_off: int, row_off: int) -> None:
    from_elm = anchor.find(_XDR_FROM)
    if from_elm is not None:
        _set_corner_offset_emu(from_elm, col_off, row_off)


def _marker(marker: _Element) -> tuple[int, int, int, int]:
    """``(col, colOff, row, rowOff)`` of an ``xdr:from`` or ``xdr:to`` marker."""
    col, col_off, row, row_off = (
        int(marker.findtext(tag) or "0") for tag in (_XDR_COL, _XDR_COLOFF, _XDR_ROW, _XDR_ROWOFF)
    )
    return col, col_off, row, row_off


def _set_marker(marker: _Element, col: int, col_off: int, row: int, row_off: int) -> None:
    for tag, value in ((_XDR_COL, col), (_XDR_COLOFF, col_off), (_XDR_ROW, row), (_XDR_ROWOFF, row_off)):
        child = marker.find(tag)
        if child is None:
            raise ValueError(
                f"malformed anchor: xdr:{etree.QName(marker).localname} has no xdr:{etree.QName(tag).localname}"
            )
        child.text = str(value)


def _move_anchor(anchor: _Element, address: str) -> None:
    """Move ``anchor``'s top-left cell to ``address`` and ``xdr:to`` by the same delta.

    Offsets within the cells stay, so a two-cell anchor keeps its span of cells.
    """
    col, row = split_address(address)
    new_col, new_row = col_to_index(col), row - 1
    from_elm = anchor.find(_XDR_FROM)
    old_col, col_off, old_row, row_off = _marker(from_elm)
    moves = [(from_elm, new_col, col_off, new_row, row_off)]
    to_elm = anchor.find(_XDR_TO)
    if to_elm is not None:
        to_col, to_col_off, to_row, to_row_off = _marker(to_elm)
        moves.append(
            (to_elm, to_col + new_col - old_col, to_col_off, to_row + new_row - old_row, to_row_off)
        )
    for _marker_elm, marker_col, _col_off, marker_row, _row_off in moves:
        if not (0 <= marker_col < _MAX_COL and 0 <= marker_row < _MAX_ROW):
            raise ValueError(f"anchoring at {address!r} puts a corner outside the sheet")
    for marker_elm, *position in moves:
        _set_marker(marker_elm, *position)


def _col_width_emu(width: float) -> int:
    """EMU of a column ``width`` (``col/@width``), per ECMA-376 Part 1, 18.3.1.13."""
    mdw = _MAX_DIGIT_WIDTH_PX
    return int(((256 * width + int(128 / mdw)) / 256) * mdw) * EMU_PER_PIXEL


class _SheetGrid:
    """Column widths and row heights of a worksheet, in EMU, for two-cell anchors.

    Pixel widths assume Excel's default 11-point body font. Hidden columns and
    rows are zero wide or high, as Excel draws them.
    """

    def __init__(self, worksheet: Worksheet):
        root = worksheet._part.element
        fmt = root.find(_SHEET_FORMAT_PR)
        default_width = fmt.get("defaultColWidth") if fmt is not None else None
        if default_width is not None:
            self._default_col = _col_width_emu(float(default_width))
        else:
            base = float(fmt.get("baseColWidth", "8")) if fmt is not None else 8.0
            # Excel rounds a default column up to a multiple of 8 pixels.
            px = math.ceil((base * _MAX_DIGIT_WIDTH_PX + 5) / 8) * 8
            self._default_col = px * EMU_PER_PIXEL
        default_height = fmt.get("defaultRowHeight") if fmt is not None else None
        self._default_row = round(float(default_height or 15) * _EMU_PER_POINT)

        self._cols: list[tuple[int, int, int]] = []
        cols = root.find(_COLS)
        for col in cols.iterchildren(_COL) if cols is not None else ():
            if col.get("hidden") in ("1", "true"):
                emu = 0
            elif col.get("width") is not None:
                emu = _col_width_emu(float(col.get("width")))
            else:
                emu = self._default_col
            self._cols.append((int(col.get("min", "1")) - 1, int(col.get("max", "1")) - 1, emu))

        self._rows: dict[int, int] = {}
        sheet_data = root.find(_SHEET_DATA)
        number = 0
        for row in sheet_data.iterchildren(_ROW) if sheet_data is not None else ():
            number = int(row.get("r") or number + 1)
            if row.get("hidden") in ("1", "true"):
                self._rows[number - 1] = 0
            elif row.get("ht") is not None:
                self._rows[number - 1] = round(float(row.get("ht")) * _EMU_PER_POINT)

    def col_emu(self, index: int) -> int:
        for low, high, emu in self._cols:
            if low <= index <= high:
                return emu
        return self._default_col

    def row_emu(self, index: int) -> int:
        return self._rows.get(index, self._default_row)

    def _end(self, size_of, limit: int, index: int, offset: int, length: int) -> tuple[int, int]:
        remaining = offset + length
        while True:
            size = size_of(index)
            if remaining < size:
                return index, remaining
            if index == limit - 1:
                raise ValueError("the resized object would extend beyond the sheet")
            remaining -= size
            index += 1

    @staticmethod
    def _corners(anchor: _Element) -> tuple[_Element, _Element]:
        from_elm, to_elm = anchor.find(_XDR_FROM), anchor.find(_XDR_TO)
        if from_elm is None or to_elm is None:
            raise ValueError("malformed anchor: a two-cell anchor needs xdr:from and xdr:to")
        return from_elm, to_elm

    def place_to(self, anchor: _Element, cx: int, cy: int) -> None:
        """Set ``xdr:to`` of a two-cell ``anchor`` to ``cx`` by ``cy`` EMU from ``xdr:from``."""
        from_elm, to_elm = self._corners(anchor)
        col, col_off, row, row_off = _marker(from_elm)
        to_col, to_col_off = self._end(self.col_emu, _MAX_COL, col, col_off, cx)
        to_row, to_row_off = self._end(self.row_emu, _MAX_ROW, row, row_off, cy)
        _set_marker(to_elm, to_col, to_col_off, to_row, to_row_off)

    def extent(self, anchor: _Element) -> tuple[int, int]:
        """``(cx, cy)`` in EMU of the box between ``xdr:from`` and ``xdr:to``."""
        from_elm, to_elm = self._corners(anchor)
        col, col_off, row, row_off = _marker(from_elm)
        to_col, to_col_off, to_row, to_row_off = _marker(to_elm)
        cx = sum(self.col_emu(i) for i in range(col, to_col)) - col_off + to_col_off
        cy = sum(self.row_emu(i) for i in range(row, to_row)) - row_off + to_row_off
        return max(0, cx), max(0, cy)


def _title_text_from_tx(tx: _Element) -> str | None:
    rich = tx.find(_C_RICH)
    if rich is not None:
        parts = [t.text for t in rich.iter(_A_T) if t.text]
        if parts:
            return "".join(parts)
    v = tx.find(_C_V)
    return v.text if v is not None else None


def _build_title_rich(title_elm: _Element, text: str) -> _Element:
    tx_pr = title_elm.find(_C_TX_PR)
    body_pr = tx_pr.find(_A_BODY_PR) if tx_pr is not None else None
    def_r_pr = tx_pr.find(f".//{_A_DEF_R_PR}") if tx_pr is not None else None

    rich = etree.Element(_C_RICH)
    if body_pr is not None:
        rich.append(deepcopy(body_pr))
    else:
        etree.SubElement(
            rich,
            _A_BODY_PR,
            rot="0",
            spcFirstLastPara="1",
            vertOverflow="ellipsis",
            vert="horz",
            wrap="square",
            anchor="ctr",
            anchorCtr="1",
        )
    etree.SubElement(rich, _A_LST_STYLE)

    para = etree.SubElement(rich, _A_P)
    p_pr = etree.SubElement(para, _A_P_PR)
    if def_r_pr is not None:
        p_pr.append(deepcopy(def_r_pr))
    else:
        etree.SubElement(
            p_pr,
            _A_DEF_R_PR,
            sz="1400",
            b="0",
            i="0",
            u="none",
            strike="noStrike",
            kern="1200",
            spc="0",
            baseline="0",
        )

    run = etree.SubElement(para, _A_R)
    etree.SubElement(run, _A_R_PR, lang="en-GB", sz="1400", b="1", i="0", u="none", strike="noStrike", baseline="0")
    t_elm = etree.SubElement(run, _A_T)
    t_elm.text = text
    etree.SubElement(para, _A_END_PARA_R_PR, lang="en-GB")
    return rich


def _parse_chart_formula(formula: str) -> tuple[str, str, int, str, int]:
    if "!" not in formula:
        raise ValueError(f"invalid chart formula: {formula!r}")
    sheet_part, range_part = formula.split("!", 1)
    if sheet_part.startswith("'") and sheet_part.endswith("'"):
        sheet_name = sheet_part[1:-1].replace("''", "'")
    else:
        sheet_name = sheet_part
    clean = range_part.replace("$", "")
    c1, r1, c2, r2 = parse_range(clean)
    return sheet_name, c1, r1, c2, r2


def _iter_range_cells(worksheet: Worksheet, c1: str, r1: int, c2: str, r2: int):
    from xlsxedit.oxml.address import col_to_index, index_to_col, join_address

    col_start = col_to_index(c1)
    col_end = col_to_index(c2)
    row_start, row_end = min(r1, r2), max(r1, r2)
    col_lo, col_hi = min(col_start, col_end), max(col_start, col_end)
    for row in range(row_start, row_end + 1):
        for col_idx in range(col_lo, col_hi + 1):
            yield worksheet[join_address(index_to_col(col_idx), row)]


def _rebuild_series_cache(parent: _Element, formula: str, worksheet: Worksheet) -> None:
    _, c1, r1, c2, r2 = _parse_chart_formula(formula)
    cells = list(_iter_range_cells(worksheet, c1, r1, c2, r2))

    if parent.tag == _C_STR_REF:
        cache = parent.find(_C_STR_CACHE)
        if cache is not None:
            parent.remove(cache)
        cache = etree.SubElement(parent, _C_STR_CACHE)
        etree.SubElement(cache, _C_PT_COUNT, val=str(len(cells)))
        for idx, cell in enumerate(cells):
            pt = etree.SubElement(cache, _C_PT, idx=str(idx))
            v = etree.SubElement(pt, _C_V)
            value = cell.value
            v.text = "" if value is None else str(value)
    elif parent.tag == _C_NUM_REF:
        cache = parent.find(_C_NUM_CACHE)
        if cache is not None:
            parent.remove(cache)
        cache = etree.SubElement(parent, _C_NUM_CACHE)
        etree.SubElement(cache, f"{{{CHART_NS}}}formatCode").text = "General"
        etree.SubElement(cache, _C_PT_COUNT, val=str(len(cells)))
        for idx, cell in enumerate(cells):
            pt = etree.SubElement(cache, _C_PT, idx=str(idx))
            v = etree.SubElement(pt, _C_V)
            value = cell.value
            v.text = "0" if value is None else str(value)


def _anchor_address(from_elm: _Element) -> str:
    col = int(from_elm.findtext(_XDR_COL, "0"))
    row = int(from_elm.findtext(_XDR_ROW, "0")) + 1
    return join_address(index_to_col(col), row)


def _cnvpr(obj_elm: _Element) -> _Element | None:
    """``cNvPr`` of a drawing object (``xdr:nvPicPr/xdr:cNvPr``, ``xdr:nvSpPr/xdr:cNvPr``, …)."""
    for nv_props in obj_elm:
        if isinstance(nv_props.tag, str) and etree.QName(nv_props).localname.startswith("nv"):
            for child in nv_props:
                if isinstance(child.tag, str) and etree.QName(child).localname == "cNvPr":
                    return child
            return None
    return None


class _AnchoredObject:
    """Cell-anchor geometry shared by :class:`Picture` and :class:`Chart`."""

    _kind = "object"

    def __init__(self, anchor_elm: _Element, element: _Element, drawing_part: DrawingPart):
        self._anchor = anchor_elm
        self._element = element
        self._drawing_part = drawing_part

    def _changed(self) -> None:
        self._drawing_part.mark_dirty()

    def _cell_anchor(self, action: str) -> _Element:
        anchor = self._anchor
        if anchor.find(_XDR_FROM) is None:
            raise ValueError(
                f"cannot {action} {self._kind} {self.name!r}: it has an absolute anchor, "
                "not a cell anchor"
            )
        return anchor

    @property
    def name(self) -> str | None:
        cnv = _cnvpr(self._element)
        return cnv.get("name") if cnv is not None else None

    @name.setter
    def name(self, value: str) -> None:
        cnv = _cnvpr(self._element)
        if cnv is None:
            raise ValueError(f"{self._kind} has no cNvPr to name")
        cnv.set("name", value)
        self._changed()

    @property
    def anchor(self) -> str:
        """Top-left cell; ``""`` for an absolute anchor."""
        from_elm = self._anchor.find(_XDR_FROM)
        if from_elm is None:
            return ""
        return _anchor_address(from_elm)

    @anchor.setter
    def anchor(self, address: str) -> None:
        """Move the top-left cell; keep the offsets and the span of cells to ``xdr:to``."""
        _move_anchor(self._cell_anchor("move"), address)
        self._changed()

    @property
    def offset_x(self) -> int:
        """Horizontal inset from the anchor cell, in pixels."""
        return _emu_to_px(_get_from_offset_emu(self._anchor)[0])

    @offset_x.setter
    def offset_x(self, pixels: int) -> None:
        pixels = _validate_offset_px(pixels)
        anchor = self._cell_anchor("offset")
        _, row_off = _get_from_offset_emu(anchor)
        _set_from_offset_emu(anchor, _px_to_emu(pixels), row_off)
        self._changed()

    @property
    def offset_y(self) -> int:
        """Vertical inset from the anchor cell, in pixels."""
        return _emu_to_px(_get_from_offset_emu(self._anchor)[1])

    @offset_y.setter
    def offset_y(self, pixels: int) -> None:
        pixels = _validate_offset_px(pixels)
        anchor = self._cell_anchor("offset")
        col_off, _ = _get_from_offset_emu(anchor)
        _set_from_offset_emu(anchor, col_off, _px_to_emu(pixels))
        self._changed()


class Picture(_AnchoredObject):
    """Proxy for an ``xdr:pic`` in a worksheet drawing."""

    _kind = "picture"

    def __init__(self, anchor_elm: _Element, drawing_part: DrawingPart, worksheet: Worksheet):
        pic = anchor_elm.find(_XDR_PIC)
        if pic is None:
            raise ValueError("anchor has no picture")
        super().__init__(anchor_elm, pic, drawing_part)
        self._worksheet = worksheet
        self._pic = pic

    def __repr__(self) -> str:
        return f"<Picture name={self.name!r} anchor={self.anchor!r}>"

    def _xfrm_ext(self) -> _Element | None:
        return self._pic.find(f"{_XDR_SPPR}/{_A_XFRM}/{_A_EXT}")

    def _extent_emu(self) -> tuple[int, int]:
        ext = self._xfrm_ext()
        if ext is None:
            ext = self._anchor.find(_XDR_EXT)
        if ext is not None:
            return int(ext.get("cx", "0")), int(ext.get("cy", "0"))
        if self._anchor.tag == _XDR_TWO_CELL:
            return _SheetGrid(self._worksheet).extent(self._anchor)
        return 0, 0

    @property
    def width_emu(self) -> int:
        return self._extent_emu()[0]

    @property
    def height_emu(self) -> int:
        return self._extent_emu()[1]

    @property
    def width(self) -> int:
        return max(1, round(self.width_emu / EMU_PER_PIXEL))

    @width.setter
    def width(self, pixels: int) -> None:
        pixels = _validate_size_px(pixels)
        self._set_ext(cx=int(pixels * EMU_PER_PIXEL), cy=self.height_emu or int(150 * EMU_PER_PIXEL))

    @property
    def height(self) -> int:
        return max(1, round(self.height_emu / EMU_PER_PIXEL))

    @height.setter
    def height(self, pixels: int) -> None:
        pixels = _validate_size_px(pixels)
        self._set_ext(cx=self.width_emu or int(200 * EMU_PER_PIXEL), cy=int(pixels * EMU_PER_PIXEL))

    def _set_ext(self, *, cx: int, cy: int) -> None:
        """Resize to ``cx`` by ``cy`` EMU: ``spPr/a:xfrm/a:ext`` and the anchor's extent."""
        anchor = self._anchor
        sp_pr = self._pic.find(_XDR_SPPR)
        if sp_pr is None:
            raise ValueError(f"picture {self.name!r} has no xdr:spPr")
        if anchor.tag == _XDR_TWO_CELL:
            _SheetGrid(self._worksheet).place_to(anchor, cx, cy)
        else:
            anchor_ext = anchor.find(_XDR_EXT)
            if anchor_ext is None:
                raise ValueError(f"malformed anchor: picture {self.name!r} has no xdr:ext")
            anchor_ext.set("cx", str(cx))
            anchor_ext.set("cy", str(cy))
        # Earlier releases appended a bare a:ext here, which CT_ShapeProperties does not allow.
        for stray in sp_pr.findall(_A_EXT):
            sp_pr.remove(stray)
        xfrm = sp_pr.find(_A_XFRM)
        if xfrm is None:
            xfrm = etree.Element(_A_XFRM)
            sp_pr.insert(0, xfrm)  # a:xfrm comes first in CT_ShapeProperties
        ext = xfrm.find(_A_EXT)
        if ext is None:
            ext = etree.SubElement(xfrm, _A_EXT)  # after a:off, the only other child
        ext.set("cx", str(cx))
        ext.set("cy", str(cy))
        self._changed()

    @property
    def media_path(self) -> str | None:
        part = self._media_part()
        return str(part.partname) if part is not None else None

    def _media_part(self) -> Part | None:
        blip_fill = self._pic.find(_XDR_BLIP_FILL)
        if blip_fill is None:
            return None
        blip = blip_fill.find(_A_BLIP)
        if blip is None:
            return None
        embed = blip.get(f"{{{OFFICE_REL_NS}}}embed")
        if embed is None:
            return None
        try:
            return self._drawing_part.rels[embed].target_part
        except KeyError:
            return None

    def replace(self, image_path: str | Path) -> None:
        """Swap image bytes in the linked media part."""
        part = self._media_part()
        if part is None:
            raise ValueError("picture has no linked media part")
        data = Path(image_path).read_bytes()
        part._blob = data
        ext = Path(image_path).suffix.lower().lstrip(".")
        ct = mimetypes.types_map.get(f".{ext}", CT.JPEG)
        part._content_type = ct


class Chart(_AnchoredObject):
    """Chart reference from a drawing graphicFrame."""

    _kind = "chart"

    def __init__(self, anchor_elm: _Element, drawing_part: DrawingPart, chart_part: Part | None = None):
        frame = anchor_elm.find(_XDR_GRAPHIC_FRAME)
        if frame is None:
            raise ValueError("not a chart frame")
        super().__init__(anchor_elm, frame, drawing_part)
        self._chart_part = chart_part
        self._frame = frame

    def __repr__(self) -> str:
        return f"<Chart name={self.name!r}>"

    def _resolve_chart_part(self) -> Part | None:
        if self._chart_part is None:
            self._chart_part = _chart_part_for(self._frame, self._drawing_part)
        return self._chart_part

    def _chart_root(self) -> _Element | None:
        part = self._resolve_chart_part()
        return part.element if isinstance(part, ChartPart) else None

    def _save_chart_root(self, root: _Element) -> None:
        part = self._resolve_chart_part()
        if isinstance(part, ChartPart):
            if root is not part.element:
                part._element = root
            part.mark_dirty()

    @property
    def to_anchor(self) -> str:
        """Bottom-right corner cell of the chart box."""
        to_elm = self._anchor.find(_XDR_TO)
        if to_elm is None:
            return ""
        return _anchor_address(to_elm)

    @to_anchor.setter
    def to_anchor(self, address: str) -> None:
        """Set the bottom-right corner cell (does not move ``anchor``)."""
        col, row = split_address(address)
        to_elm = self._anchor.find(_XDR_TO)
        if to_elm is None:
            raise ValueError(f"cannot resize chart {self.name!r}: its anchor has no xdr:to corner")
        _, col_off, _, row_off = _marker(to_elm)
        _set_marker(to_elm, col_to_index(col), col_off, row - 1, row_off)
        self._changed()

    @property
    def partname(self) -> str | None:
        part = self._resolve_chart_part()
        return str(part.partname) if part is not None else None

    @property
    def title(self) -> str | None:
        root = self._chart_root()
        if root is None:
            return None
        title_elm = root.find(f".//{_C_TITLE}")
        if title_elm is None:
            return None
        tx = title_elm.find(_C_TX)
        if tx is None:
            return None
        return _title_text_from_tx(tx)

    @title.setter
    def title(self, value: str | None) -> None:
        root = self._chart_root()
        if root is None:
            raise ValueError("chart has no linked chart part")
        chart_elm = root.find(f".//{{{CHART_NS}}}chart")
        if chart_elm is None:
            raise ValueError("invalid chart xml")
        title_elm = chart_elm.find(_C_TITLE)
        if title_elm is None:
            title_elm = etree.Element(_C_TITLE)
            insert_ordered(chart_elm, title_elm, ORDER["CT_Chart"])

        old_tx = title_elm.find(_C_TX)
        if old_tx is not None:
            title_elm.remove(old_tx)

        tx = etree.Element(_C_TX)
        title_elm.insert(0, tx)
        if value:
            tx.append(_build_title_rich(title_elm, value))

        self._save_chart_root(root)

    def set_series_formula(
        self,
        index: int,
        formula: str,
        worksheet: Worksheet | None = None,
    ) -> None:
        """Set the ``c:f`` formula for a series value reference by index."""
        root = self._chart_root()
        if root is None:
            raise ValueError("chart has no linked chart part")
        formulas = [f for f in root.iter(_C_F) if f.text and "$" in (f.text or "")]
        if index < 0 or index >= len(formulas):
            raise IndexError(f"series formula index out of range: {index}")
        target = formulas[index]
        parent = target.getparent()
        if parent is not None and worksheet is not None:
            # Parses the formula first, so a bad one raises before the tree changes.
            _rebuild_series_cache(parent, formula, worksheet)
        target.text = formula
        self._save_chart_root(root)


class Table:
    """Excel table metadata with resize support."""

    def __init__(self, element: _Element, part: Part | None = None):
        self._element = element
        self._part = part

    def __repr__(self) -> str:
        return f"<Table name={self.name!r} ref={self.ref!r}>"

    @property
    def name(self) -> str | None:
        return self._element.get("name")

    @property
    def ref(self) -> str | None:
        return self._element.get("ref")

    @property
    def columns(self) -> list[str]:
        cols = self._element.find(_TABLE_COLUMNS)
        if cols is None:
            return []
        return [
            c.get("name", "")
            for c in cols.findall(_TABLE_COLUMN)
            if c.get("name")
        ]

    def resize(self, ref: str) -> None:
        """Update table ``ref`` and ``autoFilter`` range."""
        self._element.set("ref", ref)
        af = self._element.find(_AUTO_FILTER)
        if af is not None:
            af.set("ref", ref)
        if self._part is not None:
            self._part._blob = serialize_xml(self._element)


def _chart_part_for(frame: _Element, drawing_part: DrawingPart) -> Part | None:
    chart = frame.find(f"{_A_GRAPHIC}/{_A_GRAPHIC_DATA}/{_C_CHART}")
    r_id = chart.get(f"{{{OFFICE_REL_NS}}}id") if chart is not None else None
    if r_id is None:
        return None
    try:
        return drawing_part.rels[r_id].target_part
    except KeyError:
        return None


def iter_pictures(drawing_part: DrawingPart, worksheet: Worksheet) -> Iterator[Picture]:
    root = drawing_part.element
    for tag in (_XDR_TWO_CELL, _XDR_ONE_CELL):
        for anchor in root.findall(tag):
            if anchor.find(_XDR_PIC) is not None:
                yield Picture(anchor, drawing_part, worksheet)


def iter_charts(drawing_part: DrawingPart) -> Iterator[Chart]:
    root = drawing_part.element
    for anchor in root.findall(_XDR_TWO_CELL):
        frame = anchor.find(_XDR_GRAPHIC_FRAME)
        if frame is not None:
            yield Chart(anchor, drawing_part, _chart_part_for(frame, drawing_part))


def drawing_parts_for_worksheet(worksheet: Worksheet) -> list[DrawingPart]:
    parts = []
    ws_elm = worksheet._part.element
    for drawing_elm in ws_elm.findall(_WS_DRAWING):
        r_id = drawing_elm.get(f"{{{OFFICE_REL_NS}}}id")
        if r_id is None:
            continue
        try:
            part = worksheet._part.rels[r_id].target_part
        except KeyError:
            continue
        if isinstance(part, DrawingPart):
            parts.append(part)
    return parts


def ensure_drawing_part(worksheet: Worksheet) -> DrawingPart:
    """The worksheet's drawing part, created and related on first use."""
    parts = drawing_parts_for_worksheet(worksheet)
    if parts:
        return parts[0]
    if worksheet._part.element.find(_WS_DRAWING) is not None:
        raise ValueError(f"worksheet {worksheet.name!r} links a drawing that is not a drawing part")
    package = worksheet._workbook._package
    root = etree.Element(_XDR_WSDR, nsmap={"xdr": XDR_NS, "a": A_NS, "r": OFFICE_REL_NS})
    part = DrawingPart(package.next_partname("/xl/drawings/drawing%d.xml"), CT.DRAWING, root, package)
    package._add_part(part)
    drawing_elm = etree.Element(_WS_DRAWING)
    drawing_elm.set(f"{{{OFFICE_REL_NS}}}id", worksheet._part.relate_to(part, RT.DRAWING))
    insert_worksheet_child(worksheet._part.element, drawing_elm)
    return part


def table_parts_for_worksheet(worksheet: Worksheet) -> list[Table]:
    tables: list[Table] = []
    tp = worksheet._part.element.find(_TABLE_PARTS)
    if tp is None:
        return tables
    for rel_elm in tp.findall(_TABLE_PART):
        r_id = rel_elm.get(f"{{{OFFICE_REL_NS}}}id")
        if r_id is None:
            continue
        try:
            part = worksheet._part.rels[r_id].target_part
            element = parse_xml(part.blob)
            tables.append(Table(element, part))
        except KeyError:
            continue
    return tables

"""Targeted facts about a workbook package, for comparisons and the fixture manifest.

Each extractor reads a package with lxml and returns plain data (tuples,
dictionaries and strings) keyed by sheet name. Two packages that Excel would
show the same way give equal facts: what Excel rewrites on every save is left
out or resolved, such as ``xr:uid`` and other revision ids, relationship ids,
drawing object ids, style indexes and rule priorities (kept only as order).

``tests/test_excel_reference.py`` compares these facts with Excel's own
result; ``scripts/describe_fixtures.py`` prints them in ``fixtures/MANIFEST.md``.
Like ``tests/preservation.py``, nothing here imports pytest or xlsxedit.
"""

from __future__ import annotations

import posixpath
import re
from collections import Counter
from collections.abc import Iterator, Mapping
from functools import cached_property
from typing import NamedTuple

from lxml import etree

from tests.preservation import content_types, read_pkg

SML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM = "http://schemas.microsoft.com/office/excel/2006/main"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
TC = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"

_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True)
_REVISION_NS = (
    "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
)

# ECMA-376 Part 1, 18.8.30: the number formats that numFmtId < 164 names without a numFmt.
BUILTIN_FORMATS = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%",
    11: "0.00E+00", 12: "# ?/?", 13: "# ??/??", 14: "mm-dd-yy", 15: "d-mmm-yy", 16: "d-mmm",
    17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss",
    22: "m/d/yy h:mm", 37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 45: "mm:ss",
    46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
}  # fmt: skip


def sml(name: str) -> str:
    return f"{{{SML}}}{name}"


def _local(tag) -> str:
    return etree.QName(tag).localname if isinstance(tag, str) else str(tag)


class Sheet(NamedTuple):
    name: str
    member: str
    state: str


class Book:
    """Read-only navigation over a package: workbook, sheets, related parts, styles, strings."""

    def __init__(self, source):
        self.pkg: Mapping[str, bytes] = source if isinstance(source, Mapping) else read_pkg(source)
        self._trees: dict[str, etree._Element] = {}

    def tree(self, member: str) -> etree._Element:
        if member not in self._trees:
            self._trees[member] = etree.fromstring(self.pkg[member], _PARSER)
        return self._trees[member]

    def rels(self, source: str) -> list[tuple[str, str, str]]:
        """``(rId, type, target member)`` for the internal relationships of ``source``."""
        folder, name = posixpath.split(source)
        rels_member = posixpath.join(folder, "_rels", f"{name}.rels")
        if rels_member not in self.pkg:
            return []
        found = []
        for rel in self.tree(rels_member).iter(f"{{{PKG_REL}}}Relationship"):
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target", "")
            if target.startswith("/"):
                member = target[1:]
            else:
                member = posixpath.normpath(posixpath.join(folder, target))
            found.append((rel.get("Id", ""), rel.get("Type", ""), member))
        return found

    def related(self, source: str, kind: str) -> list[str]:
        """Members ``source`` relates to with a relationship type ending in ``/kind``."""
        return [m for _, t, m in self.rels(source) if t.endswith(f"/{kind}") and m in self.pkg]

    def target(self, source: str, r_id: str | None) -> str | None:
        for rid, _, member in self.rels(source):
            if rid == r_id and member in self.pkg:
                return member
        return None

    @cached_property
    def workbook(self) -> str:
        return self.related("", "officeDocument")[0]

    @cached_property
    def sheets(self) -> list[Sheet]:
        """Worksheets in tab order (chartsheets and macro sheets are left out)."""
        worksheets = {
            r_id: member
            for r_id, kind, member in self.rels(self.workbook)
            if kind.endswith("/worksheet") and member in self.pkg
        }
        found = []
        for sheet in self.tree(self.workbook).iter(sml("sheet")):
            member = worksheets.get(sheet.get(f"{{{R}}}id"))
            if member is not None:
                found.append(Sheet(sheet.get("name", ""), member, sheet.get("state", "visible")))
        return found

    def sheet(self, name: str) -> Sheet:
        for sheet in self.sheets:
            if sheet.name == name:
                return sheet
        raise KeyError(name)

    @cached_property
    def shared_strings(self) -> list[str]:
        members = self.related(self.workbook, "sharedStrings")
        if not members:
            return []
        return [_text(si) for si in self.tree(members[0]).iterchildren(sml("si"))]

    @cached_property
    def _styles(self) -> etree._Element | None:
        members = self.related(self.workbook, "styles")
        return self.tree(members[0]) if members else None

    def _style_list(self, collection: str, item: str) -> list[etree._Element]:
        root = self._styles
        if root is None:
            return []
        parent = root.find(sml(collection))
        return [] if parent is None else list(parent.iterchildren(sml(item)))

    @cached_property
    def number_formats(self) -> dict[int, str]:
        formats = dict(BUILTIN_FORMATS)
        for fmt in self._style_list("numFmts", "numFmt"):
            formats[int(fmt.get("numFmtId", "0"))] = fmt.get("formatCode", "")
        return formats

    def number_format(self, fmt_id: str | int | None) -> str:
        fmt = int(fmt_id or 0)
        return self.number_formats.get(fmt, f"builtin {fmt}")

    def dxf(self, dxf_id: str | None) -> str | None:
        """A differential format as text, or None without one."""
        if dxf_id is None:
            return None
        dxfs = self._style_list("dxfs", "dxf")
        index = int(dxf_id)
        return describe(dxfs[index], drop={"numFmtId"}) if index < len(dxfs) else f"dxf {index}?"

    @cached_property
    def cell_styles(self) -> list[dict[str, str]]:
        """Every ``cellXfs`` entry resolved to the format it shows, one text per component."""
        fonts = self._style_list("fonts", "font")
        fills = self._style_list("fills", "fill")
        borders = self._style_list("borders", "border")

        def pick(items: list, index: str | None) -> str:
            i = int(index or 0)
            return describe(items[i]) if i < len(items) else f"#{i}?"

        resolved = []
        for xf in self._style_list("cellXfs", "xf"):
            parts = {
                "numFmt": f"numFmt={self.number_format(xf.get('numFmtId'))!r}",
                "font": pick(fonts, xf.get("fontId")),
                "fill": pick(fills, xf.get("fillId")),
                "border": pick(borders, xf.get("borderId")),
                "alignment": "",
                "protection": "",
                "quotePrefix": "quotePrefix" if xf.get("quotePrefix") in ("1", "true") else "",
            }
            for child in xf:
                if isinstance(child.tag, str) and _local(child.tag) in ("alignment", "protection"):
                    parts[_local(child.tag)] = describe(child)
            resolved.append(parts)
        return resolved

    def cell_style(self, index: str | None) -> str:
        """The format ``cellXfs[index]`` shows, as the components that differ from ``cellXfs[0]``.

        ``""`` is the default format.
        """
        styles = self.cell_styles
        i = int(index or 0)
        if i >= len(styles):
            return f"xf {i}?"
        default = styles[0] if styles else {}
        return " ".join(v for k, v in styles[i].items() if v != default.get(k))


def describe(elm: etree._Element, drop: set[str] | frozenset[str] = frozenset()) -> str:
    """Compact, order-stable text for an element: ``name(attr=value,...)[children]``.

    Namespaces, revision ids and the ``drop`` attributes are left out.
    """
    attrs = [
        f"{_local(k)}={v}"
        for k, v in sorted(elm.attrib.items())
        if _local(k) not in drop and etree.QName(k).namespace not in _REVISION_NS
    ]
    children = [describe(c, drop) for c in elm if isinstance(c.tag, str)]
    text = (elm.text or "").strip()
    out = _local(elm.tag)
    if attrs:
        out += f"({','.join(attrs)})"
    if text:
        out += f"={text!r}"
    if children:
        out += f"[{';'.join(children)}]"
    return out


def _text(elm: etree._Element) -> str:
    """Concatenated ``t`` text of a string item, skipping phonetic runs."""
    return "".join(
        t.text or ""
        for t in elm.iter(sml("t"))
        if t.getparent() is not None and _local(t.getparent().tag) != "rPh"
    )


# references ----------------------------------------------------------------

_CELL = re.compile(r"^\$?([A-Z]{1,3})\$?([0-9]+)$")


def col_index(letters: str) -> int:
    index = 0
    for ch in letters.upper():
        index = index * 26 + ord(ch) - 64
    return index


def col_letters(index: int) -> str:
    letters = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def parse_cell(ref: str) -> tuple[int, int]:
    """``(row, col)`` of an A1 reference."""
    m = _CELL.match(ref.upper())
    if m is None:
        raise ValueError(f"not a cell reference: {ref!r}")
    return int(m.group(2)), col_index(m.group(1))


def parse_box(ref: str) -> tuple[int, int, int, int]:
    """``(row1, col1, row2, col2)`` of ``A1`` or ``A1:B2``."""
    first, _, last = ref.partition(":")
    r1, c1 = parse_cell(first)
    r2, c2 = parse_cell(last or first)
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def sqref(text: str | None) -> str:
    """A space-separated area list in reading order, so area order does not matter."""
    areas = (text or "").split()

    def key(area: str):
        try:
            return parse_box(area.replace("$", ""))
        except ValueError:
            return (0, 0, 0, 0)

    return " ".join(sorted(areas, key=lambda a: (key(a), a)))


# tables --------------------------------------------------------------------


class TableColumn(NamedTuple):
    name: str
    formula: str | None
    totals: str | None
    data_format: str | None


class Table(NamedTuple):
    name: str
    ref: str
    header_rows: int
    totals_rows: int
    autofilter: str | None
    sort: tuple[str, tuple[tuple[str, bool], ...]] | None
    style: str | None
    columns: tuple[TableColumn, ...]


def tables(book: Book) -> dict[str, tuple[Table, ...]]:
    found = {}
    for sheet in book.sheets:
        items = []
        for member in book.related(sheet.member, "table"):
            items.append(_table(book, book.tree(member)))
        if items:
            found[sheet.name] = tuple(sorted(items))
    return found


def _table(book: Book, root: etree._Element) -> Table:
    af = root.find(sml("autoFilter"))
    sort_state = root.find(sml("sortState"))
    if sort_state is None and af is not None:
        sort_state = af.find(sml("sortState"))
    sort = None
    if sort_state is not None:
        sort = (
            sort_state.get("ref", ""),
            tuple(
                (c.get("ref", ""), c.get("descending") in ("1", "true"))
                for c in sort_state.iter(sml("sortCondition"))
            ),
        )
    style = root.find(sml("tableStyleInfo"))
    columns = []
    for col in root.iter(sml("tableColumn")):
        calc = col.find(sml("calculatedColumnFormula"))
        totals_formula = col.find(sml("totalsRowFormula"))
        if col.get("totalsRowFunction") == "custom" and totals_formula is not None:
            totals = f"={totals_formula.text}"
        elif col.get("totalsRowFunction") not in (None, "none"):
            totals = col.get("totalsRowFunction")
        elif col.get("totalsRowLabel") is not None:
            totals = f"label {col.get('totalsRowLabel')!r}"
        else:
            totals = None
        columns.append(
            TableColumn(
                col.get("name", ""),
                calc.text if calc is not None else None,
                totals,
                book.dxf(col.get("dataDxfId")),
            )
        )
    return Table(
        name=root.get("name", ""),
        ref=root.get("ref", ""),
        header_rows=int(root.get("headerRowCount", "1")),
        totals_rows=int(root.get("totalsRowCount", "0")),
        autofilter=af.get("ref") if af is not None else None,
        sort=sort,
        style=describe(style) if style is not None else None,
        columns=tuple(columns),
    )


# conditional formatting ----------------------------------------------------


class Rule(NamedTuple):
    sqref: str
    source: str  # "" for the base element, "x14" for an extension-only rule, or an orphan twin
    type: str
    detail: str


def conditional_formats(book: Book) -> dict[str, tuple[Rule, ...]]:
    """Rules in priority order; a base rule and its x14 twin count as one rule.

    The twin's own ``xm:sqref`` is part of the rule, and a twin no base rule
    claims is listed on its own, so neither can drift unnoticed.
    """
    found = {}
    for sheet in book.sheets:
        root = book.tree(sheet.member)
        twins: dict[str, tuple[str, etree._Element]] = {}
        ranked: list[tuple[int, int, Rule]] = []
        for block in root.iter(f"{{{X14}}}conditionalFormatting"):
            ref = block.find(f"{{{XM}}}sqref")
            area = sqref(ref.text if ref is not None else None)
            for rule in block.iterchildren(f"{{{X14}}}cfRule"):
                if rule.get("priority") is None and rule.get("id"):
                    twins[rule.get("id")] = (area, rule)
                    continue
                ranked.append(
                    (_priority(rule), len(ranked), Rule(area, "x14", rule.get("type", ""), _rule(book, rule)))
                )
        claimed = set()
        for block in root.iterchildren(sml("conditionalFormatting")):
            area = sqref(block.get("sqref"))
            for rule in block.iterchildren(sml("cfRule")):
                detail = _rule(book, rule)
                twin_id = rule.find(f"{sml('extLst')}/{sml('ext')}/{{{X14}}}id")
                if twin_id is not None and twin_id.text in twins:
                    twin_area, twin = twins[twin_id.text]
                    detail += f" x14@{twin_area}:{_rule(book, twin)}"
                    claimed.add(twin_id.text)
                ranked.append((_priority(rule), len(ranked), Rule(area, "", rule.get("type", ""), detail)))
        for twin_id, (area, twin) in twins.items():
            if twin_id not in claimed:
                orphan = Rule(area, "x14 twin without base rule", twin.get("type", ""), _rule(book, twin))
                ranked.append((_priority(twin), len(ranked), orphan))
        if ranked:
            found[sheet.name] = tuple(r for _, _, r in sorted(ranked))
    return found


def _priority(rule: etree._Element) -> int:
    try:
        return int(rule.get("priority", ""))
    except ValueError:
        return 1 << 30


def _rule(book: Book, rule: etree._Element) -> str:
    drop = {"priority", "dxfId", "id", "type"}
    attrs = ",".join(
        f"{_local(k)}={v}" for k, v in sorted(rule.attrib.items()) if _local(k) not in drop
    )
    parts = [attrs] if attrs else []
    for child in rule:
        if not isinstance(child.tag, str) or _local(child.tag) == "extLst":
            continue
        parts.append(f"={child.text}" if _local(child.tag) in ("formula", "f") else describe(child))
    dxf = rule.get("dxfId")
    if dxf is not None:
        parts.append(f"dxf {book.dxf(dxf)}")
    else:
        inline = rule.find(f"{{{X14}}}dxf")
        if inline is not None:
            parts.append(f"dxf {describe(inline)}")
    return " ".join(parts)


# data validation -----------------------------------------------------------


class Validation(NamedTuple):
    sqref: str
    source: str
    type: str
    detail: str


def data_validations(book: Book) -> dict[str, tuple[Validation, ...]]:
    found = {}
    for sheet in book.sheets:
        root = book.tree(sheet.member)
        items = []
        for dv in root.iter(sml("dataValidation")):
            items.append(Validation(sqref(dv.get("sqref")), "", dv.get("type", "none"), _validation(dv)))
        for dv in root.iter(f"{{{X14}}}dataValidation"):
            ref = dv.find(f"{{{XM}}}sqref")
            area = sqref(ref.text if ref is not None else None)
            items.append(Validation(area, "x14", dv.get("type", "none"), _validation(dv)))
        if items:
            found[sheet.name] = tuple(sorted(items))
    return found


def _validation(dv: etree._Element) -> str:
    attrs = ",".join(
        f"{_local(k)}={v}"
        for k, v in sorted(dv.attrib.items())
        if _local(k) not in ("sqref", "type", "uid")
    )
    parts = [attrs] if attrs else []
    for name in ("formula1", "formula2"):
        for formula in dv.iter(sml(name), f"{{{X14}}}{name}"):
            text = formula.findtext(f"{{{XM}}}f") if formula.find(f"{{{XM}}}f") is not None else formula.text
            parts.append(f"{name}={text}")
    return " ".join(parts)


# drawings ------------------------------------------------------------------


class Anchor(NamedTuple):
    kind: str  # twoCellAnchor, oneCellAnchor, absoluteAnchor
    edit_as: str | None
    start: tuple[int, int, int, int] | None  # (col, colOff, row, rowOff), zero-based like the XML
    end: tuple[int, int, int, int] | None
    position: tuple[int, int] | None  # absoluteAnchor xdr:pos
    extent: tuple[int, int] | None  # oneCellAnchor/absoluteAnchor xdr:ext
    objects: tuple[tuple[str, str], ...]  # (kind, name), the anchored object first, then group members


_OBJECT_KINDS = {
    "sp": "shape",
    "pic": "picture",
    "graphicFrame": "graphicFrame",
    "grpSp": "group",
    "cxnSp": "connector",
    "contentPart": "contentPart",
}


def drawings(book: Book) -> dict[str, tuple[Anchor, ...]]:
    """Anchors in document order; an anchor inside ``mc:AlternateContent`` counts by its Choice."""
    found = {}
    for sheet in book.sheets:
        anchors = []
        for member in book.related(sheet.member, "drawing"):
            anchors += [_anchor(a) for a in _anchors(book.tree(member))]
        if anchors:
            found[sheet.name] = tuple(anchors)
    return found


def _anchors(parent: etree._Element) -> Iterator[etree._Element]:
    for child in parent:
        if not isinstance(child.tag, str):
            continue
        if child.tag == f"{{{MC}}}AlternateContent":
            choice = child.find(f"{{{MC}}}Choice")
            if choice is not None:
                yield from _anchors(choice)
        elif etree.QName(child).namespace == XDR:
            yield child


def _marker(elm: etree._Element | None) -> tuple[int, int, int, int] | None:
    if elm is None:
        return None
    return tuple(int(elm.findtext(f"{{{XDR}}}{n}") or 0) for n in ("col", "colOff", "row", "rowOff"))


def _anchor(anchor: etree._Element) -> Anchor:
    pos = anchor.find(f"{{{XDR}}}pos")
    ext = anchor.find(f"{{{XDR}}}ext")
    return Anchor(
        kind=_local(anchor.tag),
        edit_as=anchor.get("editAs"),
        start=_marker(anchor.find(f"{{{XDR}}}from")),
        end=_marker(anchor.find(f"{{{XDR}}}to")),
        position=(int(pos.get("x", 0)), int(pos.get("y", 0))) if pos is not None else None,
        extent=(int(ext.get("cx", 0)), int(ext.get("cy", 0))) if ext is not None else None,
        objects=tuple(_objects(anchor)),
    )


def _objects(parent: etree._Element) -> Iterator[tuple[str, str]]:
    """Drawing objects under ``parent`` in document order, following ``mc:Choice`` only."""
    for child in parent:
        if not isinstance(child.tag, str):
            continue
        if child.tag == f"{{{MC}}}AlternateContent":
            choice = child.find(f"{{{MC}}}Choice")
            if choice is not None:
                yield from _objects(choice)
            continue
        kind = _OBJECT_KINDS.get(_local(child.tag)) if etree.QName(child).namespace == XDR else None
        if kind is None:
            continue
        props = child.find(f"*/{{{XDR}}}cNvPr")
        name = props.get("name", "") if props is not None else ""
        if kind == "shape" and child.find(f"{{{XDR}}}nvSpPr/{{{XDR}}}cNvSpPr[@txBox='1']") is not None:
            kind = "textBox"
        if kind == "graphicFrame":
            data = child.find(f"{{{A}}}graphic/{{{A}}}graphicData")
            uri = data.get("uri", "") if data is not None else ""
            kind = uri.rstrip("/").rsplit("/", 1)[-1] or kind
        yield kind, name
        if kind == "group":
            yield from _objects(child)


def drawings_by_name(book: Book) -> dict[str, dict[str, Anchor]]:
    """Anchors keyed by the name of the object they hold, per sheet.

    A name used twice gets ``#n``, its count so far, appended from the second time on.
    """
    found = {}
    for sheet, anchors in drawings(book).items():
        keyed: dict[str, Anchor] = {}
        seen: Counter[str] = Counter()
        for anchor in anchors:
            name = anchor.objects[0][1] if anchor.objects else ""
            seen[name] += 1
            keyed[name if seen[name] == 1 else f"{name} #{seen[name]}"] = anchor
        found[sheet] = keyed
    return found


# names, merges -------------------------------------------------------------


def defined_names(book: Book) -> dict[tuple[str, str], str]:
    """``{(name, scope): formula}``; scope is the sheet name of a local name, else ``""``."""
    tabs = [s.get("name", "") for s in book.tree(book.workbook).iter(sml("sheet"))]
    found = {}
    for name in book.tree(book.workbook).iter(sml("definedName")):
        local = name.get("localSheetId")
        scope = tabs[int(local)] if local is not None and int(local) < len(tabs) else ""
        found[(name.get("name", ""), scope)] = name.text or ""
    return found


def merges(book: Book) -> dict[str, tuple[str, ...]]:
    found = {}
    for sheet in book.sheets:
        refs = [m.get("ref", "") for m in book.tree(sheet.member).iter(sml("mergeCell"))]
        if refs:
            found[sheet.name] = tuple(sorted(refs, key=lambda r: (parse_box(r), r)))
    return found


# cells ---------------------------------------------------------------------


class Formula(NamedTuple):
    kind: str  # normal, shared, array, dynamic, dataTable
    text: str | None  # None for a shared-formula follower
    ref: str | None
    options: str  # other attributes, such as a data table's input cells r1 and r2


def _formula(cell: etree._Element) -> Formula | None:
    f = cell.find(sml("f"))
    if f is None:
        return None
    kind = f.get("t", "normal")
    if kind == "array" and cell.get("cm") is not None:
        kind = "dynamic"
    options = ",".join(
        f"{_local(k)}={v}" for k, v in sorted(f.attrib.items()) if _local(k) not in ("t", "ref", "si")
    )
    return Formula(kind, f.text, f.get("ref"), options)


class CellFact(NamedTuple):
    type: str | None  # n, s, b, e or d; None for a formula cell
    value: object  # None for a formula cell
    formula: Formula | None


_VALUE_TYPES = {"s": "s", "inlineStr": "s", "str": "s", "b": "b", "e": "e", "d": "d"}


def _value(book: Book, cell: etree._Element) -> tuple[str, object]:
    """``(type, value)``: ``n`` number, ``s`` text, ``b`` boolean, ``e`` error, ``d`` ISO date."""
    t = cell.get("t", "n")
    kind = _VALUE_TYPES.get(t, "n")
    if t == "inlineStr":
        inline = cell.find(sml("is"))
        return kind, _text(inline) if inline is not None else ""
    v = cell.findtext(sml("v"))
    if v is None:
        return kind, None
    if t == "s":
        index = int(v)
        strings = book.shared_strings
        return kind, strings[index] if index < len(strings) else f"#sst {index}?"
    if t == "b":
        return kind, v in ("1", "true")
    if kind != "n":
        return kind, v
    try:
        return kind, float(v)
    except ValueError:
        return kind, v


def _rows(root: etree._Element) -> Iterator[tuple[int, etree._Element]]:
    data = root.find(sml("sheetData"))
    previous = 0
    for row in [] if data is None else data.iterchildren(sml("row")):
        previous = int(row.get("r") or previous + 1)
        yield previous, row


def _cells(row_number: int, row: etree._Element) -> Iterator[tuple[int, etree._Element]]:
    previous = 0
    for cell in row.iterchildren(sml("c")):
        ref = cell.get("r")
        previous = parse_cell(ref)[1] if ref else previous + 1
        yield previous, cell


def cells(book: Book, band: str) -> dict[str, CellFact]:
    """``{address: CellFact}`` for the cells in ``band`` (``Sheet!A1:F20``) with content.

    A formula cell's cached value and its type are left out: they are only as
    fresh as the last recalculation.
    """
    sheet, box = _band(book, band)
    r1, c1, r2, c2 = box
    found = {}
    for number, row in _rows(book.tree(sheet.member)):
        if not r1 <= number <= r2:
            continue
        for col, cell in _cells(number, row):
            if not c1 <= col <= c2:
                continue
            formula = _formula(cell)
            kind, value = (None, None) if formula is not None else _value(book, cell)
            if formula is None and value is None:
                continue
            found[f"{col_letters(col)}{number}"] = CellFact(kind, value, formula)
    return found


def styles(book: Book, band: str) -> dict[str, str]:
    """``{address: resolved format}`` for the cells in ``band`` that are not in the default format.

    A cell takes its own ``s``, else its row's (``customFormat``), else its column's.
    """
    sheet, box = _band(book, band)
    r1, c1, r2, c2 = box
    root = book.tree(sheet.member)
    column_style: dict[int, str] = {}
    for col in root.iter(sml("col")):
        if col.get("style") is not None:
            for index in range(int(col.get("min", "1")), int(col.get("max", "1")) + 1):
                column_style[index] = col.get("style")
    present: dict[tuple[int, int], str | None] = {}
    row_style: dict[int, str] = {}
    for number, row in _rows(root):
        if row.get("customFormat") in ("1", "true") and row.get("s") is not None:
            row_style[number] = row.get("s")
        for col, cell in _cells(number, row):
            present[(number, col)] = cell.get("s")
    found = {}
    for number in range(r1, r2 + 1):
        for col in range(c1, c2 + 1):
            if (number, col) in present:
                index = present[(number, col)]
            else:
                index = row_style.get(number, column_style.get(col))
            style = book.cell_style(index)
            if style:
                found[f"{col_letters(col)}{number}"] = style
    return found


def rows(book: Book, band: str) -> dict[int, str]:
    """Row attributes Excel shows in ``band``'s rows: height, hidden, outline, row format."""
    sheet, (r1, _, r2, _) = _band(book, band)
    found = {}
    for number, row in _rows(book.tree(sheet.member)):
        if not r1 <= number <= r2:
            continue
        parts = []
        if row.get("customHeight") in ("1", "true"):
            parts.append(f"ht={row.get('ht')}")
        for flag in ("hidden", "collapsed", "thickTop", "thickBot"):
            if row.get(flag) in ("1", "true"):
                parts.append(flag)
        if row.get("outlineLevel") not in (None, "0"):
            parts.append(f"outlineLevel={row.get('outlineLevel')}")
        if row.get("customFormat") in ("1", "true"):
            parts.append(f"format {book.cell_style(row.get('s'))}")
        if parts:
            found[number] = " ".join(parts)
    return found


def _band(book: Book, band: str) -> tuple[Sheet, tuple[int, int, int, int]]:
    name, _, ref = band.rpartition("!")
    return book.sheet(name.strip("'")), parse_box(ref)


# sheet- and package-level features, for the manifest -------------------------


def formula_summary(book: Book) -> dict[str, dict[str, object]]:
    """Per sheet: formula counts by kind, shared groups, and the ``_xlfn.`` functions used."""
    found = {}
    for sheet in book.sheets:
        kinds: Counter[str] = Counter()
        groups = set()
        future: set[str] = set()
        cross_sheet = 0
        for number, row in _rows(book.tree(sheet.member)):
            for _, cell in _cells(number, row):
                formula = _formula(cell)
                if formula is None:
                    continue
                kinds[formula.kind] += 1
                if formula.kind == "shared":
                    groups.add(cell.find(sml("f")).get("si"))
                text = formula.text or ""
                future.update(re.findall(r"_xlfn\.(?:_xlws\.)?([A-Z][A-Z0-9.]*)", text))
                if "!" in text:
                    cross_sheet += 1
        if kinds:
            found[sheet.name] = {
                "kinds": dict(sorted(kinds.items())),
                "shared_groups": len(groups),
                "future_functions": sorted(future),
                "cross_sheet": cross_sheet,
            }
    return found


def sheet_features(book: Book) -> dict[str, list[str]]:
    """Per sheet, one line per notable feature: view, filters, outline, comments, pivots..."""
    found = {}
    for sheet in book.sheets:
        root = book.tree(sheet.member)
        lines = []
        pane = root.find(f"{sml('sheetViews')}/{sml('sheetView')}/{sml('pane')}")
        if pane is not None:
            lines.append(
                f"pane: {pane.get('state', 'split')} at {pane.get('topLeftCell')} "
                f"(xSplit={pane.get('xSplit', '0')}, ySplit={pane.get('ySplit', '0')})"
            )
        af = root.find(sml("autoFilter"))
        if af is not None:
            columns = [describe(c, drop={"uid"}) for c in af.iterchildren(sml("filterColumn"))]
            lines.append(f"autoFilter: {af.get('ref')}" + (f" {'; '.join(columns)}" if columns else ""))
        widths = [
            f"{c.get('min')}-{c.get('max')}={c.get('width')}"
            if c.get("min") != c.get("max")
            else f"{col_letters(int(c.get('min', '0')))}={c.get('width')}"
            for c in root.iter(sml("col"))
            if c.get("customWidth") in ("1", "true")
        ]
        if widths:
            lines.append(f"custom column widths: {', '.join(widths)}")
        hidden_cols = [c.get("min") for c in root.iter(sml("col")) if c.get("hidden") in ("1", "true")]
        if hidden_cols:
            lines.append(f"hidden columns from: {', '.join(hidden_cols)}")
        heights, hidden, outline = [], [], []
        for number, row in _rows(root):
            if row.get("customHeight") in ("1", "true"):
                heights.append(f"{number}={row.get('ht')}")
            if row.get("hidden") in ("1", "true"):
                hidden.append(str(number))
            if row.get("outlineLevel") not in (None, "0"):
                outline.append(f"{number}:{row.get('outlineLevel')}")
        if heights:
            lines.append(f"custom row heights: {', '.join(heights)}")
        if hidden:
            lines.append(f"hidden rows: {', '.join(hidden)}")
        if outline:
            fmt = root.find(sml("sheetFormatPr"))
            level = fmt.get("outlineLevelRow") if fmt is not None else None
            lines.append(f"row outline (row:level): {', '.join(outline)}; outlineLevelRow={level}")
        sparklines = []
        for group in root.iter(f"{{{X14}}}sparklineGroup"):
            refs = [
                f"{s.findtext(f'{{{XM}}}sqref')}<-{s.findtext(f'{{{XM}}}f')}"
                for s in group.iter(f"{{{X14}}}sparkline")
            ]
            sparklines.append(f"{group.get('type', 'line')} {', '.join(refs)}")
        if sparklines:
            lines.append(f"sparkline groups: {'; '.join(sparklines)}")
        for member in book.related(sheet.member, "comments"):
            refs = [c.get("ref") for c in book.tree(member).iter(sml("comment"))]
            lines.append(f"legacy comments (VML): {', '.join(refs)}")
        for member in book.related(sheet.member, "threadedComment"):
            refs = [c.get("ref") for c in book.tree(member).iter(f"{{{TC}}}threadedComment")]
            lines.append(f"threaded comments: {', '.join(refs)}")
        for member in book.related(sheet.member, "pivotTable"):
            pivot = book.tree(member)
            location = pivot.find(sml("location"))
            ref = location.get("ref") if location is not None else "?"
            lines.append(f"pivot table: {pivot.get('name')} at {ref}")
        slicers = [
            s.get("name")
            for member in book.related(sheet.member, "slicer")
            for s in book.tree(member)
            if isinstance(s.tag, str)
        ]
        if slicers:
            lines.append(f"slicers: {', '.join(slicers)}")
        links = root.find(sml("hyperlinks"))
        if links is not None:
            lines.append(f"hyperlinks: {', '.join(h.get('ref', '') for h in links)}")
        if lines:
            found[sheet.name] = lines
    return found


_PART_KINDS = (
    ("vbaProject", "application/vnd.ms-office.vbaProject"),
    ("customXml", "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"),
    ("metadata", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"),
    ("calcChain", "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"),
    ("sharedStrings", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"),
    ("table", "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"),
    ("drawing", "application/vnd.openxmlformats-officedocument.drawing+xml"),
    ("chart", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"),
    ("vmlDrawing", "application/vnd.openxmlformats-officedocument.vmlDrawing"),
    ("comments", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"),
    ("threadedComments", "application/vnd.ms-excel.threadedcomments+xml"),
    ("persons", "application/vnd.ms-excel.person+xml"),
    ("pivotTable", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"),
    ("pivotCacheDefinition", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"),
    ("slicer", "application/vnd.ms-excel.slicer+xml"),
    ("slicerCache", "application/vnd.ms-excel.slicerCache+xml"),
    ("externalLink", "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"),
)


def package_parts(book: Book) -> dict[str, int]:
    """How many parts of each notable kind the package has, by content type."""
    counts: Counter[str] = Counter()
    for member, ctype in content_types(book.pkg).items():
        for label, wanted in _PART_KINDS:
            if ctype == wanted:
                counts[label] += 1
        if member.startswith("xl/media/"):
            counts["media"] += 1
    return dict(sorted(counts.items()))

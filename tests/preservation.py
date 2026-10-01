"""Preservation-verification harness and package consistency checker.

``assert_preserved`` reports exactly which package members an operation changed,
``worksheet_diff`` breaks one worksheet's change down to cells, rows and children,
and ``check_consistency`` enforces the cross-part invariants Excel repairs on open.
``checking_workbooks`` runs those checks on every workbook the test suite builds
(``tests/conftest.py``), honouring invariants a test declares with ``waives``.
Packages are ``read_pkg`` dictionaries of ZIP member name to bytes. Nothing here
imports pytest, so scripts can reuse the module.
"""

from __future__ import annotations

import functools
import io
import os
import posixpath
import re
import zipfile
from collections import Counter, defaultdict
from collections.abc import Iterable, Iterator, Mapping
from contextlib import contextmanager
from dataclasses import dataclass, field
from typing import NamedTuple
from urllib.parse import unquote

from lxml import etree

Package = Mapping[str, bytes]

CONTENT_TYPES = "[Content_Types].xml"
XML_EXTENSIONS = frozenset({"xml", "rels", "vml"})
MAX_ROW = 1_048_576
MAX_COL = 16_384

_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_SML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
_XM_NS = "http://schemas.microsoft.com/office/excel/2006/main"

_RT_OFFICE_DOCUMENT = f"{_R_NS}/officeDocument"
_RT_WORKSHEET = f"{_R_NS}/worksheet"
_RT_CALC_CHAIN = f"{_R_NS}/calcChain"
_RT_TABLE = f"{_R_NS}/table"
_RT_SHEET_METADATA = f"{_R_NS}/sheetMetadata"
_CT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

_R_ID = f"{{{_R_NS}}}id"


def _sml(localname: str) -> str:
    return f"{{{_SML_NS}}}{localname}"


_SHEET_DATA = _sml("sheetData")
_ROW = _sml("row")
_C = _sml("c")
_F = _sml("f")
_V = _sml("v")
_IS = _sml("is")
_T = _sml("t")
_R = _sml("r")

# entity resolution off: the harness parses whatever the library under test wrote
_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True)

# what xlsxedit's unpacked-folder reader skips (opc/phys_pkg.py)
_NOT_PACKAGE_DIRS = frozenset({".git", ".venv", "venv", "__pycache__"})
_NOT_PACKAGE_FILES = frozenset({".DS_Store", "Thumbs.db"})


def _ext(name: str) -> str:
    """OPC extension: after the last dot of the last segment, so ``_rels/.rels`` is ``rels``."""
    segment = name.rsplit("/", 1)[-1]
    return segment.rpartition(".")[2].lower() if "." in segment else ""


def _parse(blob: bytes) -> etree._Element:
    return etree.fromstring(blob, _PARSER)


def read_pkg(source) -> dict[str, bytes]:
    """Return ``{member name: bytes}`` for every file in a package.

    ``source`` is a path to a ZIP package or an unpacked package folder, the
    package bytes, a binary file-like, or an ``xlsxedit.Workbook`` (saved to
    memory first). Directory entries are skipped, and so are the folders and
    files the library itself skips in an unpacked package (``.git``,
    ``__pycache__``, ``.DS_Store``...). A member name that occurs twice raises
    ``ValueError``: no dictionary can represent that package.
    """
    from xlsxedit import Workbook

    if isinstance(source, Workbook):
        buf = io.BytesIO()
        source.save(buf)
        source = buf.getvalue()
    if isinstance(source, (str, os.PathLike)) and os.path.isdir(source):
        root = os.fspath(source)
        pkg: dict[str, bytes] = {}
        for folder, dirs, files in os.walk(root):
            dirs[:] = [d for d in dirs if d not in _NOT_PACKAGE_DIRS]
            for filename in files:
                if filename in _NOT_PACKAGE_FILES:
                    continue
                path = os.path.join(folder, filename)
                with open(path, "rb") as f:
                    pkg[os.path.relpath(path, root).replace(os.sep, "/")] = f.read()
        return pkg
    if isinstance(source, (bytes, bytearray, memoryview)):
        source = io.BytesIO(bytes(source))
    with zipfile.ZipFile(source) as z:
        names = [n for n in z.namelist() if not n.endswith("/")]
        duplicates = sorted(n for n, k in Counter(names).items() if k > 1)
        if duplicates:
            raise ValueError(f"duplicate ZIP members: {', '.join(duplicates)}")
        return {n: z.read(n) for n in names}


def _as_pkg(source) -> Package:
    return source if isinstance(source, Mapping) else read_pkg(source)


def canon(name: str, blob: bytes) -> bytes:
    """Canonical form of a package member, for equality comparison.

    XML-family members, classified by extension (``.xml``, ``.rels``, ``.vml``)
    rather than by whether the library models the part, become inclusive
    C14N 1.0 with comments. That forgives the XML declaration's quotes and
    lxml writing namespace declarations ahead of attributes, and nothing that
    changes the infoset: whitespace text inside ``<t xml:space="preserve">``,
    prefixes, comments, and namespace declarations that only ``mc:Ignorable``
    references all stay significant. (``etree.canonicalize`` is C14N 2.0,
    which drops unused declarations like exclusive C14N; it is not used here.)
    Other members, and XML that does not parse, are returned unchanged.
    """
    if _ext(name) not in XML_EXTENSIONS:
        return blob
    try:
        root = _parse(blob)
    except etree.XMLSyntaxError:
        return blob
    return etree.tostring(root.getroottree(), method="c14n", exclusive=False, with_comments=True)


def _well_formed(blob: bytes) -> bool:
    try:
        _parse(blob)
    except etree.XMLSyntaxError:
        return False
    return True


def content_types(pkg: Package) -> dict[str, str | None]:
    """Effective content type of every member: its ``Override``, else its extension's ``Default``.

    ``[Content_Types].xml`` itself is not a part and is left out; a member with
    neither entry maps to ``None``.
    """
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    if CONTENT_TYPES in pkg:
        root = _parse(pkg[CONTENT_TYPES])
        for elm in root.iterchildren(f"{{{_CT_NS}}}Default"):
            defaults[(elm.get("Extension") or "").lower()] = elm.get("ContentType")
        for elm in root.iterchildren(f"{{{_CT_NS}}}Override"):
            overrides[(elm.get("PartName") or "").lower()] = elm.get("ContentType")
    return {
        name: overrides.get(f"/{name}".lower(), defaults.get(_ext(name)))
        for name in pkg
        if name != CONTENT_TYPES
    }


@dataclass(frozen=True)
class PreservationReport:
    """Outcome of :func:`assert_preserved`.

    ``reserialised`` is informational: members whose bytes differ while their
    canonical form (for ``[Content_Types].xml``, the effective content-type
    map) is unchanged.
    """

    added: frozenset[str]
    removed: frozenset[str]
    changed: frozenset[str]
    reserialised: frozenset[str]


def assert_preserved(
    before,
    after,
    *,
    expected_changed: Iterable[str] = frozenset(),
    expected_added: Iterable[str] = frozenset(),
    expected_removed: Iterable[str] = frozenset(),
) -> PreservationReport:
    """Assert that exactly the expected members changed between two packages.

    ``before`` and ``after`` are ``read_pkg`` dictionaries or anything
    ``read_pkg`` accepts. Every common member is compared by :func:`canon`;
    ``[Content_Types].xml`` counts as changed only when a common member's
    effective content type differs, never because entries were rewritten.
    The changed, added and removed member sets must equal the expectations
    exactly. Changed and added XML members must be well-formed, and added
    members must have a content type.
    """
    before, after = _as_pkg(before), _as_pkg(after)
    expected_changed = frozenset(expected_changed)
    expected_added = frozenset(expected_added)
    expected_removed = frozenset(expected_removed)

    added = frozenset(after.keys() - before.keys())
    removed = frozenset(before.keys() - after.keys())
    common = before.keys() & after.keys()
    types_before, types_after = content_types(before), content_types(after)
    retyped = [
        (name, types_before[name], types_after[name])
        for name in sorted(common - {CONTENT_TYPES})
        if types_before[name] != types_after[name]
    ]

    changed: set[str] = set()
    reserialised: set[str] = set()
    for name in sorted(common):
        if name == CONTENT_TYPES:
            if retyped:
                changed.add(name)
            elif before[name] != after[name]:
                reserialised.add(name)
        elif before[name] != after[name]:
            if canon(name, before[name]) == canon(name, after[name]):
                reserialised.add(name)
            else:
                changed.add(name)

    problems = []
    for label, actual, expected in (
        ("changed", changed, expected_changed),
        ("added", added, expected_added),
        ("removed", removed, expected_removed),
    ):
        for name in sorted(actual - expected):
            problems.append(f"{label} but not expected: {name}")
        for name in sorted(expected - actual):
            where = "" if name in before or name in after else " (in neither package)"
            problems.append(f"expected {label} but was not: {name}{where}")
    for name in sorted(added - {CONTENT_TYPES}):
        if types_after[name] is None:
            problems.append(f"added member has no content type: {name}")
    for name in sorted(changed | added):
        if _ext(name) in XML_EXTENSIONS and not _well_formed(after[name]):
            problems.append(f"changed or added member is not well-formed XML: {name}")

    report = PreservationReport(added, removed, frozenset(changed), frozenset(reserialised))
    if problems:
        details = []
        for name in sorted(changed - expected_changed):
            if name == CONTENT_TYPES:
                details += [f"  {m}: {a} -> {b}" for m, a, b in retyped]
            else:
                details.append(_describe_change(name, before[name], after[name], types_after))
        message = ["package members differ from expectation:"]
        message += [f"  {p}" for p in problems]
        if details:
            message += ["unexpected changes:", *details]
        if reserialised:
            message.append(f"re-serialised, canonically equal: {', '.join(sorted(reserialised))}")
        raise AssertionError("\n".join(message))
    return report


def _describe_change(name: str, old: bytes, new: bytes, types: Mapping[str, str | None]) -> str:
    if types.get(name) == _CT_WORKSHEET:
        try:
            diff = worksheet_diff(old, new)
        except (etree.XMLSyntaxError, ValueError):
            diff = None
        if diff:  # else the change is one worksheet_diff does not model
            return f"  {name}:\n{diff.summary(indent='    ')}"
    a, b = canon(name, old), canon(name, new)
    at = next((i for i, (x, y) in enumerate(zip(a, b)) if x != y), min(len(a), len(b)))
    start = max(0, at - 60)
    return (
        f"  {name}: first difference at canonical offset {at}\n"
        f"    before: {a[start : at + 60].decode('utf-8', 'replace')}\n"
        f"    after:  {b[start : at + 60].decode('utf-8', 'replace')}"
    )


class CellState(NamedTuple):
    """What one ``<c>`` holds, as XML text.

    ``formula`` is the ``<f>`` text (``""`` for a shared-formula follower),
    ``None`` without ``<f>``; ``value`` is the ``<v>`` text, or the ``<t>``
    text of an inline string.
    """

    t: str | None
    s: str | None
    formula: str | None
    value: str | None


@dataclass
class WorksheetDiff:
    """Differences between two versions of a worksheet part.

    Cells are keyed by address and rows by number; ``changed`` holds
    ``(before, after)`` states. A cell compares by its whole ``<c>`` element,
    so a ``changed`` entry whose two states are equal differs in something
    the state leaves out (``<f>`` attributes such as ``ref``/``si``, ``cm``,
    ``vm``, ``ph``, rich-text runs). ``rows`` maps a row number to its
    ``(before, after)`` attributes, ``None`` where the row is absent.
    ``children`` maps ``(localname, ordinal)`` of every other root child to
    its ``(before, after)`` exclusive-C14N XML, so a namespace declaration
    added to the root shows once, under ``root``, instead of in every child.
    ``root`` and ``order`` are set when the root's attributes and namespace
    declarations, or the sequence of root-child names, differ. Order inside
    ``sheetData``, whitespace text and nodes outside the root element are not
    modelled; :func:`assert_preserved` still catches them.
    """

    added: dict[str, CellState] = field(default_factory=dict)
    removed: dict[str, CellState] = field(default_factory=dict)
    changed: dict[str, tuple[CellState, CellState]] = field(default_factory=dict)
    rows: dict[int, tuple[dict[str, str] | None, dict[str, str] | None]] = field(
        default_factory=dict
    )
    children: dict[tuple[str, int], tuple[str | None, str | None]] = field(default_factory=dict)
    root: tuple[dict[str, str], dict[str, str]] | None = None
    order: tuple[tuple[str, ...], tuple[str, ...]] | None = None
    _cell_xml: dict[str, tuple[str, str]] = field(default_factory=dict, repr=False, compare=False)

    def __bool__(self) -> bool:
        return bool(
            self.added
            or self.removed
            or self.changed
            or self.rows
            or self.children
            or self.root
            or self.order
        )

    def summary(self, indent: str = "", limit: int = 20) -> str:
        """Human-readable listing, at most ``limit`` entries per section."""
        lines: list[str] = []

        def section(title: str, entries: list[str]) -> None:
            if entries:
                lines.append(f"{indent}{title}:")
                lines.extend(f"{indent}  {e}" for e in entries[:limit])
                if len(entries) > limit:
                    lines.append(f"{indent}  ... {len(entries) - limit} more")

        section("cells added", [f"{a} {s}" for a, s in self.added.items()])
        section("cells removed", [f"{a} {s}" for a, s in self.removed.items()])
        changed = []
        for addr, (old, new) in self.changed.items():
            if old == new:
                old_xml, new_xml = self._cell_xml[addr]
                changed.append(f"{addr} {old_xml} -> {new_xml}")
            else:
                changed.append(f"{addr} {old} -> {new}")
        section("cells changed", changed)
        section("rows", [f"{n} {old} -> {new}" for n, (old, new) in self.rows.items()])
        section("children", [f"{k} {old} -> {new}" for k, (old, new) in self.children.items()])
        if self.root:
            lines.append(f"{indent}root: {self.root[0]} -> {self.root[1]}")
        if self.order:
            lines.append(f"{indent}child order: {self.order[0]} -> {self.order[1]}")
        return "\n".join(lines) if lines else f"{indent}(no differences)"


def child_order(blob: bytes) -> tuple[str, ...]:
    """Local names of the root element's children, in document order.

    For ``CT_Worksheet`` sequence assertions on a worksheet part.
    """
    return _child_names(_parse(blob))


def _child_names(root: etree._Element) -> tuple[str, ...]:
    return tuple(etree.QName(child).localname for child in root.iterchildren(etree.Element))


def worksheet_diff(before: bytes, after: bytes) -> WorksheetDiff:
    """Cell-, row- and child-level differences between two worksheet XML blobs."""
    old_root, new_root = _parse(before), _parse(after)
    diff = WorksheetDiff()

    old_cells, new_cells = _cell_map(old_root), _cell_map(new_root)
    for addr in sorted(old_cells.keys() | new_cells.keys(), key=_address_sort_key):
        if addr not in new_cells:
            diff.removed[addr] = _cell_state(old_cells[addr])
        elif addr not in old_cells:
            diff.added[addr] = _cell_state(new_cells[addr])
        else:
            old_xml, new_xml = _exclusive_xml(old_cells[addr]), _exclusive_xml(new_cells[addr])
            if old_xml != new_xml:
                diff.changed[addr] = (_cell_state(old_cells[addr]), _cell_state(new_cells[addr]))
                diff._cell_xml[addr] = (old_xml, new_xml)

    old_rows, new_rows = _row_map(old_root), _row_map(new_root)
    for number in sorted(old_rows.keys() | new_rows.keys()):
        old_attrs, new_attrs = old_rows.get(number), new_rows.get(number)
        if old_attrs != new_attrs:
            diff.rows[number] = (old_attrs, new_attrs)

    old_children, new_children = _child_map(old_root), _child_map(new_root)
    for key in sorted(old_children.keys() | new_children.keys()):
        old_xml, new_xml = old_children.get(key), new_children.get(key)
        if old_xml != new_xml:
            diff.children[key] = (old_xml, new_xml)

    old_head = (_attributes(old_root), dict(old_root.nsmap))
    new_head = (_attributes(new_root), dict(new_root.nsmap))
    if old_head != new_head:
        diff.root = (
            {**old_head[0], **{f"xmlns:{p}" if p else "xmlns": u for p, u in old_head[1].items()}},
            {**new_head[0], **{f"xmlns:{p}" if p else "xmlns": u for p, u in new_head[1].items()}},
        )
    old_order, new_order = _child_names(old_root), _child_names(new_root)
    if old_order != new_order:
        diff.order = (old_order, new_order)
    return diff


def _cell_state(c: etree._Element) -> CellState:
    f = c.find(_F)
    if c.get("t") == "inlineStr":
        is_elm = c.find(_IS)
        value = (
            None
            if is_elm is None
            else "".join(t.text or "" for t in is_elm.findall(_T) + is_elm.findall(f"{_R}/{_T}"))
        )
    else:
        v = c.find(_V)
        value = None if v is None else v.text
    return CellState(c.get("t"), c.get("s"), None if f is None else f.text or "", value)


def _exclusive_xml(elm: etree._Element) -> str:
    return etree.tostring(elm, method="c14n", exclusive=True, with_comments=True).decode("utf-8")


def _attributes(elm: etree._Element) -> dict[str, str]:
    prefixes = {uri: prefix for prefix, uri in elm.nsmap.items() if prefix}
    named = {}
    for key, value in elm.attrib.items():
        qname = etree.QName(key)
        prefix = prefixes.get(qname.namespace) if qname.namespace else None
        named[f"{prefix}:{qname.localname}" if prefix else key] = value
    return named


def _cell_map(root: etree._Element) -> dict[str, etree._Element]:
    cells: dict[str, etree._Element] = {}
    for addr, c in _cells(root):
        if addr in cells:
            raise ValueError(f"worksheet has two cells at {addr}")
        cells[addr] = c
    return cells


def _row_map(root: etree._Element) -> dict[int, dict[str, str]]:
    rows: dict[int, dict[str, str]] = {}
    for number, row in _rows(root):
        if number in rows:
            raise ValueError(f"worksheet has two rows numbered {number}")
        rows[number] = _attributes(row)
    return rows


def _child_map(root: etree._Element) -> dict[tuple[str, int], str]:
    children: dict[tuple[str, int], str] = {}
    seen: Counter[str] = Counter()
    for child in root.iterchildren(etree.Element):
        if child.tag == _SHEET_DATA:
            continue
        name = etree.QName(child).localname
        children[(name, seen[name])] = _exclusive_xml(child)
        seen[name] += 1
    return children


_CELL_RE = re.compile(r"^\$?([A-Za-z]+)\$?([0-9]+)$")
_REF_END_RE = re.compile(r"^\$?([A-Za-z]*)\$?([0-9]*)$")

Box = tuple[int, int, int, int]


def _col_index(letters: str) -> int:
    index = 0
    for ch in letters.upper():
        index = index * 26 + ord(ch) - 64
    return index


def _col_letters(index: int) -> str:
    letters = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _parse_cell(text: str | None) -> tuple[int, int] | None:
    """``(col, row)`` of an A1 cell reference (``$`` allowed), unbounded; None if malformed."""
    m = _CELL_RE.match(text or "")
    return (_col_index(m.group(1)), int(m.group(2))) if m else None


def _parse_ref(text: str | None) -> Box | None:
    """``(col1, row1, col2, row2)`` of a cell, range, whole-column or whole-row ref.

    Unbounded, so the caller can check the grid; None if malformed.
    """
    ends = (text or "").strip().split(":")
    if len(ends) > 2:
        return None
    parsed = []
    for end in ends:
        m = _REF_END_RE.match(end)
        if m is None:
            return None
        parsed.append((_col_index(m.group(1)) or None, int(m.group(2)) if m.group(2) else None))
    if len(parsed) == 1:
        col, row = parsed[0]
        return None if col is None or row is None else (col, row, col, row)
    (c1, r1), (c2, r2) = parsed
    if None not in (c1, r1, c2, r2):
        return min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2)
    if r1 is None and r2 is None and c1 is not None and c2 is not None:
        return min(c1, c2), 1, max(c1, c2), MAX_ROW
    if c1 is None and c2 is None and r1 is not None and r2 is not None:
        return 1, min(r1, r2), MAX_COL, max(r1, r2)
    return None


def _in_grid(box: Box) -> bool:
    c1, r1, c2, r2 = box
    return 1 <= c1 <= c2 <= MAX_COL and 1 <= r1 <= r2 <= MAX_ROW


def _contains(box: Box, col: int, row: int) -> bool:
    return box[0] <= col <= box[2] and box[1] <= row <= box[3]


def _format_ref(box: Box) -> str:
    first = f"{_col_letters(box[0])}{box[1]}"
    last = f"{_col_letters(box[2])}{box[3]}"
    return first if first == last else f"{first}:{last}"


def _address_sort_key(addr: str) -> tuple[int, int, str]:
    pos = _parse_cell(addr)
    return (pos[1], pos[0], "") if pos else (MAX_ROW + 1, 0, addr)


def _sqref_key(text: str | None) -> tuple:
    """Order-insensitive identity of a space-separated sqref."""
    boxes = [_parse_ref(token) or token.upper() for token in (text or "").split()]
    return tuple(sorted(boxes, key=repr))


def _rows(root: etree._Element) -> Iterator[tuple[int, etree._Element]]:
    """``(row number, <row>)``, numbering rows without ``r`` after their predecessor."""
    sheet_data = root.find(_SHEET_DATA)
    if sheet_data is None:
        return
    number = 0
    for row in sheet_data.iterchildren(_ROW):
        r = row.get("r")
        number = int(r) if r is not None and r.isdigit() else number + 1
        yield number, row


def _cells(root: etree._Element) -> Iterator[tuple[str, etree._Element]]:
    """``(address, <c>)`` in document order, placing cells without ``r`` after their predecessor.

    Addresses are uppercase without ``$``; an unparseable ``r`` is yielded as written.
    """
    for number, row in _rows(root):
        col = 0
        for c in row.iterchildren(_C):
            r = c.get("r")
            pos = _parse_cell(r)
            if pos is not None:
                col = pos[0]
                yield f"{_col_letters(pos[0])}{pos[1]}", c
            elif r is not None:
                yield r, c
            else:
                col += 1
                yield f"{_col_letters(col)}{number}", c


INVARIANTS: dict[str, str] = {
    "relationship": (
        "every internal relationship target is in the package, and every sheet and "
        "tablePart r:id names a relationship of its kind"
    ),
    "shared-formula": (
        "every shared-formula si has exactly one master carrying ref and text, "
        "and every cell with that si lies inside the master's ref"
    ),
    "formula-ref": "every array/dataTable formula master lies inside its own ref",
    "calc-chain": "every calcChain entry names a cell with <f> on a worksheet whose sheetId exists",
    "grid-bounds": (
        "every row, cell and reference in worksheets and tables parses and lies "
        "within A1:XFD1048576"
    ),
    "table-columns": "tableColumns/@count == number of tableColumn children == ref column width",
    "table-name": "table displayNames are unique workbook-wide (case-insensitively)",
    "table-autofilter": "table autoFilter/@ref == table ref minus totalsRowCount trailing rows",
    "cell-metadata": "every cell cm/vm index resolves to a bk of the workbook's metadata part",
    "x14-cf": "every cfRule x14:id has an x14:cfRule twin whose xm:sqref equals the base sqref",
}


class Violation(NamedTuple):
    """One broken invariant: its ``INVARIANTS`` code, the member it was found in, and why."""

    code: str
    part: str
    detail: str

    def __str__(self) -> str:
        return f"[{self.code}] {self.part}: {self.detail}"


def check_consistency(pkg, *, waive: Iterable[str] = ()) -> list[Violation]:
    """Assert that a package satisfies every invariant in ``INVARIANTS``.

    ``pkg`` is a ``read_pkg`` dictionary or anything ``read_pkg`` accepts.
    Violations of the codes in ``waive`` are tolerated and returned; any
    other violation raises ``AssertionError`` listing all of them.
    """
    waive = frozenset(waive)
    unknown = waive - INVARIANTS.keys()
    if unknown:
        raise ValueError(f"unknown invariant codes: {', '.join(sorted(unknown))}")
    found = consistency_violations(_as_pkg(pkg))
    failing = [v for v in found if v.code not in waive]
    if failing:
        lines = [f"package violates {len(failing)} consistency invariant(s):"]
        lines += [f"  {v}" for v in failing]
        raise AssertionError("\n".join(lines))
    return found


def worksheet_members(pkg) -> dict[str, str]:
    """Sheet name -> worksheet member, in workbook order, for building expectations."""
    return {sheet.name: sheet.member for sheet in _Book(_as_pkg(pkg)).worksheets()}


def consistency_violations(pkg: Package) -> list[Violation]:
    """Every invariant violation in ``pkg``, without raising."""
    book = _Book(pkg)
    for check in (
        _check_relationships,
        _check_formulas,
        _check_calc_chain,
        _check_grid_bounds,
        _check_tables,
        _check_cell_metadata,
        _check_x14_cf,
    ):
        check(book)
    return book.violations


def _rels_member(part: str) -> str:
    folder, filename = posixpath.split(part)
    return posixpath.join(folder, "_rels", f"{filename}.rels")


def _summarise(items: list[str], limit: int = 10) -> str:
    head = ", ".join(items[:limit])
    return head if len(items) <= limit else f"{head}, ... ({len(items)} in all)"


class _Sheet(NamedTuple):
    name: str
    sheet_id: str | None
    member: str | None


class _Book:
    """Read-only navigation of a SpreadsheetML package for the consistency checks."""

    def __init__(self, pkg: Package):
        self.pkg = pkg
        self.violations: list[Violation] = []
        self._trees: dict[str, etree._Element] = {}
        self._cells: dict[str, list[tuple[str, etree._Element]]] = {}
        self._tables: dict[str, list[str]] = {}
        package_rels = self.relationships("")
        self.workbook = next(
            (t for typ, t in package_rels.values() if typ == _RT_OFFICE_DOCUMENT and t), None
        )
        if self.workbook is None or self.workbook not in pkg:
            raise ValueError("package has no workbook part")
        self.workbook_rels = self.relationships(self.workbook)
        self.sheets: list[_Sheet] = []
        for elm in self.tree(self.workbook).iterfind(f"{_sml('sheets')}/{_sml('sheet')}"):
            name, r_id = elm.get("name") or "", elm.get(_R_ID)
            reltype, target = self.workbook_rels.get(r_id, (None, None))
            member = None
            if reltype == _RT_WORKSHEET:
                member = self.member(target)
            elif reltype is None:
                self.flag(
                    "relationship",
                    self.workbook,
                    f"sheet {name!r} r:id {r_id!r} has no relationship",
                )
            self.sheets.append(_Sheet(name, elm.get("sheetId"), member))

    def flag(self, code: str, part: str, detail: str) -> None:
        self.violations.append(Violation(code, part, detail))

    def tree(self, member: str) -> etree._Element:
        if member not in self._trees:
            self._trees[member] = _parse(self.pkg[member])
        return self._trees[member]

    def cells(self, member: str) -> list[tuple[str, etree._Element]]:
        """``(address, <c>)`` of a worksheet part, computed once."""
        if member not in self._cells:
            self._cells[member] = list(_cells(self.tree(member)))
        return self._cells[member]

    def relationships(self, source: str) -> dict[str, tuple[str, str | None]]:
        """``rId -> (type, target member)``; the target is None for external targets.

        Targets resolve like the library's ``PackURI.from_rel_ref``: absolute
        ones from the package root, and ``..`` never climbs above it.
        """
        blob = self.pkg.get(_rels_member(source))
        if blob is None:
            return {}
        rels = {}
        base = "/" + posixpath.dirname(source)
        for elm in _parse(blob).iterchildren(f"{{{_PKG_REL_NS}}}Relationship"):
            member = None
            if elm.get("TargetMode") != "External":
                target = elm.get("Target") or ""
                member = posixpath.normpath(posixpath.join(base, target)).lstrip("/")
            rels[elm.get("Id")] = (elm.get("Type"), member)
        return rels

    def member(self, name: str | None) -> str | None:
        """``name``, or its percent-decoded form, if the package has that member."""
        if name is None:
            return None
        if name in self.pkg:
            return name
        decoded = unquote(name)
        return decoded if decoded in self.pkg else None

    def workbook_part(self, reltype: str) -> str | None:
        for typ, target in self.workbook_rels.values():
            if typ == reltype:
                return self.member(target)
        return None

    def worksheets(self) -> Iterator[_Sheet]:
        return (s for s in self.sheets if s.member is not None)

    def tables(self, sheet: _Sheet) -> list[str]:
        """Table parts the worksheet's ``tableParts`` reference (resolved once)."""
        if sheet.member in self._tables:
            return self._tables[sheet.member]
        found = self._tables[sheet.member] = []
        rels = self.relationships(sheet.member)
        for elm in self.tree(sheet.member).iterfind(f"{_sml('tableParts')}/{_sml('tablePart')}"):
            r_id = elm.get(_R_ID)
            reltype, target = rels.get(r_id, (None, None))
            if reltype != _RT_TABLE:
                self.flag(
                    "relationship",
                    sheet.member,
                    f"tablePart r:id {r_id!r} is not a table relationship",
                )
                continue
            member = self.member(target)
            if member is not None:
                found.append(member)
        return found


def _check_relationships(book: _Book) -> None:
    for rels in sorted(m for m in book.pkg if _ext(m) == "rels"):
        folder, filename = posixpath.split(rels)
        if posixpath.basename(folder) != "_rels":
            continue
        source = posixpath.join(posixpath.dirname(folder), filename[: -len(".rels")])
        if source and source not in book.pkg:
            continue  # relationships of a part the package does not have
        missing = [
            f"{r_id} ({(typ or '?').rsplit('/', 1)[-1]}) targets missing part {target!r}"
            for r_id, (typ, target) in book.relationships(source).items()
            if target is not None and book.member(target) is None
        ]
        if missing:
            book.flag("relationship", rels, _summarise(missing))


def _check_formulas(book: _Book) -> None:
    for sheet in book.worksheets():
        groups: dict[str, list[tuple[str, etree._Element]]] = defaultdict(list)
        for addr, c in book.cells(sheet.member):
            f = c.find(_F)
            if f is None:
                continue
            kind = f.get("t")
            if kind == "shared":
                si = f.get("si")
                if si is None:
                    book.flag("shared-formula", sheet.member, f"{addr}: shared formula without si")
                else:
                    groups[si].append((addr, f))
            elif kind in ("array", "dataTable"):
                box, pos = _parse_ref(f.get("ref")), _parse_cell(addr)
                if box is None or pos is None or not _contains(box, *pos):
                    book.flag(
                        "formula-ref",
                        sheet.member,
                        f"{addr}: {kind} formula master lies outside its ref {f.get('ref')!r}",
                    )
        for si, members in groups.items():
            masters = [
                (a, f) for a, f in members if f.get("ref") is not None or (f.text or "").strip()
            ]
            if len(masters) != 1:
                book.flag(
                    "shared-formula",
                    sheet.member,
                    f"si={si}: {len(masters)} definitions ({_summarise([a for a, _ in masters])}); "
                    "expected exactly one master carrying ref and text",
                )
                continue
            master, f = masters[0]
            ref = f.get("ref")
            if ref is None or not (f.text or "").strip():
                missing = "ref" if ref is None else "formula text"
                book.flag(
                    "shared-formula", sheet.member, f"si={si}: master {master} has no {missing}"
                )
                continue
            box = _parse_ref(ref)
            if box is None:
                book.flag(
                    "shared-formula",
                    sheet.member,
                    f"si={si}: master {master} ref {ref!r} is malformed",
                )
                continue
            outside = [
                a for a, _ in members if (p := _parse_cell(a)) is None or not _contains(box, *p)
            ]
            if outside:
                book.flag(
                    "shared-formula",
                    sheet.member,
                    f"si={si}: {_summarise(outside)} outside master {master}'s ref {ref}",
                )


def _check_calc_chain(book: _Book) -> None:
    chain = book.workbook_part(_RT_CALC_CHAIN)
    if chain is None:
        return
    sheets = {s.sheet_id: s for s in book.sheets}
    formula_cells: dict[str, set[str]] = {}
    missing_sheets: dict[str, list[str]] = defaultdict(list)
    no_formula: dict[str, list[str]] = defaultdict(list)
    sheet_id = "0"  # schema default; an entry without i inherits the previous entry's
    for entry in book.tree(chain).iterchildren(_C):
        sheet_id = entry.get("i", sheet_id)
        pos = _parse_cell(entry.get("r"))
        addr = f"{_col_letters(pos[0])}{pos[1]}" if pos else entry.get("r") or "?"
        sheet = sheets.get(sheet_id)
        if sheet is None or sheet.member is None:
            missing_sheets[sheet_id].append(addr)
            continue
        if sheet.member not in formula_cells:
            formula_cells[sheet.member] = {
                a for a, c in book.cells(sheet.member) if c.find(_F) is not None
            }
        if addr not in formula_cells[sheet.member]:
            no_formula[sheet.name].append(addr)
    for sheet_id, addrs in missing_sheets.items():
        book.flag(
            "calc-chain",
            chain,
            f"entries {_summarise(addrs)} name sheetId {sheet_id}, which no worksheet has",
        )
    for name, addrs in no_formula.items():
        book.flag(
            "calc-chain", chain, f"entries name cells without <f> on {name!r}: {_summarise(addrs)}"
        )


_REF_ATTRIBUTES = ("ref", "sqref", "r1", "r2", "topLeftCell", "activeCell")


def _check_grid_bounds(book: _Book) -> None:
    parts = []
    for sheet in book.worksheets():
        parts.append(sheet.member)
        parts.extend(book.tables(sheet))
    for member in dict.fromkeys(parts):
        root = book.tree(member)
        bad: list[str] = []
        if root.tag == _sml("worksheet"):
            for number, row in _rows(root):
                r = row.get("r")
                if (r is not None and not r.isdigit()) or not 1 <= number <= MAX_ROW:
                    bad.append(f"row {r if r is not None else number}")
            for addr, _c in book.cells(member):
                pos = _parse_cell(addr)
                if pos is None or not _in_grid((*pos, *pos)):
                    bad.append(f"cell {addr}")
            for col in root.iterfind(f"{_sml('cols')}/{_sml('col')}"):
                lo, hi = col.get("min", ""), col.get("max", "")
                if not (lo.isdigit() and hi.isdigit() and 1 <= int(lo) <= int(hi) <= MAX_COL):
                    bad.append(f"col min={lo!r} max={hi!r}")
        for elm in _reference_carriers(root):
            local = etree.QName(elm).localname
            for attr in _REF_ATTRIBUTES:
                value = elm.get(attr)
                if value is not None:
                    bad += [f"{local}/@{attr} {t!r}" for t in _out_of_grid(value)]
            if elm.tag == f"{{{_XM_NS}}}sqref":
                bad += [f"xm:sqref {t!r}" for t in _out_of_grid(elm.text or "")]
        if bad:
            book.flag(
                "grid-bounds", member, f"outside A1:XFD1048576 or malformed: {_summarise(bad)}"
            )


def _reference_carriers(root: etree._Element) -> Iterator[etree._Element]:
    """The root, every element outside ``sheetData``, and the ``<f>`` elements inside it."""
    yield root
    for child in root.iterchildren(etree.Element):
        yield from child.iter(_F) if child.tag == _SHEET_DATA else child.iter(etree.Element)


def _out_of_grid(value: str) -> list[str]:
    tokens = value.split() or [value]
    return [t for t in tokens if (box := _parse_ref(t)) is None or not _in_grid(box)]


def _check_tables(book: _Book) -> None:
    owners: dict[str, list[str]] = defaultdict(list)
    for sheet in book.worksheets():
        for member in book.tables(sheet):
            table = book.tree(member)
            display_name = table.get("displayName")
            if display_name:
                owners[display_name.casefold()].append(f"{display_name} ({member})")
            else:
                book.flag("table-name", member, "table has no displayName")
            ref = table.get("ref")
            box = _parse_ref(ref)
            if box is None:
                book.flag("table-columns", member, f"table ref {ref!r} is malformed")
                continue
            columns = table.find(_sml("tableColumns"))
            children = 0 if columns is None else len(columns.findall(_sml("tableColumn")))
            count = None if columns is None else columns.get("count")
            width = box[2] - box[0] + 1
            if count is not None and not (count.isdigit() and int(count) == children):
                book.flag(
                    "table-columns",
                    member,
                    f"tableColumns/@count={count} but {children} tableColumn children",
                )
            if children != width:
                book.flag(
                    "table-columns",
                    member,
                    f"{children} tableColumn children but ref {ref} is {width} wide",
                )
            auto_filter = table.find(_sml("autoFilter"))
            if auto_filter is not None:
                totals = table.get("totalsRowCount", "0")
                totals = int(totals) if totals.isdigit() else 0
                expected = (box[0], box[1], box[2], box[3] - totals)
                if _parse_ref(auto_filter.get("ref")) != expected:
                    book.flag(
                        "table-autofilter",
                        member,
                        f"autoFilter ref {auto_filter.get('ref')!r} but table ref {ref} minus "
                        f"{totals} totals row(s) is {_format_ref(expected)}",
                    )
    for tables in owners.values():
        if len(tables) > 1:
            book.flag(
                "table-name",
                book.workbook,
                f"displayName used by {len(tables)} tables: {', '.join(tables)}",
            )


def _check_cell_metadata(book: _Book) -> None:
    metadata = book.workbook_part(_RT_SHEET_METADATA)
    limits = None
    if metadata is not None:
        root = book.tree(metadata)
        limits = {
            "cm": len(root.findall(f"{_sml('cellMetadata')}/{_sml('bk')}")),
            "vm": len(root.findall(f"{_sml('valueMetadata')}/{_sml('bk')}")),
        }
    for sheet in book.worksheets():
        bad = []
        for addr, c in book.cells(sheet.member):
            for attr in ("cm", "vm"):
                index = c.get(attr)
                if index is None:
                    continue
                if limits is None:
                    bad.append(f"{addr} {attr}={index} (no metadata part)")
                elif not (index.isdigit() and 1 <= int(index) <= limits[attr]):
                    bad.append(f"{addr} {attr}={index} (metadata has {limits[attr]})")
        if bad:
            book.flag("cell-metadata", sheet.member, f"unresolved indexes: {_summarise(bad)}")


def _check_x14_cf(book: _Book) -> None:
    for sheet in book.worksheets():
        root = book.tree(sheet.member)
        twins: dict[str, list[tuple]] = defaultdict(list)
        for x14_cf in root.iter(f"{{{_X14_NS}}}conditionalFormatting"):
            sqref = x14_cf.find(f"{{{_XM_NS}}}sqref")
            key = _sqref_key(None if sqref is None else sqref.text)
            for rule in x14_cf.iterchildren(f"{{{_X14_NS}}}cfRule"):
                if rule.get("id"):
                    twins[rule.get("id").strip().casefold()].append(key)
        for cf in root.iterchildren(_sml("conditionalFormatting")):
            base = _sqref_key(cf.get("sqref"))
            for rule in cf.iterchildren(_sml("cfRule")):
                for x14_id in rule.iter(f"{{{_X14_NS}}}id"):
                    rule_id = (x14_id.text or "").strip()
                    found = twins.get(rule_id.casefold())
                    if not found:
                        book.flag(
                            "x14-cf",
                            sheet.member,
                            f"cfRule x14:id {rule_id} has no x14:cfRule twin",
                        )
                    elif base not in found:
                        book.flag(
                            "x14-cf",
                            sheet.member,
                            f"x14:cfRule {rule_id} xm:sqref differs from base sqref "
                            f"{cf.get('sqref')!r}",
                        )


def waives(issue: str, *codes: str):
    """Mark a test whose workbooks still break the invariants ``codes`` until ``issue`` lands.

    The suite guard in ``tests/conftest.py`` tolerates those violations in that
    test, and fails it once none of them turns up any more, so the mark goes
    away with the fix. Every parametrisation of a marked test must break them.
    """
    unknown = set(codes) - INVARIANTS.keys()
    if not codes or unknown:
        raise ValueError(f"waives() needs invariant codes from INVARIANTS, got {codes!r}")

    def mark(test):
        test.consistency_waivers = {**waived(test), **dict.fromkeys(codes, issue)}
        return test

    return mark


def waived(test) -> dict[str, str]:
    """``{invariant code: issue}`` that :func:`waives` declared on a test function."""
    return dict(getattr(test, "consistency_waivers", None) or {})


class ConsistencyGuardError(Exception):
    """Raised by :func:`checking_workbooks`.

    Deliberately not an ``AssertionError``, so ``pytest.raises(AssertionError)``
    and ``xfail(raises=AssertionError)`` in a test cannot absorb it.
    """


class _Guard:
    abandoned = False

    def abandon(self) -> None:
        """Skip the exit checks: the code under test has already failed."""
        self.abandoned = True


@contextmanager
def checking_workbooks(*, waive: Mapping[str, str] | Iterable[str] = ()):
    """Run :func:`check_consistency` on every workbook built inside the block.

    ``Workbook.save`` checks the package it wrote (read back from a path or an
    in-memory buffer; other file objects are left alone). When the block
    exits, every ``Workbook`` constructed in it is saved to memory and checked
    again, so workbooks that are never saved are covered too; one that can no
    longer be saved is skipped. The block yields a guard whose ``abandon()``
    skips that exit pass, for when the code under test has already failed.
    ``waive`` holds invariant codes (or :func:`waived` ``{code: issue}``)
    whose violations are tolerated, but each must turn up at least once: a
    waiver that matches nothing is stale and fails the block. Failures raise
    :class:`ConsistencyGuardError`.
    """
    from xlsxedit import Workbook

    issues = dict(waive) if isinstance(waive, Mapping) else dict.fromkeys(waive)
    codes = frozenset(issues)
    seen: set[str] = set()
    built: list[Workbook] = []
    guard = _Guard()
    original_init, original_save = Workbook.__init__, Workbook.save

    def check(pkg) -> None:
        try:
            found = check_consistency(pkg, waive=codes)
        except AssertionError as exc:
            raise ConsistencyGuardError(str(exc)) from None
        seen.update(v.code for v in found)

    @functools.wraps(original_init)
    def init(self, *args, **kwargs):
        original_init(self, *args, **kwargs)
        built.append(self)

    @functools.wraps(original_save)
    def save(self, path, *args, **kwargs):
        original_save(self, path, *args, **kwargs)
        if isinstance(path, (str, os.PathLike)):
            check(read_pkg(path))
        elif hasattr(path, "getvalue") and hasattr(path, "tell"):
            check(read_pkg(path.getvalue()[: path.tell()]))

    Workbook.__init__, Workbook.save = init, save
    try:
        yield guard
    finally:
        Workbook.__init__, Workbook.save = original_init, original_save
    if guard.abandoned:
        return
    for wb in built:
        buf = io.BytesIO()
        try:
            wb.save(buf)
        except Exception:
            continue
        check(buf.getvalue())
    stale = sorted(codes - seen)
    if stale:
        named = ", ".join(f"{c} ({issues[c]})" if issues[c] else c for c in stale)
        raise ConsistencyGuardError(
            f"waived invariants are no longer violated: {named}; remove the waiver"
        )

"""Clone a worksheet part and its sheet-owned related parts."""

from __future__ import annotations

import re
import uuid
from copy import deepcopy

from xlsxedit.drawing import _A16_CREATION_ID, chart_references
from xlsxedit.opc.constants import CT, RT, SML_NS
from xlsxedit.opc.packuri import PackURI
from xlsxedit.opc.part import Part, XmlPart
from xlsxedit.parts import ChartPart, DrawingPart, WorksheetPart
from xlsxedit.row_shift import _FORMULA_TOKEN_RE, _same_sheet

_DEFINED_NAMES = f"{{{SML_NS}}}definedNames"
_DEFINED_NAME = f"{{{SML_NS}}}definedName"
_DIGITS_RE = re.compile(r"(\d+)(\.[^/]+)?$")
_PLAIN_SHEET_NAME = re.compile(r"[^\W\d][\w.]*")
# Unquoted, these would read as a cell, an R1C1 reference or a boolean.
_READS_AS_REFERENCE = re.compile(r"[a-z]{1,3}\d+|r\d*c?\d*|c\d*|true|false", re.IGNORECASE)
_DEFINED_NAME_TOKEN = re.compile(r"[^\W\d][\w.?\\]*|\\[\w.?\\]*")

_SHARE_RELTYPES = {RT.IMAGE}
_SHARE_CONTENT_TYPES = {CT.STYLES, CT.SHARED_STRINGS, CT.THEME, CT.WORKBOOK}


def _partname_template(partname: PackURI) -> str:
    s = str(partname)
    m = _DIGITS_RE.search(s)
    if m:
        return s[: m.start(1)] + "%d" + (m.group(2) or "")
    stem, ext = s.rsplit(".", 1) if "." in s.rsplit("/", 1)[-1] else (s, "")
    if ext:
        return f"{stem}%d.{ext}"
    return s + "%d"


def _should_share(rel) -> bool:
    if rel.is_external:
        return False
    if rel.reltype in _SHARE_RELTYPES:
        return True
    part = rel.target_part
    ct = getattr(part, "content_type", "") or ""
    if ct in _SHARE_CONTENT_TYPES or ct.startswith("image/"):
        return True
    return False


def _clone_part(package, source: Part, cloned: dict[int, Part], workbook) -> Part:
    key = id(source)
    if key in cloned:
        return cloned[key]

    partname = package.next_partname(_partname_template(source.partname))
    if isinstance(source, XmlPart):
        dest = type(source)(
            partname, source.content_type, deepcopy(source.element), package
        )
    else:
        dest = Part(partname, source.content_type, source.blob, package)
    package._add_part(dest)
    cloned[key] = dest
    _copy_relationships(source, dest, package, cloned, workbook)
    if source.content_type == CT.TABLE:
        _uniquify_table(dest, workbook)
    return dest


def _copy_relationships(source: Part, dest: Part, package, cloned: dict[int, Part], workbook) -> None:
    for rel in list(source.rels):
        if rel.is_external:
            dest.rels.add_relationship(rel.reltype, rel.target_ref, rel.rId, "External")
            continue
        if _should_share(rel):
            dest.rels.add_relationship(rel.reltype, rel.target_part, rel.rId)
            continue
        target = _clone_part(package, rel.target_part, cloned, workbook)
        dest.rels.add_relationship(rel.reltype, target, rel.rId)


def _uniquify_table(part: XmlPart, workbook) -> None:
    """Give a cloned table part an unused id and name, and fresh revision uids."""
    from xlsxedit.workbook import _next_table_name, _refresh_revision_uids, _table_names_and_ids

    elm = part.element
    used_names, used_ids = _table_names_and_ids(workbook, exclude=part)
    new_id = max(used_ids, default=0) + 1
    name = _next_table_name(used_names, new_id)
    elm.set("id", str(new_id))
    elm.set("name", name)
    elm.set("displayName", name)
    _refresh_revision_uids(elm)


def clone_sheet_relationships(
    source_part: WorksheetPart, dest_part: WorksheetPart, workbook
) -> list[Part]:
    """Copy sheet rels onto ``dest_part`` (same rIds; share images, clone the rest);
    return the parts cloned."""
    cloned: dict[int, Part] = {}
    _copy_relationships(source_part, dest_part, workbook._package, cloned, workbook)
    return list(cloned.values())


def sheet_prefix(name: str) -> str:
    """``name`` as a formula writes it before ``!``: quoted unless it is a plain name
    that cannot read as a reference."""
    if _PLAIN_SHEET_NAME.fullmatch(name) and not _READS_AS_REFERENCE.fullmatch(name):
        return name
    return "'" + name.replace("'", "''") + "'"


def repoint_sheet_refs(text: str, old_name: str, new_name: str) -> str:
    """``text`` with every reference and sheet-scoped name qualified with sheet
    ``old_name`` qualified with ``new_name`` instead; other sheets, 3D spans and
    other workbooks stay."""
    tokens = list(_FORMULA_TOKEN_RE.finditer(text))
    out = [m.group(0) for m in tokens]
    for i, m in enumerate(tokens):
        if i and tokens[i - 1].group("bracket") is not None:
            continue  # a sheet of another workbook
        if m.group("qref") is not None:
            if _same_sheet(m.group("sheet"), old_name):
                out[i] = f"{sheet_prefix(new_name)}!{m.group('qref')}"
        elif (
            i + 2 < len(tokens)
            and out[i + 1] == "!"
            and _DEFINED_NAME_TOKEN.fullmatch(out[i + 2])
            and _same_sheet(out[i], old_name)
        ):
            out[i] = sheet_prefix(new_name)  # a name scoped to the sheet: Sheet1!Sales
    return "".join(out)


def refresh_cloned_drawings(cloned: list[Part], old_name: str, new_name: str) -> None:
    """Make the drawings and charts cloned for a copy of sheet ``old_name`` its own:
    a chart's references to ``old_name`` name ``new_name``, as Excel's copy plots
    the copy's cells, and every drawing object gets a fresh ``a16:creationId``.
    ``cNvPr/@id`` stays: it is unique within its drawing part, and connectors
    and form controls refer to it."""
    for part in cloned:
        if isinstance(part, ChartPart):
            for elm in chart_references(part.element):
                elm.text = repoint_sheet_refs(elm.text, old_name, new_name)
        elif isinstance(part, DrawingPart):
            for creation_id in part.element.iter(_A16_CREATION_ID):
                creation_id.set("id", f"{{{str(uuid.uuid4()).upper()}}}")


def local_defined_name_copies(
    workbook, source_index: int, dest_index: int, source_name: str, dest_name: str
) -> list:
    """Copies of the definedNames with ``localSheetId`` equal to ``source_index``,
    scoped to ``dest_index``, their references to sheet ``source_name`` naming
    ``dest_name`` instead, as Excel copies them; :func:`append_defined_names` adds
    them to the workbook."""
    block = workbook._workbook_part.element.find(_DEFINED_NAMES)
    if block is None:
        return []
    copies = []
    for elm in block.findall(_DEFINED_NAME):
        local_id = elm.get("localSheetId")
        if local_id is None or int(local_id) != source_index:
            continue
        clone = deepcopy(elm)
        clone.set("localSheetId", str(dest_index))
        if clone.text:
            clone.text = repoint_sheet_refs(clone.text, source_name, dest_name)
        copies.append(clone)
    return copies


def append_defined_names(workbook, names: list) -> None:
    """Append ``names`` from :func:`local_defined_name_copies` to the workbook's definedNames."""
    if names:
        workbook._workbook_part.element.find(_DEFINED_NAMES).extend(names)

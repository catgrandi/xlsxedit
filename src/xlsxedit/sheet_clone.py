"""Clone a worksheet part and its sheet-owned related parts."""

from __future__ import annotations

import re
from copy import deepcopy

from xlsxedit.opc.constants import CT, RT, SML_NS
from xlsxedit.opc.packuri import PackURI
from xlsxedit.opc.part import Part, XmlPart
from xlsxedit.opc.serialize import serialize_part_xml
from xlsxedit.oxml.parser import parse_xml
from xlsxedit.parts import WorksheetPart

_DEFINED_NAMES = f"{{{SML_NS}}}definedNames"
_DEFINED_NAME = f"{{{SML_NS}}}definedName"
_DIGITS_RE = re.compile(r"(\d+)(\.[^/]+)?$")

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


def _uniquify_table(part: Part, workbook) -> None:
    """Give a cloned table part an unused id and name, and fresh revision uids."""
    from xlsxedit.workbook import _next_table_name, _refresh_revision_uids, _table_names_and_ids

    elm = parse_xml(part.blob)
    used_names, used_ids = _table_names_and_ids(workbook, exclude=part)
    new_id = max(used_ids, default=0) + 1
    name = _next_table_name(used_names, new_id)
    elm.set("id", str(new_id))
    elm.set("name", name)
    elm.set("displayName", name)
    _refresh_revision_uids(elm)
    part._blob = serialize_part_xml(elm)


def clone_sheet_relationships(source_part: WorksheetPart, dest_part: WorksheetPart, workbook) -> None:
    """Copy sheet rels onto ``dest_part`` (same rIds; share images, clone the rest)."""
    cloned: dict[int, Part] = {}
    _copy_relationships(source_part, dest_part, workbook._package, cloned, workbook)


def copy_local_defined_names(workbook, source_index: int, dest_index: int) -> None:
    """Duplicate definedNames with ``localSheetId`` equal to ``source_index``."""
    wb_elm = workbook._workbook_part.element
    block = wb_elm.find(_DEFINED_NAMES)
    if block is None:
        return
    for elm in list(block.findall(_DEFINED_NAME)):
        local_id = elm.get("localSheetId")
        if local_id is None or int(local_id) != source_index:
            continue
        clone = deepcopy(elm)
        clone.set("localSheetId", str(dest_index))
        block.append(clone)

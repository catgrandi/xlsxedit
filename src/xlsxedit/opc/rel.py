"""OPC relationships for a package or part."""

from __future__ import annotations

from typing import Iterator

from lxml import etree

from xlsxedit.opc.constants import REL_NS
from xlsxedit.opc.packuri import PackURI
from xlsxedit.opc.serialize import serialize_part_xml
from xlsxedit.oxml.parser import parse_xml

_REL_TAG = f"{{{REL_NS}}}Relationship"
_RELS_TAG = f"{{{REL_NS}}}Relationships"


class Relationship:
    """One relationship from a source part (or the package) to a target."""

    def __init__(
        self,
        rId: str,
        reltype: str,
        target_ref: str,
        target_mode: str = "Internal",
        target_part: object | None = None,
        base_uri: str = "/",
    ):
        self.rId = rId
        self.reltype = reltype
        self.target_ref = target_ref
        self.target_mode = target_mode
        self._target_part = target_part
        self._base_uri = base_uri

    @property
    def is_external(self) -> bool:
        return self.target_mode == "External"

    @property
    def target_part(self):
        if self.is_external:
            raise ValueError("External relationships have no target part")
        return self._target_part

    @target_part.setter
    def target_part(self, part) -> None:
        self._target_part = part

    @property
    def target_partname(self) -> PackURI:
        if self.is_external:
            raise ValueError("External relationships have no target partname")
        if self._target_part is not None:
            return self._target_part.partname
        return PackURI.from_rel_ref(self._base_uri, self.target_ref)


class Relationships:
    """Collection of relationships for a package or part."""

    def __init__(self, base_uri: str = "/"):
        self._base_uri = base_uri
        self._rels: dict[str, Relationship] = {}
        self._xml: bytes | None = None
        self._source: tuple[bytes, dict[str, tuple[str, str, str]]] | None = None

    def __len__(self) -> int:
        return len(self._rels)

    def __iter__(self) -> Iterator[Relationship]:
        return iter(self._rels.values())

    def __getitem__(self, rId: str) -> Relationship:
        return self._rels[rId]

    def get(self, rId: str, default=None):
        return self._rels.get(rId, default)

    def add_relationship(
        self,
        reltype: str,
        target: object | str,
        rId: str,
        target_mode: str = "Internal",
    ) -> Relationship:
        if target_mode == "External":
            target_ref = str(target)
            target_part = None
        else:
            target_part = target
            target_ref = target.partname.relative_ref(self._base_uri)
        rel = Relationship(
            rId, reltype, target_ref, target_mode, target_part, self._base_uri
        )
        self._rels[rId] = rel
        return rel

    def get_or_add(self, reltype: str, target_part) -> str:
        """Return rId of existing (reltype, part) relationship, or add one."""
        for rel in self._rels.values():
            if rel.reltype == reltype and not rel.is_external and rel.target_part is target_part:
                return rel.rId
        r_id = self.next_rId()
        self.add_relationship(reltype, target_part, r_id)
        return r_id

    def get_or_add_ext_rel(self, reltype: str, target_ref: str) -> str:
        """Return rId of existing external (reltype, ref) relationship, or add one."""
        for rel in self._rels.values():
            if rel.reltype == reltype and rel.is_external and rel.target_ref == target_ref:
                return rel.rId
        r_id = self.next_rId()
        self.add_relationship(reltype, target_ref, r_id, target_mode="External")
        return r_id

    def part_with_reltype(self, reltype: str):
        for rel in self._rels.values():
            if rel.reltype == reltype and not rel.is_external:
                return rel.target_part
        raise KeyError(f"no relationship of type {reltype!r}")

    def parts_with_reltype(self, reltype: str) -> list:
        return [
            rel.target_part
            for rel in self._rels.values()
            if rel.reltype == reltype and not rel.is_external
        ]

    def next_rId(self) -> str:
        max_id = 0
        for r_id in self._rels:
            if r_id.startswith("rId") and r_id[3:].isdigit():
                max_id = max(max_id, int(r_id[3:]))
        return f"rId{max_id + 1}"

    def keep_source(self, srels: Relationships) -> None:
        """Write the rels item ``srels`` was read from while it still lists these relationships."""
        if srels._xml is not None:
            self._source = (srels._xml, srels._signature())

    def _signature(self) -> dict[str, tuple[str, str, str]]:
        """``rId -> (type, mode, target)``; the target is a part name unless external."""
        return {
            rel.rId: (
                rel.reltype,
                rel.target_mode,
                rel.target_ref if rel.is_external else str(rel.target_partname),
            )
            for rel in self._rels.values()
        }

    @property
    def xml(self) -> bytes:
        if self._source is not None and self._source[1] == self._signature():
            return self._source[0]
        root = etree.Element(_RELS_TAG, nsmap={None: REL_NS})
        for rel in self._rels.values():
            elm = etree.SubElement(root, _REL_TAG)
            elm.set("Id", rel.rId)
            elm.set("Type", rel.reltype)
            elm.set("Target", rel.target_ref)
            if rel.is_external:
                elm.set("TargetMode", "External")
        return serialize_part_xml(root)

    @classmethod
    def from_xml(cls, base_uri: str, xml: bytes | None) -> Relationships:
        rels = cls(base_uri)
        if xml is None:
            return rels
        rels._xml = xml
        root = parse_xml(xml)
        for elm in root.findall(_REL_TAG):
            rId = elm.get("Id")
            reltype = elm.get("Type")
            target_ref = elm.get("Target")
            target_mode = elm.get("TargetMode", "Internal")
            rels._rels[rId] = Relationship(
                rId, reltype, target_ref, target_mode, None, base_uri
            )
        return rels

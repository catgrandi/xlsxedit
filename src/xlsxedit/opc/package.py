"""OPC package: load, walk parts, save."""

from __future__ import annotations

from xlsxedit.opc.packuri import PACKAGE_URI, PackURI
from xlsxedit.opc.part import Part, PartFactory
from xlsxedit.opc.pkgreader import ContentTypeMap, PackageReader
from xlsxedit.opc.pkgwriter import PackageWriter
from xlsxedit.opc.rel import Relationships


class OpcPackage:
    """In-memory OPC package (ZIP of related parts)."""

    def __init__(self):
        self._rels = Relationships("/")
        self._parts: dict[PackURI, Part] = {}
        self._orphan_parts: dict[PackURI, Part] = {}
        self._content_types: ContentTypeMap | None = None
        self._large = False

    @classmethod
    def open(cls, pkg_file, *, large: bool = False) -> OpcPackage:
        pkg = cls()
        pkg._large = large
        reader = PackageReader.from_file(pkg_file)
        pkg._content_types = reader.content_types
        Unmarshaller.unmarshal(reader, pkg, PartFactory)
        return pkg

    def save(self, pkg_file, *, include_orphans: bool = False) -> None:
        parts = list(self.iter_parts())
        if include_orphans:
            parts.extend(self.iter_orphan_parts())
        PackageWriter.write(pkg_file, self.rels, parts, self._content_types)

    @property
    def rels(self) -> Relationships:
        return self._rels

    def load_rel(
        self, reltype: str, target: object | str, rId: str, is_external: bool = False
    ) -> None:
        mode = "External" if is_external else "Internal"
        self._rels.add_relationship(reltype, target, rId, mode)

    def iter_parts(self):
        return iter(self._parts.values())

    def iter_orphan_parts(self):
        return iter(self._orphan_parts.values())

    @property
    def orphan_partnames(self) -> tuple[str, ...]:
        return tuple(sorted(str(p) for p in self._orphan_parts))

    def get_parts_of_type(self, content_type: str) -> list[Part]:
        return [p for p in self._parts.values() if p.content_type == content_type]

    def main_document_part(self) -> Part:
        from xlsxedit.opc.constants import RT

        return self.rels.part_with_reltype(RT.OFFICE_DOCUMENT)

    def _add_part(self, part: Part) -> None:
        self._parts[part.partname] = part

    def _remove_part(self, part: Part) -> None:
        self._parts.pop(part.partname, None)

    def _add_orphan_part(self, part: Part) -> None:
        self._orphan_parts[part.partname] = part

    def next_partname(self, template: str) -> PackURI:
        """Return the next part name matching ``template`` (``%d`` suffix) that no part,
        orphans included, already uses; part names compare case-insensitively."""
        partnames = {name.lower() for name in (*self._parts, *self._orphan_parts)}
        for n in range(1, len(partnames) + 2):
            candidate = PackURI(template % n)
            if candidate.lower() not in partnames:
                return candidate
        raise RuntimeError(f"could not allocate partname from {template!r}")


class Unmarshaller:
    """Assemble Part objects and wire relationships from a PackageReader."""

    @staticmethod
    def unmarshal(pkg_reader: PackageReader, package: OpcPackage, part_factory):
        parts = Unmarshaller._unmarshal_parts(pkg_reader, package, part_factory)
        Unmarshaller._unmarshal_relationships(pkg_reader, package, parts)
        Unmarshaller._unmarshal_orphans(pkg_reader, package)
        for part in parts.values():
            part.after_unmarshal()

    @staticmethod
    def _unmarshal_parts(pkg_reader, package, part_factory):
        parts = {}
        for partname, content_type, reltype, blob, _srels in pkg_reader.iter_sparts():
            part = part_factory(partname, content_type, reltype, blob, package)
            parts[partname] = part
            package._add_part(part)
        return parts

    @staticmethod
    def _unmarshal_orphans(pkg_reader: PackageReader, package: OpcPackage) -> None:
        for partname, content_type, blob in pkg_reader.iter_orphan_sparts():
            # Always opaque — orphans are cargo, not live XML.
            part = Part.load(partname, content_type, blob, package)
            package._add_orphan_part(part)

    @staticmethod
    def _unmarshal_relationships(pkg_reader, package, parts):
        for source_uri, srel in pkg_reader.iter_srels():
            source = package if source_uri == PACKAGE_URI else parts[source_uri]
            if srel.is_external:
                source.load_rel(srel.reltype, srel.target_ref, srel.rId, is_external=True)
            else:
                target = parts[srel.target_partname]
                source.load_rel(srel.reltype, target, srel.rId)

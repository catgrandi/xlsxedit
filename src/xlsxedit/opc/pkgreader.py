"""Read an OPC package into serialized parts and relationships."""

from __future__ import annotations

from xlsxedit.opc.constants import CT, CT_NS
from xlsxedit.opc.packuri import CONTENT_TYPES_URI, PACKAGE_URI, PackURI
from xlsxedit.opc.phys_pkg import PhysPkgReader
from xlsxedit.opc.rel import Relationships
from xlsxedit.oxml.parser import parse_xml

_DEFAULT_TAG = f"{{{CT_NS}}}Default"
_OVERRIDE_TAG = f"{{{CT_NS}}}Override"

_OCTET_STREAM = "application/octet-stream"


def _is_bookkeeping_part(partname: PackURI) -> bool:
    """True for content types / relationships items (regenerated on save)."""
    if partname == CONTENT_TYPES_URI:
        return True
    member = partname.membername
    if "/_rels/" in f"/{member}" or member.startswith("_rels/"):
        return True
    if partname.filename.endswith(".rels"):
        return True
    return False


class ContentTypeMap:
    def __init__(self):
        self._defaults: dict[str, str] = {}
        self._overrides: dict[str, str] = {}

    def __getitem__(self, partname: PackURI) -> str:
        key = str(partname).lower()
        if key in self._overrides:
            return self._overrides[key]
        ext = partname.ext.lower()
        if ext in self._defaults:
            return self._defaults[ext]
        raise KeyError(f"no content type for partname {partname!r}")

    def content_type_for(self, partname: PackURI) -> str:
        """Return content type, falling back for unlisted orphan members."""
        try:
            return self[partname]
        except KeyError:
            ext = partname.ext.lower()
            if ext == "xml":
                return CT.XML
            if ext == "rels":
                return CT.OPC_RELATIONSHIPS
            return _OCTET_STREAM

    @classmethod
    def from_xml(cls, xml: bytes) -> ContentTypeMap:
        root = parse_xml(xml)
        ct_map = cls()
        for elm in root.findall(_DEFAULT_TAG):
            ct_map._defaults[elm.get("Extension").lower()] = elm.get("ContentType")
        for elm in root.findall(_OVERRIDE_TAG):
            ct_map._overrides[elm.get("PartName").lower()] = elm.get("ContentType")
        return ct_map


class PackageReader:
    """Load parts by walking the OPC relationship graph."""

    def __init__(
        self,
        content_types: ContentTypeMap,
        pkg_rels: Relationships,
        sparts,
        orphan_sparts=(),
    ):
        self.content_types = content_types
        self.pkg_rels = pkg_rels
        self._sparts = sparts
        self._orphan_sparts = tuple(orphan_sparts)

    @classmethod
    def from_file(cls, pkg_file) -> PackageReader:
        phys = PhysPkgReader(pkg_file)
        content_types = ContentTypeMap.from_xml(phys.content_types_xml)
        pkg_rels = Relationships.from_xml("/", phys.rels_xml_for(PACKAGE_URI))
        sparts = cls._load_parts(phys, pkg_rels, content_types)
        orphan_sparts = cls._load_orphans(phys, content_types, sparts)
        phys.close()
        return cls(content_types, pkg_rels, sparts, orphan_sparts)

    def iter_sparts(self):
        for s in self._sparts:
            yield s

    def iter_orphan_sparts(self):
        """Yield ``(partname, content_type, blob)`` for unrelated ZIP members."""
        yield from self._orphan_sparts

    def iter_srels(self):
        for rel in self.pkg_rels:
            yield (PACKAGE_URI, rel)
        for partname, _ct, _rt, _blob, srels in self._sparts:
            for rel in srels:
                yield (partname, rel)

    @classmethod
    def _load_parts(cls, phys, pkg_rels, content_types):
        sparts = []
        for partname, blob, reltype, srels in cls._walk(phys, pkg_rels):
            content_type = content_types[partname]
            sparts.append((partname, content_type, reltype, blob, srels))
        return sparts

    @classmethod
    def _load_orphans(cls, phys, content_types: ContentTypeMap, sparts):
        related = {partname for partname, *_ in sparts}
        orphans = []
        for partname in phys.iter_part_membernames():
            if _is_bookkeeping_part(partname):
                continue
            if partname in related:
                continue
            blob = phys.blob_for(partname)
            content_type = content_types.content_type_for(partname)
            orphans.append((partname, content_type, blob))
        return orphans

    @classmethod
    def _walk(cls, phys, srels, visited=None):
        if visited is None:
            visited = []
        for srel in srels:
            if srel.is_external:
                continue
            partname = srel.target_partname
            if partname in visited:
                continue
            visited.append(partname)
            part_srels = Relationships.from_xml(
                partname.baseURI, phys.rels_xml_for(partname)
            )
            blob = phys.blob_for(partname)
            yield (partname, blob, srel.reltype, part_srels)
            yield from cls._walk(phys, part_srels, visited)

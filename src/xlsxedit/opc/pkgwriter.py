"""Write an OPC package as a ZIP archive."""

from __future__ import annotations

from lxml import etree

from xlsxedit.opc.constants import CT, CT_NS
from xlsxedit.opc.packuri import CONTENT_TYPES_URI, PACKAGE_URI
from xlsxedit.opc.phys_pkg import PhysPkgWriter
from xlsxedit.oxml.parser import serialize_xml

_DEFAULT_CONTENT_TYPES = {
    ("rels", CT.OPC_RELATIONSHIPS),
    ("xml", CT.XML),
}


class PackageWriter:
    @staticmethod
    def write(pkg_file, pkg_rels, parts) -> None:
        parts = list(parts)
        phys = PhysPkgWriter(pkg_file)
        PackageWriter._write_content_types(phys, parts)
        phys.write(PACKAGE_URI.rels_uri, pkg_rels.xml)
        for part in parts:
            part.before_marshal()
            phys.write(part.partname, part.blob)
            if len(part.rels):
                phys.write(part.partname.rels_uri, part.rels.xml)
        phys.close()

    @staticmethod
    def _write_content_types(phys, parts) -> None:
        defaults: dict[str, str] = {
            "rels": CT.OPC_RELATIONSHIPS,
            "xml": CT.XML,
        }
        overrides: list[tuple[str, str]] = []
        for part in parts:
            ext = part.partname.ext.lower()
            ct = part.content_type
            if (ext, ct) in _DEFAULT_CONTENT_TYPES:
                defaults[ext] = ct
            else:
                overrides.append((str(part.partname), ct))

        nsmap = {None: CT_NS}
        root = etree.Element(f"{{{CT_NS}}}Types", nsmap=nsmap)
        for ext, ct in sorted(defaults.items()):
            elm = etree.SubElement(root, f"{{{CT_NS}}}Default")
            elm.set("Extension", ext)
            elm.set("ContentType", ct)
        for partname, ct in overrides:
            elm = etree.SubElement(root, f"{{{CT_NS}}}Override")
            elm.set("PartName", partname)
            elm.set("ContentType", ct)
        phys.write(CONTENT_TYPES_URI, serialize_xml(root))

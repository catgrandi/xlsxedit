"""Write an OPC package as a ZIP archive."""

from __future__ import annotations

from collections import Counter

from lxml import etree

from xlsxedit.opc.constants import CT, CT_NS
from xlsxedit.opc.packuri import CONTENT_TYPES_URI, PACKAGE_URI
from xlsxedit.opc.phys_pkg import PhysPkgWriter
from xlsxedit.opc.pkgreader import ContentTypeMap
from xlsxedit.oxml.parser import serialize_xml


class PackageWriter:
    @staticmethod
    def write(pkg_file, pkg_rels, parts, content_types: ContentTypeMap | None = None) -> None:
        """Write ``parts`` and their relationships as a ZIP package.

        ``content_types`` is the source package's map: its bytes are written
        back while they still type every part exactly, otherwise its
        ``Default`` entries are kept and ``Override``s are added only where a
        part's type differs. Raises ``ValueError`` before writing anything if
        two items would share a name.
        """
        parts = list(parts)
        PackageWriter._check_unique_names(parts)
        phys = PhysPkgWriter(pkg_file)
        phys.write(CONTENT_TYPES_URI, PackageWriter._content_types_xml(parts, content_types))
        phys.write(PACKAGE_URI.rels_uri, pkg_rels.xml)
        for part in parts:
            part.before_marshal()
            phys.write(part.partname, part.blob)
            if len(part.rels):
                phys.write(part.partname.rels_uri, part.rels.xml)
        phys.close()

    @staticmethod
    def _check_unique_names(parts) -> None:
        """Part names are case-insensitive, so ``/xl/A.xml`` and ``/xl/a.xml`` collide too."""
        names = [CONTENT_TYPES_URI, PACKAGE_URI.rels_uri]
        for part in parts:
            names.append(part.partname)
            if len(part.rels):
                names.append(part.partname.rels_uri)
        duplicates = sorted(n for n, k in Counter(n.lower() for n in names).items() if k > 1)
        if duplicates:
            raise ValueError(f"package would contain duplicate part names: {', '.join(duplicates)}")

    @staticmethod
    def _content_types_xml(parts, source: ContentTypeMap | None) -> bytes:
        if source is not None and source.xml is not None and source.describes(parts):
            return source.xml
        defaults: dict[str, tuple[str, str]] = {}
        for ext, ct in source.default_entries if source is not None else ():
            defaults.setdefault(ext.lower(), (ext, ct))
        rels_ext = defaults.get("rels", ("rels", ""))[0]
        defaults["rels"] = (rels_ext, CT.OPC_RELATIONSHIPS)
        defaults.setdefault("xml", ("xml", CT.XML))

        root = etree.Element(f"{{{CT_NS}}}Types", nsmap={None: CT_NS})
        for ext, ct in defaults.values():
            elm = etree.SubElement(root, f"{{{CT_NS}}}Default")
            elm.set("Extension", ext)
            elm.set("ContentType", ct)
        for part in parts:
            default = defaults.get(part.partname.ext.lower())
            if default is not None and default[1] == part.content_type:
                continue
            elm = etree.SubElement(root, f"{{{CT_NS}}}Override")
            elm.set("PartName", str(part.partname))
            elm.set("ContentType", part.content_type)
        return serialize_xml(root)

"""OPC Part and XmlPart."""

from __future__ import annotations

from typing import TYPE_CHECKING

from xlsxedit.opc.packuri import PackURI
from xlsxedit.opc.rel import Relationships
from xlsxedit.oxml.parser import parse_xml, serialize_xml

if TYPE_CHECKING:
    from lxml.etree import _Element


class Part:
    """A package part. Unknown content types keep the original blob."""

    def __init__(
        self,
        partname: PackURI,
        content_type: str,
        blob: bytes | None = None,
        package=None,
    ):
        self._partname = partname
        self._content_type = content_type
        self._blob = blob or b""
        self._package = package
        self._rels = Relationships(partname.baseURI)

    @classmethod
    def load(cls, partname: PackURI, content_type: str, blob: bytes, package=None):
        return cls(partname, content_type, blob, package)

    @property
    def partname(self) -> PackURI:
        return self._partname

    @property
    def content_type(self) -> str:
        return self._content_type

    @property
    def blob(self) -> bytes:
        return self._blob

    @property
    def package(self):
        return self._package

    @property
    def rels(self) -> Relationships:
        return self._rels

    def load_rel(
        self, reltype: str, target: object | str, rId: str, is_external: bool = False
    ) -> None:
        mode = "External" if is_external else "Internal"
        self._rels.add_relationship(reltype, target, rId, mode)

    def relate_to(self, target_part, reltype: str) -> str:
        """Relate this part to ``target_part``, reusing an existing rId if present."""
        return self._rels.get_or_add(reltype, target_part)

    def after_unmarshal(self) -> None:
        pass

    def before_marshal(self) -> None:
        pass


class XmlPart(Part):
    """Package part whose payload is a live lxml tree."""

    def __init__(
        self,
        partname: PackURI,
        content_type: str,
        element: "_Element",
        package=None,
    ):
        super().__init__(partname, content_type, None, package)
        self._element = element

    @classmethod
    def load(cls, partname: PackURI, content_type: str, blob: bytes, package=None):
        return cls(partname, content_type, parse_xml(blob), package)

    @property
    def element(self) -> "_Element":
        return self._element

    @property
    def blob(self) -> bytes:
        return serialize_xml(self._element)


class PartFactory:
    """Map content types to Part subclasses."""

    part_type_for: dict[str, type[Part]] = {}

    def __new__(
        cls,
        partname: PackURI,
        content_type: str,
        reltype: str | None,
        blob: bytes,
        package=None,
    ) -> Part:
        from xlsxedit.opc.constants import CT
        from xlsxedit.parts import WorksheetPart

        if getattr(package, "_large", False) and content_type == CT.WORKSHEET:
            return WorksheetPart.load_deferred(partname, content_type, blob, package)
        part_cls = cls.part_type_for.get(content_type, Part)
        return part_cls.load(partname, content_type, blob, package)

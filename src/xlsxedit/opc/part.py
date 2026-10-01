"""OPC Part and XmlPart."""

from __future__ import annotations

import hashlib
from typing import TYPE_CHECKING

from xlsxedit.opc.packuri import PackURI
from xlsxedit.opc.rel import Relationships
from xlsxedit.opc.serialize import serialize_part_xml
from xlsxedit.oxml.parser import parse_xml

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
    """Package part whose payload is a live lxml tree.

    A part loaded from a package keeps its source bytes and writes them back
    while its tree is canonically unchanged: never handed out by
    :attr:`element`, or serializing exactly as a fresh parse of the source
    bytes does. An untouched part therefore round-trips byte-identically. A
    part whose tree changed, or that was marked with :meth:`mark_dirty`, is
    serialized from the tree.
    """

    _orig_blob: bytes | None = None
    _orig_digest: bytes | None = None  # how a fresh parse of _orig_blob serializes
    _handed_out: bool = False
    _dirty: bool = False

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
        part = cls(partname, content_type, parse_xml(blob), package)
        part._orig_blob = blob
        return part

    @property
    def element(self) -> "_Element":
        self._handed_out = True
        return self._element

    def mark_dirty(self) -> None:
        """Serialize the tree on save without comparing it to the source bytes."""
        self._dirty = True

    @property
    def blob(self) -> bytes:
        if self._orig_blob is None or self._dirty:
            return serialize_part_xml(self.element)
        if not self._handed_out:
            return self._orig_blob
        xml = serialize_part_xml(self._element)
        if xml == self._orig_blob:
            return self._orig_blob
        if self._orig_digest is None:
            loaded = serialize_part_xml(parse_xml(self._orig_blob))
            self._orig_digest = hashlib.sha256(loaded).digest()
        return self._orig_blob if hashlib.sha256(xml).digest() == self._orig_digest else xml


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

"""PackURI value type for OPC part names."""

from __future__ import annotations

import posixpath


class PackURI(str):
    """Absolute pack URI such as ``/xl/workbook.xml``."""

    def __new__(cls, pack_uri_str: str) -> PackURI:
        if not pack_uri_str or pack_uri_str[0] != "/":
            raise ValueError(f"PackURI must begin with slash, got {pack_uri_str!r}")
        return str.__new__(cls, pack_uri_str)

    @staticmethod
    def from_rel_ref(base_uri: str, relative_ref: str) -> PackURI:
        joined = posixpath.join(base_uri, relative_ref)
        return PackURI(posixpath.abspath(joined))

    @property
    def baseURI(self) -> str:
        return posixpath.split(self)[0]

    @property
    def ext(self) -> str:
        raw = posixpath.splitext(self)[1]
        return raw[1:] if raw.startswith(".") else raw

    @property
    def filename(self) -> str:
        return posixpath.split(self)[1]

    @property
    def membername(self) -> str:
        return self[1:]

    def relative_ref(self, base_uri: str) -> str:
        if base_uri == "/":
            return self[1:]
        return posixpath.relpath(self, base_uri)

    @property
    def rels_uri(self) -> PackURI:
        if self == "/":
            return PackURI("/_rels/.rels")
        return PackURI(f"{self.baseURI}/_rels/{self.filename}.rels")


PACKAGE_URI = PackURI("/")
CONTENT_TYPES_URI = PackURI("/[Content_Types].xml")

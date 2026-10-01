"""PackURI value type for OPC part names."""

from __future__ import annotations

import posixpath
import re

# ECMA-376 Part 2 part-name segment: one or more pchar (RFC 3986), not ending in "."
_SEGMENT_RE = re.compile(r"(?:[A-Za-z0-9\-._~!$&'()*+,;=:@]|%[0-9A-Fa-f]{2})+")
_NOT_PCHAR_RE = re.compile(r"%(?![0-9A-Fa-f]{2})|[^A-Za-z0-9\-._~!$&'()*+,;=:@/%]")


def is_partname(name: str) -> bool:
    """True if ``name`` is a legal OPC part name such as ``/xl/workbook.xml``."""
    if not name.startswith("/"):
        return False
    return all(
        _SEGMENT_RE.fullmatch(segment) and not segment.endswith(".")
        for segment in name[1:].split("/")
    )


def encode_partname(name: str) -> str | None:
    """Percent-encode the characters that make ``name`` an illegal part name.

    ``/[trash]/0000.dat`` becomes ``/%5Btrash%5D/0000.dat``. Returns None when
    encoding cannot help (an empty segment, or one ending in ``.``).
    """
    encoded = _NOT_PCHAR_RE.sub(
        lambda m: "".join(f"%{b:02X}" for b in m.group().encode("utf-8")), name
    )
    return encoded if is_partname(encoded) else None


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

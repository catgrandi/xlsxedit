"""Physical OPC package I/O (ZIP and directory)."""

from __future__ import annotations

import os
from collections.abc import Iterator
from zipfile import ZIP_DEFLATED, ZipFile, is_zipfile

from xlsxedit.opc.packuri import CONTENT_TYPES_URI, PackURI

_JUNK_BASENAMES = {".DS_Store", "Thumbs.db"}


class PackageNotFoundError(OSError):
    """Raised when a package path cannot be opened."""


def _is_junk_member(membername: str) -> bool:
    """True for directory markers and OS junk files."""
    if not membername or membername.endswith("/"):
        return True
    base = membername.rsplit("/", 1)[-1]
    return base in _JUNK_BASENAMES


class PhysPkgReader:
    """Factory for physical package readers."""

    def __new__(cls, pkg_file):
        if isinstance(pkg_file, (str, os.PathLike)):
            path = os.fspath(pkg_file)
            if os.path.isdir(path):
                reader_cls = _DirPkgReader
            elif is_zipfile(path):
                reader_cls = _ZipPkgReader
            else:
                raise PackageNotFoundError(f"Package not found at {path!r}")
        else:
            reader_cls = _ZipPkgReader
        return super().__new__(reader_cls)

    def iter_part_membernames(self) -> Iterator[PackURI]:
        raise NotImplementedError


class PhysPkgWriter:
    """Factory for ZIP package writers."""

    def __new__(cls, pkg_file):
        return super().__new__(_ZipPkgWriter)


class _DirPkgReader(PhysPkgReader):
    def __init__(self, path):
        self._path = os.path.abspath(os.fspath(path))

    def blob_for(self, pack_uri: PackURI) -> bytes:
        path = os.path.join(self._path, pack_uri.membername)
        with open(path, "rb") as f:
            return f.read()

    def close(self) -> None:
        pass

    @property
    def content_types_xml(self) -> bytes:
        return self.blob_for(CONTENT_TYPES_URI)

    def rels_xml_for(self, source_uri: PackURI) -> bytes | None:
        try:
            return self.blob_for(source_uri.rels_uri)
        except OSError:
            return None

    def iter_part_membernames(self) -> Iterator[PackURI]:
        for root, dirs, files in os.walk(self._path):
            dirs[:] = [d for d in dirs if d not in {".git", ".venv", "venv", "__pycache__"}]
            for name in files:
                path = os.path.join(root, name)
                rel = os.path.relpath(path, self._path).replace(os.sep, "/")
                if _is_junk_member(rel):
                    continue
                yield PackURI("/" + rel)


class _ZipPkgReader(PhysPkgReader):
    def __init__(self, pkg_file):
        self._zipf = ZipFile(pkg_file, "r")

    def blob_for(self, pack_uri: PackURI) -> bytes:
        return self._zipf.read(pack_uri.membername)

    def close(self) -> None:
        self._zipf.close()

    @property
    def content_types_xml(self) -> bytes:
        return self.blob_for(CONTENT_TYPES_URI)

    def rels_xml_for(self, source_uri: PackURI) -> bytes | None:
        try:
            return self.blob_for(source_uri.rels_uri)
        except KeyError:
            return None

    def iter_part_membernames(self) -> Iterator[PackURI]:
        for name in self._zipf.namelist():
            if _is_junk_member(name):
                continue
            yield PackURI("/" + name if not name.startswith("/") else name)


class _ZipPkgWriter(PhysPkgWriter):
    def __init__(self, pkg_file):
        self._zipf = ZipFile(pkg_file, "w", compression=ZIP_DEFLATED)

    def close(self) -> None:
        self._zipf.close()

    def write(self, pack_uri: PackURI, blob: bytes) -> None:
        self._zipf.writestr(pack_uri.membername, blob)

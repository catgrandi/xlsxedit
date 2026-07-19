"""Read image pixel dimensions and fit sizes preserving aspect ratio."""

from __future__ import annotations

import struct
from pathlib import Path

from xlsxedit.exceptions import InvalidImageError


def read_pixel_size(path: str | Path) -> tuple[int, int]:
    """Return ``(width, height)`` for a JPEG or PNG file."""
    data = Path(path).read_bytes()
    if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
        width, height = struct.unpack(">II", data[16:24])
        return int(width), int(height)
    if data.startswith(b"\xff\xd8"):
        return _jpeg_pixel_size(data)
    raise InvalidImageError(f"unsupported or invalid image file: {path}")


def _jpeg_pixel_size(data: bytes) -> tuple[int, int]:
    i = 2
    while i < len(data) - 8:
        if data[i] != 0xFF:
            raise InvalidImageError("invalid JPEG data")
        marker = data[i + 1]
        if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
            height = int.from_bytes(data[i + 5 : i + 7], "big")
            width = int.from_bytes(data[i + 7 : i + 9], "big")
            return width, height
        length = int.from_bytes(data[i + 2 : i + 4], "big")
        if length < 2:
            raise InvalidImageError("invalid JPEG segment length")
        i += 2 + length
    raise InvalidImageError("JPEG size marker not found")


def fit_pixel_size(
    width: int,
    height: int,
    *,
    max_width: int = 200,
    max_height: int = 200,
) -> tuple[int, int]:
    """Scale ``(width, height)`` down to fit inside the max box, preserving aspect ratio."""
    if width < 1 or height < 1:
        raise ValueError("width and height must be positive")
    if max_width < 1 or max_height < 1:
        raise ValueError("max_width and max_height must be positive")
    scale = min(max_width / width, max_height / height, 1.0)
    fitted_w = max(1, round(width * scale))
    fitted_h = max(1, round(height * scale))
    return fitted_w, fitted_h


def resolve_display_size(
    path: str | Path,
    *,
    width: int | None = None,
    height: int | None = None,
    max_width: int = 200,
    max_height: int = 200,
) -> tuple[int, int]:
    """Resolve display width/height for ``add_image``."""
    natural_w, natural_h = read_pixel_size(path)
    if width is not None and height is not None:
        return width, height
    if width is None and height is None:
        return fit_pixel_size(natural_w, natural_h, max_width=max_width, max_height=max_height)
    if width is not None:
        scale = width / natural_w
        return width, max(1, round(natural_h * scale))
    assert height is not None
    scale = height / natural_h
    return max(1, round(natural_w * scale)), height

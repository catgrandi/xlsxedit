"""Tests for image pixel size helpers."""

from __future__ import annotations

from pathlib import Path

import pytest

from xlsxedit import Workbook
from xlsxedit.image_size import fit_pixel_size, read_pixel_size, resolve_display_size
from tests.conftest import ASSETS


@pytest.fixture
def coco_jpg() -> Path:
    path = ASSETS / "coco-happy-swiss-nature.jpg"
    if not path.is_file():
        pytest.skip("bundled coco JPEG not found")
    return path


def test_read_pixel_size_portrait(coco_jpg: Path):
    w, h = read_pixel_size(coco_jpg)
    assert w == 1106
    assert h == 1478
    assert h > w


def test_fit_pixel_size_preserves_aspect_ratio():
    fitted_w, fitted_h = fit_pixel_size(1106, 1478, max_width=200, max_height=200)
    assert fitted_h == 200
    assert fitted_w < fitted_h
    ratio_natural = 1106 / 1478
    ratio_fitted = fitted_w / fitted_h
    assert ratio_fitted == pytest.approx(ratio_natural, rel=0.02)


def test_resolve_display_size_defaults(coco_jpg: Path):
    w, h = resolve_display_size(coco_jpg)
    assert h <= 200
    assert w <= 200
    assert h > w


def test_add_image_auto_fit_aspect_ratio(tmp_path: Path, coco_jpg: Path):
    wb = Workbook.create()
    pic = wb["Sheet1"].add_image(coco_jpg, anchor="C4", name="coco")
    assert pic.height >= pic.width
    natural_w, natural_h = read_pixel_size(coco_jpg)
    assert pic.width / pic.height == pytest.approx(natural_w / natural_h, rel=0.05)
    wb.save(tmp_path / "fit.xlsx")


def test_insert_image_at_placeholder_width_keeps_aspect(tmp_path: Path, coco_jpg: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "{logo}"
    n = wb.insert_image_at_placeholder("{logo}", coco_jpg, width=96)
    assert n == 1
    pic = ws.images[0]
    assert pic.width == 96
    natural_w, natural_h = read_pixel_size(coco_jpg)
    assert pic.height == max(1, round(natural_h * (96 / natural_w)))
    wb.save(tmp_path / "sar_w.xlsx")


def test_insert_image_at_placeholder_default_fits_box(tmp_path: Path, coco_jpg: Path):
    wb = Workbook.create()
    wb["Sheet1"]["B2"].value = "{logo}"
    assert wb.insert_image_at_placeholder("{logo}", coco_jpg) == 1
    pic = wb["Sheet1"].images[0]
    assert pic.width <= 200
    assert pic.height <= 200
    wb.save(tmp_path / "sar_default.xlsx")

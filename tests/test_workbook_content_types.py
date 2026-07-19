"""Open/save workbooks whose main part uses .xlsm / .xltx content types."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from xlsxedit import Workbook
from xlsxedit.opc.constants import CT

_WORKBOOK_PARTNAME = "/xl/workbook.xml"


def _retype_workbook(src: Path, dest: Path, content_type: str) -> None:
    """Copy ``src`` to ``dest``, swapping the workbook override content type."""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dest, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    CT.WORKBOOK.encode(), content_type.encode()
                )
            zout.writestr(item, data)


def _content_type_of_workbook(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        xml = z.read("[Content_Types].xml").decode()
    for ct in (
        CT.WORKBOOK_MACRO_ENABLED,
        CT.WORKBOOK_MACRO_ENABLED_TEMPLATE,
        CT.WORKBOOK_TEMPLATE,
        CT.WORKBOOK,
    ):
        if f'PartName="{_WORKBOOK_PARTNAME}" ContentType="{ct}"' in xml:
            return ct
    raise AssertionError("workbook override not found in [Content_Types].xml")


@pytest.mark.parametrize(
    ("suffix", "content_type"),
    [
        (".xlsm", CT.WORKBOOK_MACRO_ENABLED),
        (".xltx", CT.WORKBOOK_TEMPLATE),
        (".xltm", CT.WORKBOOK_MACRO_ENABLED_TEMPLATE),
    ],
)
def test_open_edit_save_preserves_content_type(tmp_path: Path, suffix, content_type):
    plain = tmp_path / "plain.xlsx"
    wb = Workbook.create()
    wb.worksheets[0]["A1"].value = "hello"
    wb.save(plain)

    variant = tmp_path / f"variant{suffix}"
    _retype_workbook(plain, variant, content_type)

    wb2 = Workbook.open(variant)
    assert wb2.worksheets[0]["A1"].value == "hello"
    wb2.worksheets[0]["B1"].value = 42

    saved = tmp_path / f"saved{suffix}"
    wb2.save(saved)
    assert _content_type_of_workbook(saved) == content_type

    wb3 = Workbook.open(saved)
    assert wb3.worksheets[0]["A1"].value == "hello"
    assert wb3.worksheets[0]["B1"].value == 42


def test_pandas_writer_accepts_xlsm_extension():
    pd = pytest.importorskip("pandas")
    _ = pd
    from xlsxedit.pandas_io import XlsxeditWriter

    assert set(XlsxeditWriter._supported_extensions) == {".xlsx", ".xlsm", ".xltx"}

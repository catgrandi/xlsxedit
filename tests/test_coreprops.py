"""Tests for Workbook.properties (docProps/core.xml)."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

from xlsxedit import Workbook


def test_read_template_properties():
    wb = Workbook.create()
    props = wb.properties
    assert props.author  # template ships with a creator
    assert props.created is not None


def test_set_and_round_trip(tmp_path: Path):
    path = tmp_path / "props.xlsx"
    wb = Workbook.create()
    props = wb.properties
    props.title = "Quarterly report"
    props.author = "Test Author"
    props.subject = "Testing"
    props.keywords = "excel, test"
    props.comments = "A comment"
    props.category = "Reports"
    props.last_modified_by = "xlsxedit"
    props.created = datetime(2026, 1, 2, 3, 4, 5)
    props.modified = datetime(2026, 5, 6, 7, 8, 9, tzinfo=timezone.utc)
    wb.save(path)

    wb2 = Workbook.open(path)
    p2 = wb2.properties
    assert p2.title == "Quarterly report"
    assert p2.author == "Test Author"
    assert p2.subject == "Testing"
    assert p2.keywords == "excel, test"
    assert p2.comments == "A comment"
    assert p2.category == "Reports"
    assert p2.last_modified_by == "xlsxedit"
    assert p2.created == datetime(2026, 1, 2, 3, 4, 5, tzinfo=timezone.utc)
    assert p2.modified == datetime(2026, 5, 6, 7, 8, 9, tzinfo=timezone.utc)


def test_unset_fields_are_none():
    wb = Workbook.create()
    assert wb.properties.title is None
    assert wb.properties.subject is None


def test_missing_core_part_created(tmp_path: Path):
    # Simulate a package without docProps/core.xml by removing the rel + part.
    wb = Workbook.create()
    pkg = wb._package
    from xlsxedit.opc.constants import RT

    part = pkg.rels.part_with_reltype(RT.CORE_PROPERTIES)
    pkg._remove_part(part)
    for r_id, rel in list(pkg.rels._rels.items()):
        if rel.reltype == RT.CORE_PROPERTIES:
            del pkg.rels._rels[r_id]

    props = wb.properties  # creates a fresh part + relationship
    props.title = "Fresh"
    path = tmp_path / "fresh.xlsx"
    wb.save(path)

    wb2 = Workbook.open(path)
    assert wb2.properties.title == "Fresh"

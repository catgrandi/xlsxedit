from __future__ import annotations

from pathlib import Path

import pytest

from tests.preservation import checking_workbooks, waived

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"
ASSETS = Path(__file__).resolve().parents[1] / "assets"
BOOK1 = FIXTURES / "Book1.xlsx"

INSPECT_FIXTURES = {
    "TestBook": FIXTURES / "TestBook.xlsx",
    "MergedCells": FIXTURES / "MergedCells.xlsx",
    "DifferentCellTypes": FIXTURES / "DifferentCellTypes.xlsx",
    "CustomCellSize": FIXTURES / "CustomCellSize.xlsx",
    "ConditionalFormatting": FIXTURES / "ConditionalFormatting.xlsx",
    "SimpleFormula": FIXTURES / "SimpleFormula.xlsx",
    "images": FIXTURES / "images.xlsx",
    "ChartsAndTables": FIXTURES / "ChartsAndTables.xlsx",
    "TextSytle": FIXTURES / "TextSytle.xlsx",
    "hyperlink": FIXTURES / "hyperlink.xlsx",
}


@pytest.fixture(autouse=True)
def _consistent_workbooks(request):
    """Check every workbook a test builds or saves (``tests/preservation.py``).

    Waivers come from the test's ``waives()`` mark. A test that already failed
    skips the end-of-test pass; bench tests and tests that request
    ``unchecked_workbooks`` are left alone.
    """
    if request.node.get_closest_marker("bench") or "unchecked_workbooks" in request.fixturenames:
        yield
        return
    failures = request.session.testsfailed
    with checking_workbooks(waive=waived(getattr(request.node, "obj", None))) as guard:
        yield
        if request.session.testsfailed > failures:
            guard.abandon()


@pytest.fixture
def unchecked_workbooks() -> None:
    """Opt a test out of ``_consistent_workbooks``; it must check consistency itself."""


@pytest.fixture
def book1_path() -> Path:
    return BOOK1


@pytest.fixture
def fixtures_dir() -> Path:
    return FIXTURES


@pytest.fixture
def assets_dir() -> Path:
    return ASSETS

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
    # authored by Excel through scripts/excel_fixtures/build_fixtures.py
    "TableRich": FIXTURES / "TableRich.xlsx",
    "TableWithContentBelow": FIXTURES / "TableWithContentBelow.xlsx",
    "DataValidation": FIXTURES / "DataValidation.xlsx",
    "ConditionalFormattingRich": FIXTURES / "ConditionalFormattingRich.xlsx",
    "DrawingAnchors": FIXTURES / "DrawingAnchors.xlsx",
    "Formulas": FIXTURES / "Formulas.xlsx",
    "Preservation": FIXTURES / "Preservation.xlsm",
}

# Tests that a registered fixture makes fail until a roadmap issue lands, by node id:
# (exception, "#issue: reason"). Each becomes a strict xfail, so its entry has to go
# when the fix lands.
_NO_SHARED_STRINGS = (
    "#26 (no issue yet): writing text into a workbook without xl/sharedStrings.xml "
    "raises instead of adding the part; Excel writes none for a book with no text"
)
KNOWN_FIXTURE_FAILURES: dict[str, tuple[type[BaseException], str]] = {
    "tests/test_schema_validation.py::test_fixture_validates_after_insert_rows[DrawingAnchors]": (
        ValueError,
        _NO_SHARED_STRINGS,
    ),
    "tests/test_schema_validation.py::test_fixture_validates_after_insert_columns[DrawingAnchors]": (
        ValueError,
        _NO_SHARED_STRINGS,
    ),
}


def pytest_collection_modifyitems(config, items):
    for item in items:
        known = KNOWN_FIXTURE_FAILURES.get(item.nodeid)
        if known is not None:
            raises, reason = known
            item.add_marker(pytest.mark.xfail(strict=True, raises=raises, reason=reason))


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

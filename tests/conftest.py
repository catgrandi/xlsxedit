from __future__ import annotations

from pathlib import Path

import pytest

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


@pytest.fixture
def book1_path() -> Path:
    return BOOK1


@pytest.fixture
def fixtures_dir() -> Path:
    return FIXTURES


@pytest.fixture
def assets_dir() -> Path:
    return ASSETS

"""Pytest configuration and fixtures for rustypyxl tests."""

import os
import pytest
import rustypyxl


@pytest.fixture
def empty_workbook():
    """Create an empty workbook for testing."""
    return rustypyxl.Workbook()


@pytest.fixture
def workbook_with_sheet():
    """Create a workbook with one sheet."""
    wb = rustypyxl.Workbook()
    wb.create_sheet("Test")
    return wb


@pytest.fixture
def sample_xlsx_path():
    """Path to a sample xlsx file."""
    # Uses one of the test xlsx files in the project root
    path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "test_simple.xlsx")
    if os.path.exists(path):
        return path
    return None


@pytest.fixture
def temp_xlsx_path(tmp_path):
    """Temporary path for xlsx output."""
    return str(tmp_path / "test_output.xlsx")

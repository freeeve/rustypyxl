"""Tests for Cell functionality."""

import pytest
import rustypyxl


class TestCellBasics:
    """Test basic cell operations."""

    def test_cell_creation(self, workbook_with_sheet):
        """Create a cell with row and column."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        assert cell.row == 1
        assert cell.column == 1

    def test_cell_coordinate(self, workbook_with_sheet):
        """Test cell coordinate property."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        assert cell.coordinate == "A1"

        cell = ws.cell(10, 28)
        assert cell.coordinate == "AB10"

    def test_cell_column_letter(self, workbook_with_sheet):
        """Test cell column_letter property."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        assert cell.column_letter == "A"

        cell = ws.cell(1, 26)
        assert cell.column_letter == "Z"

        cell = ws.cell(1, 27)
        assert cell.column_letter == "AA"


class TestCellValue:
    """Test cell value operations."""

    def test_cell_value_none_default(self, workbook_with_sheet):
        """Cell value is None by default."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        assert cell.value is None

    def test_cell_set_string_value(self, workbook_with_sheet):
        """Set and get a string value."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        cell.value = "Hello"
        assert cell.value == "Hello"

    def test_cell_set_number_value(self, workbook_with_sheet):
        """Set and get a number value."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        cell.value = 42.5
        assert cell.value == 42.5

    def test_cell_set_boolean_value(self, workbook_with_sheet):
        """Set and get a boolean value."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        cell.value = True
        assert cell.value is True


class TestCellOffset:
    """Test cell offset method."""

    def test_offset_positive(self, workbook_with_sheet):
        """Offset by positive values."""
        ws = workbook_with_sheet.active
        cell = ws.cell(5, 3)
        offset_cell = cell.offset(2, 1)
        assert offset_cell.row == 7
        assert offset_cell.column == 4

    def test_offset_negative(self, workbook_with_sheet):
        """Offset by negative values."""
        ws = workbook_with_sheet.active
        cell = ws.cell(5, 3)
        offset_cell = cell.offset(-2, -1)
        assert offset_cell.row == 3
        assert offset_cell.column == 2

    def test_offset_clamps_to_minimum(self, workbook_with_sheet):
        """Offset clamps to row/column 1."""
        ws = workbook_with_sheet.active
        cell = ws.cell(2, 2)
        offset_cell = cell.offset(-10, -10)
        assert offset_cell.row >= 1
        assert offset_cell.column >= 1


class TestCellRepr:
    """Test cell string representation."""

    def test_cell_str(self, workbook_with_sheet):
        """Test cell string representation."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        s = str(cell)
        assert "Cell" in s
        assert "A1" in s

    def test_cell_repr(self, workbook_with_sheet):
        """Test cell repr."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        r = repr(cell)
        assert "Cell" in r

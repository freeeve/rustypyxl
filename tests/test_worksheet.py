"""Tests for Worksheet functionality."""

import pytest
import rustypyxl


class TestWorksheetBasics:
    """Test basic worksheet operations."""

    def test_worksheet_title(self, workbook_with_sheet):
        """Test getting worksheet title."""
        ws = workbook_with_sheet.active
        assert ws.title == "Test"

    def test_worksheet_repr(self, workbook_with_sheet):
        """Test worksheet string representation."""
        ws = workbook_with_sheet.active
        assert "Worksheet" in str(ws)
        assert "Test" in str(ws)


class TestCellAccess:
    """Test cell access methods."""

    def test_cell_by_coordinate(self, workbook_with_sheet):
        """Access cell using subscript notation."""
        ws = workbook_with_sheet.active
        cell = ws["A1"]
        assert cell is not None
        assert cell.coordinate == "A1"

    def test_cell_by_row_col(self, workbook_with_sheet):
        """Access cell using row and column numbers."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        assert cell is not None
        assert cell.row == 1
        assert cell.column == 1

    def test_cell_invalid_row_raises(self, workbook_with_sheet):
        """Accessing cell with row 0 raises ValueError."""
        ws = workbook_with_sheet.active
        with pytest.raises(ValueError):
            ws.cell(0, 1)

    def test_cell_invalid_column_raises(self, workbook_with_sheet):
        """Accessing cell with column 0 raises ValueError."""
        ws = workbook_with_sheet.active
        with pytest.raises(ValueError):
            ws.cell(1, 0)


class TestIterRows:
    """Test iterating over rows."""

    def test_iter_rows_basic(self, workbook_with_sheet):
        """Basic iteration over rows."""
        ws = workbook_with_sheet.active
        rows = ws.iter_rows(min_row=1, max_row=3, min_col=1, max_col=2)
        assert len(rows) == 3
        assert len(rows[0]) == 2

    def test_iter_rows_values_only(self, workbook_with_sheet):
        """Iteration with values_only=True."""
        ws = workbook_with_sheet.active
        rows = ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=2, values_only=True)
        assert len(rows) == 2


class TestIterCols:
    """Test iterating over columns."""

    def test_iter_cols_basic(self, workbook_with_sheet):
        """Basic iteration over columns."""
        ws = workbook_with_sheet.active
        cols = ws.iter_cols(min_col=1, max_col=3, min_row=1, max_row=2)
        assert len(cols) == 3
        assert len(cols[0]) == 2


class TestDimensions:
    """Test worksheet dimension properties."""

    def test_dimensions(self, workbook_with_sheet):
        """Test dimensions property."""
        ws = workbook_with_sheet.active
        dims = ws.dimensions
        assert dims is not None
        assert ":" in dims

    def test_max_row(self, workbook_with_sheet):
        """Test max_row property."""
        ws = workbook_with_sheet.active
        assert ws.max_row >= 1

    def test_max_column(self, workbook_with_sheet):
        """Test max_column property."""
        ws = workbook_with_sheet.active
        assert ws.max_column >= 1

    def test_min_row(self, workbook_with_sheet):
        """Test min_row property."""
        ws = workbook_with_sheet.active
        assert ws.min_row >= 1

    def test_min_column(self, workbook_with_sheet):
        """Test min_column property."""
        ws = workbook_with_sheet.active
        assert ws.min_column >= 1


class TestRowColumnOperations:
    """Test row and column insert/delete operations."""

    def test_insert_rows(self, workbook_with_sheet):
        """Test inserting rows."""
        ws = workbook_with_sheet.active
        ws.insert_rows(1, 2)  # Should not raise

    def test_insert_cols(self, workbook_with_sheet):
        """Test inserting columns."""
        ws = workbook_with_sheet.active
        ws.insert_cols(1, 2)  # Should not raise

    def test_delete_rows(self, workbook_with_sheet):
        """Test deleting rows."""
        ws = workbook_with_sheet.active
        ws.delete_rows(1, 2)  # Should not raise

    def test_delete_cols(self, workbook_with_sheet):
        """Test deleting columns."""
        ws = workbook_with_sheet.active
        ws.delete_cols(1, 2)  # Should not raise

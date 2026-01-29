"""Tests for Workbook functionality."""

import os
import pytest
import rustypyxl


class TestWorkbookCreation:
    """Test workbook creation and basic operations."""

    def test_create_empty_workbook(self):
        """Create an empty workbook."""
        wb = rustypyxl.Workbook()
        assert wb is not None
        assert len(wb) == 0

    def test_workbook_repr(self):
        """Test workbook string representation."""
        wb = rustypyxl.Workbook()
        assert "Workbook" in str(wb)
        assert "0 sheet" in str(wb)


class TestWorksheetManagement:
    """Test worksheet creation, access, and removal."""

    def test_create_sheet_with_name(self, empty_workbook):
        """Create a sheet with a specific name."""
        ws = empty_workbook.create_sheet("MySheet")
        assert ws.title == "MySheet"
        assert len(empty_workbook) == 1

    def test_create_sheet_without_name(self, empty_workbook):
        """Create a sheet without specifying a name."""
        ws = empty_workbook.create_sheet()
        assert ws.title.startswith("Sheet")
        assert len(empty_workbook) == 1

    def test_active_sheet(self, workbook_with_sheet):
        """Test accessing the active sheet."""
        active = workbook_with_sheet.active
        assert active is not None
        assert active.title == "Test"

    def test_get_sheet_by_name(self, workbook_with_sheet):
        """Get sheet by name using subscript notation."""
        ws = workbook_with_sheet["Test"]
        assert ws.title == "Test"

    def test_get_nonexistent_sheet_raises(self, workbook_with_sheet):
        """Getting a nonexistent sheet raises ValueError."""
        with pytest.raises(ValueError):
            _ = workbook_with_sheet["NonExistent"]

    def test_sheet_in_workbook(self, workbook_with_sheet):
        """Test 'in' operator for sheets."""
        assert "Test" in workbook_with_sheet
        assert "NonExistent" not in workbook_with_sheet

    def test_sheetnames(self, workbook_with_sheet):
        """Test getting all sheet names."""
        names = workbook_with_sheet.sheetnames
        assert names == ["Test"]

    def test_worksheets_property(self, workbook_with_sheet):
        """Test getting all worksheets."""
        sheets = workbook_with_sheet.worksheets
        assert len(sheets) == 1
        assert sheets[0].title == "Test"

    def test_iterate_over_workbook(self, workbook_with_sheet):
        """Test iterating over workbook returns sheet names."""
        names = list(workbook_with_sheet)
        assert names == ["Test"]


class TestWorkbookSaveLoad:
    """Test saving and loading workbooks."""

    def test_save_empty_workbook(self, empty_workbook, temp_xlsx_path):
        """Save an empty workbook."""
        # Need at least one sheet to save
        empty_workbook.create_sheet("Sheet1")
        empty_workbook.save(temp_xlsx_path)
        assert os.path.exists(temp_xlsx_path)

    def test_save_and_load_workbook(self, workbook_with_sheet, temp_xlsx_path):
        """Save and reload a workbook."""
        workbook_with_sheet.save(temp_xlsx_path)

        wb2 = rustypyxl.load_workbook(temp_xlsx_path)
        assert len(wb2) == 1
        assert "Test" in wb2.sheetnames

    def test_load_existing_file(self, sample_xlsx_path):
        """Load an existing xlsx file."""
        if sample_xlsx_path is None:
            pytest.skip("No sample xlsx file available")

        wb = rustypyxl.load_workbook(sample_xlsx_path)
        assert wb is not None
        assert len(wb) > 0

    def test_load_nonexistent_file_raises(self):
        """Loading a nonexistent file raises ValueError."""
        with pytest.raises(ValueError):
            rustypyxl.load_workbook("/nonexistent/path/file.xlsx")


class TestWorkbookClose:
    """Test workbook close method."""

    def test_close_workbook(self, empty_workbook):
        """Close method should not raise (no-op)."""
        empty_workbook.close()  # Should not raise

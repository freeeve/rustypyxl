"""Tests for named range support."""

import os
import pytest
import rustypyxl


class TestNamedRangeBasics:
    """Test basic named range operations."""

    def test_workbook_defined_names_empty(self, empty_workbook):
        """Empty workbook has no defined names."""
        assert empty_workbook.defined_names == []

    def test_create_named_range(self, workbook_with_sheet):
        """Create a named range."""
        ws = workbook_with_sheet.active
        workbook_with_sheet.create_named_range("MyRange", ws, "A1:B10")

        names = workbook_with_sheet.defined_names
        assert len(names) == 1
        assert names[0][0] == "MyRange"

    def test_create_multiple_named_ranges(self, workbook_with_sheet):
        """Create multiple named ranges."""
        ws = workbook_with_sheet.active
        workbook_with_sheet.create_named_range("Range1", ws, "A1:A10")
        workbook_with_sheet.create_named_range("Range2", ws, "B1:B10")

        names = workbook_with_sheet.defined_names
        assert len(names) == 2


class TestNamedRangeRoundtrip:
    """Test named ranges survive save/load cycle."""

    def test_named_range_roundtrip(self, temp_xlsx_path):
        """Save and load named ranges."""
        wb1 = rustypyxl.Workbook()
        ws = wb1.create_sheet("Data")
        wb1.create_named_range("TestRange", ws, "A1:C10")
        wb1.save(temp_xlsx_path)

        wb2 = rustypyxl.load_workbook(temp_xlsx_path)
        names = wb2.defined_names
        assert any("TestRange" in n[0] for n in names)

    def test_load_existing_named_ranges(self):
        """Load an existing file with named ranges."""
        path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "test_named_ranges.xlsx")
        if not os.path.exists(path):
            pytest.skip("test_named_ranges.xlsx not found")

        wb = rustypyxl.load_workbook(path)
        assert wb is not None
        # Check that named ranges were loaded
        names = wb.defined_names
        assert isinstance(names, list)

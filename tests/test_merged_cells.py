"""Tests for merged cell support."""

import pytest
import rustypyxl


class TestMergedCellBasics:
    """Test basic merged cell operations."""

    def test_merged_cells_empty_default(self, workbook_with_sheet):
        """No merged cells by default."""
        ws = workbook_with_sheet.active
        assert ws.merged_cells == []

    def test_merge_cells_by_range_string(self, workbook_with_sheet):
        """Merge cells using range string."""
        ws = workbook_with_sheet.active
        ws.merge_cells("A1:B2")
        # Note: actual merging requires workbook access

    def test_merge_cells_by_coordinates(self, workbook_with_sheet):
        """Merge cells using row/column coordinates."""
        ws = workbook_with_sheet.active
        ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=2)

    def test_unmerge_cells(self, workbook_with_sheet):
        """Unmerge previously merged cells."""
        ws = workbook_with_sheet.active
        ws.merge_cells("A1:B2")
        ws.unmerge_cells("A1:B2")

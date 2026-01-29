"""Tests for hyperlink support."""

import os
import pytest
import rustypyxl


class TestHyperlinkBasics:
    """Test basic hyperlink operations."""

    def test_cell_hyperlink_property_default(self, workbook_with_sheet):
        """Cell hyperlink is None by default."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        assert cell.hyperlink is None

    def test_set_cell_hyperlink(self, workbook_with_sheet):
        """Set a hyperlink on a cell."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        cell.hyperlink = "https://example.com"
        assert cell.hyperlink == "https://example.com"

    def test_set_mailto_hyperlink(self, workbook_with_sheet):
        """Set a mailto hyperlink."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        cell.hyperlink = "mailto:test@example.com"
        assert cell.hyperlink == "mailto:test@example.com"

    def test_set_internal_hyperlink(self, workbook_with_sheet):
        """Set an internal reference hyperlink."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        cell.hyperlink = "#Sheet2!A1"
        assert cell.hyperlink == "#Sheet2!A1"

    def test_clear_hyperlink(self, workbook_with_sheet):
        """Clear a hyperlink by setting to None."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        cell.hyperlink = "https://example.com"
        cell.hyperlink = None
        assert cell.hyperlink is None


class TestHyperlinkRoundtrip:
    """Test hyperlinks survive save/load cycle."""

    def test_load_existing_hyperlinks(self):
        """Load an existing file with hyperlinks."""
        path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "test_hyperlinks.xlsx")
        if not os.path.exists(path):
            pytest.skip("test_hyperlinks.xlsx not found")

        wb = rustypyxl.load_workbook(path)
        assert wb is not None
        assert len(wb) > 0

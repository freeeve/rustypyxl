"""Tests for comment support."""

import rustypyxl


class TestCommentBasics:
    """Test basic comment operations."""

    def test_cell_comment_property_default(self, workbook_with_sheet):
        """Cell comment is None by default."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        assert cell.comment is None

    def test_set_cell_comment(self, workbook_with_sheet):
        """Set a comment on a cell."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        cell.comment = "This is a comment"
        assert cell.comment == "This is a comment"

    def test_set_multiline_comment(self, workbook_with_sheet):
        """Set a multiline comment."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        cell.comment = "Line 1\nLine 2\nLine 3"
        assert "Line 1" in cell.comment
        assert "Line 2" in cell.comment

    def test_clear_comment(self, workbook_with_sheet):
        """Clear a comment by setting to None."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        cell.comment = "A comment"
        cell.comment = None
        assert cell.comment is None


class TestCommentRoundtrip:
    """Test comments survive save/load cycle."""

    def test_load_existing_comments(self, fixtures_dir):
        """Load an externally-authored file with comments."""
        wb = rustypyxl.load_workbook(str(fixtures_dir / "comments.xlsx"))
        ws = wb["Commented"]
        assert ws["A1"].comment is not None, "comment lost on load"
        assert "first comment" in ws["A1"].comment
        assert "second comment" in ws.cell(2, 2).comment

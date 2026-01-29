"""Tests for styling classes."""

import pytest
import rustypyxl


class TestFont:
    """Test Font styling class."""

    def test_font_creation_default(self):
        """Create a font with default values."""
        font = rustypyxl.Font()
        assert font.bold is False
        assert font.italic is False

    def test_font_creation_with_args(self):
        """Create a font with specific values."""
        font = rustypyxl.Font(name="Arial", size=12, bold=True, italic=True)
        assert font.name == "Arial"
        assert font.size == 12
        assert font.bold is True
        assert font.italic is True

    def test_font_underline(self):
        """Test font underline."""
        font = rustypyxl.Font(underline="single")
        assert font.underline == "single"

    def test_font_color(self):
        """Test font color."""
        font = rustypyxl.Font(color="#FF0000")
        assert font.color == "#FF0000"

    def test_font_copy(self):
        """Test font copy method."""
        font1 = rustypyxl.Font(name="Arial", bold=True)
        font2 = font1.copy()
        assert font2.name == "Arial"
        assert font2.bold is True

    def test_font_repr(self):
        """Test font string representation."""
        font = rustypyxl.Font(name="Arial")
        s = str(font)
        assert "Font" in s


class TestAlignment:
    """Test Alignment styling class."""

    def test_alignment_creation_default(self):
        """Create alignment with default values."""
        align = rustypyxl.Alignment()
        assert align.wrap_text is False

    def test_alignment_creation_with_args(self):
        """Create alignment with specific values."""
        align = rustypyxl.Alignment(horizontal="center", vertical="top", wrap_text=True)
        assert align.horizontal == "center"
        assert align.vertical == "top"
        assert align.wrap_text is True

    def test_alignment_indent(self):
        """Test alignment indent."""
        align = rustypyxl.Alignment(indent=2)
        assert align.indent == 2

    def test_alignment_text_rotation(self):
        """Test alignment text rotation."""
        align = rustypyxl.Alignment(text_rotation=45)
        assert align.text_rotation == 45

    def test_alignment_copy(self):
        """Test alignment copy method."""
        align1 = rustypyxl.Alignment(horizontal="center")
        align2 = align1.copy()
        assert align2.horizontal == "center"

    def test_alignment_repr(self):
        """Test alignment string representation."""
        align = rustypyxl.Alignment(horizontal="center")
        s = str(align)
        assert "Alignment" in s


class TestPatternFill:
    """Test PatternFill styling class."""

    def test_pattern_fill_default(self):
        """Create pattern fill with default values."""
        fill = rustypyxl.PatternFill()
        assert fill.fill_type is None

    def test_pattern_fill_solid(self):
        """Create a solid fill."""
        fill = rustypyxl.PatternFill(fill_type="solid", fgColor="#FFFF00")
        assert fill.fill_type == "solid"
        assert fill.fgColor == "#FFFF00"

    def test_pattern_fill_copy(self):
        """Test pattern fill copy method."""
        fill1 = rustypyxl.PatternFill(fill_type="solid")
        fill2 = fill1.copy()
        assert fill2.fill_type == "solid"

    def test_pattern_fill_repr(self):
        """Test pattern fill string representation."""
        fill = rustypyxl.PatternFill(fill_type="solid")
        s = str(fill)
        assert "PatternFill" in s


class TestBorder:
    """Test Border styling class."""

    def test_border_creation_default(self):
        """Create a border with default values."""
        border = rustypyxl.Border()
        assert border.left is None
        assert border.right is None
        assert border.top is None
        assert border.bottom is None

    def test_border_copy(self):
        """Test border copy method."""
        border1 = rustypyxl.Border(outline=False)
        border2 = border1.copy()
        assert border2.outline is False

    def test_border_repr(self):
        """Test border string representation."""
        border = rustypyxl.Border()
        s = str(border)
        assert "Border" in s


class TestCellStyling:
    """Test applying styles to cells."""

    def test_cell_font(self, workbook_with_sheet):
        """Apply font to a cell."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)

        font = rustypyxl.Font(bold=True)
        cell.font = font

        assert cell.font is not None
        assert cell.font.bold is True

    def test_cell_alignment(self, workbook_with_sheet):
        """Apply alignment to a cell."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)

        align = rustypyxl.Alignment(horizontal="center")
        cell.alignment = align

        assert cell.alignment is not None
        assert cell.alignment.horizontal == "center"

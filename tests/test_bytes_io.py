"""Tests for bytes I/O functionality."""

import io
import pytest
import rustypyxl


class TestSaveToBytes:
    """Tests for save_to_bytes() method."""

    def test_save_empty_workbook_to_bytes(self):
        """Empty workbook should produce valid xlsx bytes."""
        wb = rustypyxl.Workbook()
        wb.create_sheet("Sheet1")

        data = wb.save_to_bytes()

        assert isinstance(data, bytes)
        assert len(data) > 0
        # XLSX files are ZIP archives starting with PK
        assert data[:2] == b'PK'

    def test_save_workbook_with_data_to_bytes(self):
        """Workbook with data should produce valid xlsx bytes."""
        wb = rustypyxl.Workbook()
        wb.create_sheet("Data")
        wb.set_cell_value("Data", 1, 1, "Hello")
        wb.set_cell_value("Data", 1, 2, 42)
        wb.set_cell_value("Data", 2, 1, True)

        data = wb.save_to_bytes()

        assert isinstance(data, bytes)
        assert len(data) > 100  # Should have some content

    def test_save_multiple_sheets_to_bytes(self):
        """Multiple sheets should be preserved in bytes output."""
        wb = rustypyxl.Workbook()
        wb.create_sheet("Sheet1")
        wb.create_sheet("Sheet2")
        wb.set_cell_value("Sheet1", 1, 1, "First")
        wb.set_cell_value("Sheet2", 1, 1, "Second")

        data = wb.save_to_bytes()

        # Reload and verify
        wb2 = rustypyxl.load_workbook(data)
        assert len(wb2.sheetnames) == 2
        assert "Sheet1" in wb2.sheetnames
        assert "Sheet2" in wb2.sheetnames


class TestLoadFromBytes:
    """Tests for loading workbooks from bytes."""

    def test_load_from_bytes(self):
        """Should load workbook from bytes."""
        # Create and save
        wb = rustypyxl.Workbook()
        wb.create_sheet("Test")
        wb.set_cell_value("Test", 1, 1, "Hello World")
        data = wb.save_to_bytes()

        # Load from bytes
        wb2 = rustypyxl.load_workbook(data)

        assert wb2.sheetnames == ["Test"]
        assert wb2.get_cell_value("Test", 1, 1) == "Hello World"

    def test_load_from_bytes_via_workbook_class(self):
        """Should load via Workbook.load() with bytes."""
        wb = rustypyxl.Workbook()
        wb.create_sheet("Data")
        wb.set_cell_value("Data", 1, 1, 123.45)
        data = wb.save_to_bytes()

        wb2 = rustypyxl.Workbook.load(data)

        assert wb2.get_cell_value("Data", 1, 1) == 123.45

    def test_load_from_bytesio(self):
        """Should load from BytesIO file-like object."""
        wb = rustypyxl.Workbook()
        wb.create_sheet("Sheet1")
        wb.set_cell_value("Sheet1", 1, 1, "From BytesIO")
        data = wb.save_to_bytes()

        file_obj = io.BytesIO(data)
        wb2 = rustypyxl.load_workbook(file_obj)

        assert wb2.get_cell_value("Sheet1", 1, 1) == "From BytesIO"

    def test_load_invalid_bytes_raises(self):
        """Should raise error for invalid bytes."""
        with pytest.raises(ValueError):
            rustypyxl.load_workbook(b"not a valid xlsx file")

    def test_load_empty_bytes_raises(self):
        """Should raise error for empty bytes."""
        with pytest.raises(ValueError):
            rustypyxl.load_workbook(b"")


class TestBytesRoundtrip:
    """Tests for roundtrip through bytes."""

    def test_roundtrip_preserves_strings(self):
        """String values should survive roundtrip."""
        wb = rustypyxl.Workbook()
        wb.create_sheet("Test")
        wb.set_cell_value("Test", 1, 1, "Hello")
        wb.set_cell_value("Test", 1, 2, "World")

        data = wb.save_to_bytes()
        wb2 = rustypyxl.load_workbook(data)

        assert wb2.get_cell_value("Test", 1, 1) == "Hello"
        assert wb2.get_cell_value("Test", 1, 2) == "World"

    def test_roundtrip_preserves_numbers(self):
        """Numeric values should survive roundtrip."""
        wb = rustypyxl.Workbook()
        wb.create_sheet("Test")
        wb.set_cell_value("Test", 1, 1, 42)
        wb.set_cell_value("Test", 1, 2, 3.14159)
        wb.set_cell_value("Test", 1, 3, -100.5)

        data = wb.save_to_bytes()
        wb2 = rustypyxl.load_workbook(data)

        assert wb2.get_cell_value("Test", 1, 1) == 42
        assert abs(wb2.get_cell_value("Test", 1, 2) - 3.14159) < 0.0001
        assert wb2.get_cell_value("Test", 1, 3) == -100.5

    def test_roundtrip_preserves_booleans(self):
        """Boolean values should survive roundtrip."""
        wb = rustypyxl.Workbook()
        wb.create_sheet("Test")
        wb.set_cell_value("Test", 1, 1, True)
        wb.set_cell_value("Test", 1, 2, False)

        data = wb.save_to_bytes()
        wb2 = rustypyxl.load_workbook(data)

        assert wb2.get_cell_value("Test", 1, 1) is True
        assert wb2.get_cell_value("Test", 1, 2) is False

    def test_roundtrip_preserves_formulas(self):
        """Formulas should survive roundtrip."""
        wb = rustypyxl.Workbook()
        wb.create_sheet("Test")
        wb.set_cell_value("Test", 1, 1, 10)
        wb.set_cell_value("Test", 1, 2, 20)
        wb.set_cell_value("Test", 1, 3, "=A1+B1")

        data = wb.save_to_bytes()
        wb2 = rustypyxl.load_workbook(data)

        # Formula is stored without the '=' prefix internally
        # Use get_cell_value which properly returns formula values
        formula = wb2.get_cell_value("Test", 1, 3)
        assert formula is not None
        assert "A1" in formula and "B1" in formula

    def test_roundtrip_preserves_sheet_names(self):
        """Sheet names should survive roundtrip."""
        wb = rustypyxl.Workbook()
        wb.create_sheet("First Sheet")
        wb.create_sheet("Second Sheet")
        wb.create_sheet("Sheet With Spaces")

        data = wb.save_to_bytes()
        wb2 = rustypyxl.load_workbook(data)

        assert wb2.sheetnames == ["First Sheet", "Second Sheet", "Sheet With Spaces"]

    def test_roundtrip_large_data(self):
        """Large datasets should survive roundtrip."""
        wb = rustypyxl.Workbook()
        wb.create_sheet("Data")

        # Write 1000 rows
        rows = [[f"Row {i}", i, i * 1.5] for i in range(1000)]
        wb.write_rows("Data", rows)

        data = wb.save_to_bytes()
        wb2 = rustypyxl.load_workbook(data)

        # Verify a few rows
        assert wb2.get_cell_value("Data", 1, 1) == "Row 0"
        assert wb2.get_cell_value("Data", 500, 2) == 499
        assert wb2.get_cell_value("Data", 1000, 1) == "Row 999"


class TestLoadWorkbookTypeDetection:
    """Tests for load_workbook type detection."""

    def test_load_from_string_path(self, tmp_path):
        """Should load from string file path."""
        path = tmp_path / "test.xlsx"

        wb = rustypyxl.Workbook()
        wb.create_sheet("Test")
        wb.save(str(path))

        wb2 = rustypyxl.load_workbook(str(path))
        assert "Test" in wb2.sheetnames

    def test_load_rejects_invalid_type(self):
        """Should reject invalid types."""
        with pytest.raises(TypeError):
            rustypyxl.load_workbook(12345)

        with pytest.raises(TypeError):
            rustypyxl.load_workbook(['not', 'valid'])

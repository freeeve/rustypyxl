"""Tests for WriteOnlyWorkbook (streaming writes)."""

import os
import pytest
import rustypyxl


class TestWriteOnlyWorkbookBasic:
    """Basic tests for WriteOnlyWorkbook."""

    def test_create_and_close(self, tmp_path):
        """Should create a valid file with create and close."""
        path = tmp_path / "test.xlsx"

        wb = rustypyxl.WriteOnlyWorkbook(str(path))
        wb.create_sheet("Sheet1")
        wb.close()

        assert path.exists()
        assert path.stat().st_size > 0

    def test_append_single_row(self, tmp_path):
        """Should write a single row."""
        path = tmp_path / "test.xlsx"

        wb = rustypyxl.WriteOnlyWorkbook(str(path))
        wb.create_sheet("Data")
        wb.append_row(["Hello", "World", 123])
        wb.close()

        # Verify by loading
        wb2 = rustypyxl.load_workbook(str(path))
        assert wb2.get_cell_value("Data", 1, 1) == "Hello"
        assert wb2.get_cell_value("Data", 1, 2) == "World"
        assert wb2.get_cell_value("Data", 1, 3) == 123

    def test_append_multiple_rows(self, tmp_path):
        """Should write multiple rows."""
        path = tmp_path / "test.xlsx"

        wb = rustypyxl.WriteOnlyWorkbook(str(path))
        wb.create_sheet("Data")
        wb.append_row(["Name", "Age", "Score"])
        wb.append_row(["Alice", 30, 95.5])
        wb.append_row(["Bob", 25, 87.3])
        wb.close()

        wb2 = rustypyxl.load_workbook(str(path))
        assert wb2.get_cell_value("Data", 1, 1) == "Name"
        assert wb2.get_cell_value("Data", 2, 1) == "Alice"
        assert wb2.get_cell_value("Data", 3, 1) == "Bob"
        assert wb2.get_cell_value("Data", 2, 2) == 30
        assert wb2.get_cell_value("Data", 3, 3) == 87.3


class TestWriteOnlyWorkbookDataTypes:
    """Tests for different data types in WriteOnlyWorkbook."""

    def test_string_values(self, tmp_path):
        """Should handle string values."""
        path = tmp_path / "test.xlsx"

        wb = rustypyxl.WriteOnlyWorkbook(str(path))
        wb.create_sheet("Test")
        wb.append_row(["Simple", "With spaces", "Special: <>&"])
        wb.close()

        wb2 = rustypyxl.load_workbook(str(path))
        assert wb2.get_cell_value("Test", 1, 1) == "Simple"
        assert wb2.get_cell_value("Test", 1, 2) == "With spaces"
        assert wb2.get_cell_value("Test", 1, 3) == "Special: <>&"

    def test_numeric_values(self, tmp_path):
        """Should handle numeric values."""
        path = tmp_path / "test.xlsx"

        wb = rustypyxl.WriteOnlyWorkbook(str(path))
        wb.create_sheet("Test")
        wb.append_row([42, 3.14159, -100, 0])
        wb.close()

        wb2 = rustypyxl.load_workbook(str(path))
        assert wb2.get_cell_value("Test", 1, 1) == 42
        assert abs(wb2.get_cell_value("Test", 1, 2) - 3.14159) < 0.0001
        assert wb2.get_cell_value("Test", 1, 3) == -100
        assert wb2.get_cell_value("Test", 1, 4) == 0

    def test_boolean_values(self, tmp_path):
        """Should handle boolean values."""
        path = tmp_path / "test.xlsx"

        wb = rustypyxl.WriteOnlyWorkbook(str(path))
        wb.create_sheet("Test")
        wb.append_row([True, False, True])
        wb.close()

        wb2 = rustypyxl.load_workbook(str(path))
        assert wb2.get_cell_value("Test", 1, 1) is True
        assert wb2.get_cell_value("Test", 1, 2) is False
        assert wb2.get_cell_value("Test", 1, 3) is True

    def test_none_values(self, tmp_path):
        """Should handle None values as empty cells."""
        path = tmp_path / "test.xlsx"

        wb = rustypyxl.WriteOnlyWorkbook(str(path))
        wb.create_sheet("Test")
        wb.append_row(["Before", None, "After"])
        wb.close()

        wb2 = rustypyxl.load_workbook(str(path))
        assert wb2.get_cell_value("Test", 1, 1) == "Before"
        assert wb2.get_cell_value("Test", 1, 2) is None
        assert wb2.get_cell_value("Test", 1, 3) == "After"

    def test_mixed_values(self, tmp_path):
        """Should handle mixed value types in one row."""
        path = tmp_path / "test.xlsx"

        wb = rustypyxl.WriteOnlyWorkbook(str(path))
        wb.create_sheet("Test")
        wb.append_row(["String", 123, 45.67, True, None])
        wb.close()

        wb2 = rustypyxl.load_workbook(str(path))
        assert wb2.get_cell_value("Test", 1, 1) == "String"
        assert wb2.get_cell_value("Test", 1, 2) == 123
        assert abs(wb2.get_cell_value("Test", 1, 3) - 45.67) < 0.01
        assert wb2.get_cell_value("Test", 1, 4) is True
        assert wb2.get_cell_value("Test", 1, 5) is None


class TestWriteOnlyWorkbookLargeFiles:
    """Tests for large file handling."""

    def test_many_rows(self, tmp_path):
        """Should handle many rows efficiently."""
        path = tmp_path / "test.xlsx"

        wb = rustypyxl.WriteOnlyWorkbook(str(path))
        wb.create_sheet("Data")

        # Write 10,000 rows
        for i in range(10000):
            wb.append_row([f"Row {i}", i, i * 1.5])

        wb.close()

        # Verify file exists and has content
        assert path.exists()
        assert path.stat().st_size > 100000  # Should be reasonably large

        # Spot check some rows
        wb2 = rustypyxl.load_workbook(str(path))
        assert wb2.get_cell_value("Data", 1, 1) == "Row 0"
        assert wb2.get_cell_value("Data", 5000, 1) == "Row 4999"
        assert wb2.get_cell_value("Data", 10000, 1) == "Row 9999"

    def test_wide_rows(self, tmp_path):
        """Should handle wide rows (many columns)."""
        path = tmp_path / "test.xlsx"

        wb = rustypyxl.WriteOnlyWorkbook(str(path))
        wb.create_sheet("Wide")

        # Write row with 100 columns
        row = [f"Col{i}" for i in range(100)]
        wb.append_row(row)
        wb.close()

        wb2 = rustypyxl.load_workbook(str(path))
        assert wb2.get_cell_value("Wide", 1, 1) == "Col0"
        assert wb2.get_cell_value("Wide", 1, 50) == "Col49"
        assert wb2.get_cell_value("Wide", 1, 100) == "Col99"


class TestWriteOnlyWorkbookErrors:
    """Tests for error handling."""

    def test_append_before_create_sheet_raises(self, tmp_path):
        """Should raise error if appending before creating sheet."""
        path = tmp_path / "test.xlsx"

        wb = rustypyxl.WriteOnlyWorkbook(str(path))

        with pytest.raises(ValueError):
            wb.append_row(["This", "should", "fail"])

    def test_close_before_create_sheet_raises(self, tmp_path):
        """Should raise error if closing before creating sheet."""
        path = tmp_path / "test.xlsx"

        wb = rustypyxl.WriteOnlyWorkbook(str(path))

        with pytest.raises(ValueError):
            wb.close()

    def test_double_close_raises(self, tmp_path):
        """Should raise error on double close."""
        path = tmp_path / "test.xlsx"

        wb = rustypyxl.WriteOnlyWorkbook(str(path))
        wb.create_sheet("Test")
        wb.close()

        with pytest.raises(ValueError):
            wb.close()

    def test_append_after_close_raises(self, tmp_path):
        """Should raise error if appending after close."""
        path = tmp_path / "test.xlsx"

        wb = rustypyxl.WriteOnlyWorkbook(str(path))
        wb.create_sheet("Test")
        wb.close()

        with pytest.raises(ValueError):
            wb.append_row(["Too", "late"])

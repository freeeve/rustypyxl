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

    def test_close_before_create_sheet_writes_default_sheet(self, tmp_path):
        """Closing with no sheets produces a valid workbook with a default
        sheet, since xlsx requires at least one."""
        path = tmp_path / "test.xlsx"

        wb = rustypyxl.WriteOnlyWorkbook(str(path))
        wb.close()

        chk = rustypyxl.load_workbook(str(path))
        assert chk.sheetnames == ["Sheet1"]

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


class TestStreamingMultiSheet:
    """Multi-sheet streaming and context-manager support (task 008)."""

    def test_multiple_sheets(self, temp_xlsx_path):
        wb = rustypyxl.WriteOnlyWorkbook(temp_xlsx_path)
        wb.create_sheet("First")
        wb.append_row(["a", 1])
        wb.create_sheet("Second")  # finalizes "First" automatically
        wb.append_row([42])
        wb.close()

        chk = rustypyxl.load_workbook(temp_xlsx_path)
        assert chk.sheetnames == ["First", "Second"]
        assert chk["First"]["A1"].value == "a"
        assert chk["Second"]["A1"].value == 42

    def test_context_manager_closes_file(self, temp_xlsx_path):
        with rustypyxl.WriteOnlyWorkbook(temp_xlsx_path) as wb:
            wb.create_sheet("S")
            wb.append_row(["from with-block"])

        chk = rustypyxl.load_workbook(temp_xlsx_path)
        assert chk["S"]["A1"].value == "from with-block"

    def test_context_manager_does_not_mask_exceptions(self, temp_xlsx_path):
        with pytest.raises(RuntimeError, match="boom"):
            with rustypyxl.WriteOnlyWorkbook(temp_xlsx_path) as wb:
                wb.create_sheet("S")
                raise RuntimeError("boom")

    def test_too_many_columns_rejected(self, temp_xlsx_path):
        wb = rustypyxl.WriteOnlyWorkbook(temp_xlsx_path)
        wb.create_sheet("S")
        with pytest.raises(ValueError, match="column limit"):
            wb.append_row([1] * 16_385)
        wb.close()

    def test_invalid_sheet_names_rejected(self, temp_xlsx_path):
        wb = rustypyxl.WriteOnlyWorkbook(temp_xlsx_path)
        with pytest.raises(ValueError):
            wb.create_sheet("bad/name")
        with pytest.raises(ValueError):
            wb.create_sheet("x" * 32)
        wb.create_sheet("Fine")
        with pytest.raises(ValueError):
            wb.create_sheet("Fine")
        wb.close()


class TestWriteOnlyWorkbookBatch:
    """append_rows writes a batch with the GIL released once, not per row."""

    def test_append_rows_writes_every_row(self, temp_xlsx_path):
        wb = rustypyxl.WriteOnlyWorkbook(temp_xlsx_path)
        wb.create_sheet("S")
        wb.append_rows([["a", 1, True], ["b", 2, False], ["c", 3, None]])
        wb.close()

        chk = rustypyxl.load_workbook(temp_xlsx_path)["S"]
        assert chk["A1"].value == "a"
        assert chk["B2"].value == 2
        assert chk["C1"].value is True
        assert chk["A3"].value == "c"

    def test_append_rows_interleaves_with_append_row(self, temp_xlsx_path):
        wb = rustypyxl.WriteOnlyWorkbook(temp_xlsx_path)
        wb.create_sheet("S")
        wb.append_row(["first"])
        wb.append_rows([["second"], ["third"]])
        wb.append_row(["fourth"])
        wb.close()

        chk = rustypyxl.load_workbook(temp_xlsx_path)["S"]
        assert [chk[f"A{r}"].value for r in range(1, 5)] == [
            "first",
            "second",
            "third",
            "fourth",
        ]

    def test_append_rows_empty_batch_is_a_noop(self, temp_xlsx_path):
        wb = rustypyxl.WriteOnlyWorkbook(temp_xlsx_path)
        wb.create_sheet("S")
        wb.append_rows([])
        wb.append_row(["only"])
        wb.close()

        assert rustypyxl.load_workbook(temp_xlsx_path)["S"]["A1"].value == "only"

    def test_append_rows_requires_a_sheet(self, temp_xlsx_path):
        wb = rustypyxl.WriteOnlyWorkbook(temp_xlsx_path)
        with pytest.raises(ValueError, match="No sheet"):
            wb.append_rows([["x"]])
        wb.create_sheet("S")
        wb.close()

    def test_append_rows_rejects_too_many_columns(self, temp_xlsx_path):
        wb = rustypyxl.WriteOnlyWorkbook(temp_xlsx_path)
        wb.create_sheet("S")
        with pytest.raises(ValueError, match="column limit"):
            wb.append_rows([[1] * 16_385])
        wb.close()

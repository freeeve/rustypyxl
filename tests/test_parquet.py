"""Tests for Parquet import/export functionality."""

import pytest

# Check if parquet is available
try:
    import pyarrow.parquet as pq
    import pyarrow as pa
    HAS_PYARROW = True
except ImportError:
    HAS_PYARROW = False

import rustypyxl


pytestmark = pytest.mark.skipif(not HAS_PYARROW, reason="pyarrow not installed")


class TestParquetImport:
    """Tests for importing Parquet files into worksheets."""

    def test_import_basic_parquet(self, tmp_path):
        """Should import a basic parquet file."""
        # Create a parquet file
        parquet_path = tmp_path / "data.parquet"
        table = pa.table({
            "name": ["Alice", "Bob", "Charlie"],
            "age": [30, 25, 35],
            "score": [95.5, 87.3, 92.1],
        })
        pq.write_table(table, parquet_path)

        # Import into workbook
        wb = rustypyxl.Workbook()
        wb.create_sheet("Data")
        result = wb.insert_from_parquet(
            sheet_name="Data",
            path=str(parquet_path),
        )

        assert result["rows_imported"] == 3
        assert result["columns_imported"] == 3
        assert "name" in result["column_names"]

        # Verify data
        assert wb.get_cell_value("Data", 1, 1) == "name"  # Header
        assert wb.get_cell_value("Data", 2, 1) == "Alice"
        assert wb.get_cell_value("Data", 3, 2) == 25

    def test_import_without_headers(self, tmp_path):
        """Should import without headers."""
        parquet_path = tmp_path / "data.parquet"
        table = pa.table({
            "col1": [1, 2, 3],
            "col2": [4, 5, 6],
        })
        pq.write_table(table, parquet_path)

        wb = rustypyxl.Workbook()
        wb.create_sheet("Data")
        result = wb.insert_from_parquet(
            sheet_name="Data",
            path=str(parquet_path),
            include_headers=False,
        )

        assert result["rows_imported"] == 3
        # First row should be data, not headers
        assert wb.get_cell_value("Data", 1, 1) == 1

    def test_import_with_start_position(self, tmp_path):
        """Should import at specified start position."""
        parquet_path = tmp_path / "data.parquet"
        table = pa.table({"value": [10, 20, 30]})
        pq.write_table(table, parquet_path)

        wb = rustypyxl.Workbook()
        wb.create_sheet("Data")
        wb.set_cell_value("Data", 1, 1, "Title")

        result = wb.insert_from_parquet(
            sheet_name="Data",
            path=str(parquet_path),
            start_row=3,
            start_col=2,
        )

        # Title should still be there
        assert wb.get_cell_value("Data", 1, 1) == "Title"
        # Data should start at row 3, col 2
        assert wb.get_cell_value("Data", 3, 2) == "value"  # Header
        assert wb.get_cell_value("Data", 4, 2) == 10

    def test_import_with_column_selection(self, tmp_path):
        """Should import only selected columns."""
        parquet_path = tmp_path / "data.parquet"
        table = pa.table({
            "a": [1, 2],
            "b": [3, 4],
            "c": [5, 6],
        })
        pq.write_table(table, parquet_path)

        wb = rustypyxl.Workbook()
        wb.create_sheet("Data")
        result = wb.insert_from_parquet(
            sheet_name="Data",
            path=str(parquet_path),
            columns=["a", "c"],
        )

        assert result["columns_imported"] == 2
        assert wb.get_cell_value("Data", 1, 1) == "a"
        assert wb.get_cell_value("Data", 1, 2) == "c"

    def test_import_with_column_renames(self, tmp_path):
        """Should rename columns during import."""
        parquet_path = tmp_path / "data.parquet"
        table = pa.table({"old_name": [1, 2, 3]})
        pq.write_table(table, parquet_path)

        wb = rustypyxl.Workbook()
        wb.create_sheet("Data")
        result = wb.insert_from_parquet(
            sheet_name="Data",
            path=str(parquet_path),
            column_renames={"old_name": "new_name"},
        )

        assert wb.get_cell_value("Data", 1, 1) == "new_name"

    def test_import_various_types(self, tmp_path):
        """Should handle various data types."""
        parquet_path = tmp_path / "data.parquet"
        table = pa.table({
            "int_col": pa.array([1, 2, 3], type=pa.int64()),
            "float_col": pa.array([1.1, 2.2, 3.3], type=pa.float64()),
            "str_col": ["a", "b", "c"],
            "bool_col": [True, False, True],
        })
        pq.write_table(table, parquet_path)

        wb = rustypyxl.Workbook()
        wb.create_sheet("Data")
        wb.insert_from_parquet(sheet_name="Data", path=str(parquet_path))

        # Check types
        assert wb.get_cell_value("Data", 2, 1) == 1
        assert abs(wb.get_cell_value("Data", 2, 2) - 1.1) < 0.01
        assert wb.get_cell_value("Data", 2, 3) == "a"
        assert wb.get_cell_value("Data", 2, 4) is True


class TestParquetExport:
    """Tests for exporting worksheets to Parquet files."""

    def test_export_basic_worksheet(self, tmp_path):
        """Should export worksheet to parquet."""
        wb = rustypyxl.Workbook()
        wb.create_sheet("Data")
        wb.write_rows("Data", [
            ["name", "age"],
            ["Alice", 30],
            ["Bob", 25],
        ])

        parquet_path = tmp_path / "output.parquet"
        result = wb.export_to_parquet(
            sheet_name="Data",
            path=str(parquet_path),
        )

        assert result["rows_exported"] == 2  # Excluding header
        assert result["columns_exported"] == 2
        assert parquet_path.exists()

        # Verify by reading back
        table = pq.read_table(parquet_path)
        assert table.num_rows == 2
        assert "name" in table.column_names
        assert "age" in table.column_names

    def test_export_without_headers(self, tmp_path):
        """Should export without treating first row as headers."""
        wb = rustypyxl.Workbook()
        wb.create_sheet("Data")
        wb.write_rows("Data", [
            [1, 2],
            [3, 4],
        ])

        parquet_path = tmp_path / "output.parquet"
        result = wb.export_to_parquet(
            sheet_name="Data",
            path=str(parquet_path),
            has_headers=False,
        )

        assert result["rows_exported"] == 2
        table = pq.read_table(parquet_path)
        assert table.num_rows == 2

    def test_export_with_compression(self, tmp_path):
        """Should export with different compression types."""
        wb = rustypyxl.Workbook()
        wb.create_sheet("Data")
        wb.write_rows("Data", [
            ["col1"],
            ["value"],
        ])

        for compression in ["snappy", "gzip", "zstd", "none"]:
            parquet_path = tmp_path / f"output_{compression}.parquet"
            result = wb.export_to_parquet(
                sheet_name="Data",
                path=str(parquet_path),
                compression=compression,
            )
            assert parquet_path.exists()
            assert result["file_size"] > 0

    def test_export_range(self, tmp_path):
        """Should export specific range."""
        wb = rustypyxl.Workbook()
        wb.create_sheet("Data")
        wb.write_rows("Data", [
            ["A", "B", "C"],
            [1, 2, 3],
            [4, 5, 6],
            [7, 8, 9],
        ])

        parquet_path = tmp_path / "output.parquet"
        result = wb.export_range_to_parquet(
            sheet_name="Data",
            path=str(parquet_path),
            min_row=1,
            min_col=1,
            max_row=3,
            max_col=2,
        )

        # Should only include rows 1-3, cols 1-2
        table = pq.read_table(parquet_path)
        assert table.num_columns == 2
        assert table.num_rows == 2  # Excluding header

    def test_export_with_column_renames(self, tmp_path):
        """Should rename columns during export."""
        wb = rustypyxl.Workbook()
        wb.create_sheet("Data")
        wb.write_rows("Data", [
            ["old_col"],
            ["value"],
        ])

        parquet_path = tmp_path / "output.parquet"
        wb.export_to_parquet(
            sheet_name="Data",
            path=str(parquet_path),
            column_renames={"old_col": "new_col"},
        )

        table = pq.read_table(parquet_path)
        assert "new_col" in table.column_names


class TestParquetRoundtrip:
    """Tests for roundtrip: Excel -> Parquet -> Excel."""

    def test_roundtrip_preserves_data(self, tmp_path):
        """Data should survive Excel -> Parquet -> Excel roundtrip."""
        # Create Excel
        wb = rustypyxl.Workbook()
        wb.create_sheet("Data")
        wb.write_rows("Data", [
            ["name", "value"],
            ["Test1", 100],
            ["Test2", 200],
        ])

        # Export to Parquet
        parquet_path = tmp_path / "data.parquet"
        wb.export_to_parquet(sheet_name="Data", path=str(parquet_path))

        # Import back
        wb2 = rustypyxl.Workbook()
        wb2.create_sheet("Imported")
        wb2.insert_from_parquet(sheet_name="Imported", path=str(parquet_path))

        # Verify
        assert wb2.get_cell_value("Imported", 1, 1) == "name"
        assert wb2.get_cell_value("Imported", 2, 1) == "Test1"
        assert wb2.get_cell_value("Imported", 2, 2) == 100

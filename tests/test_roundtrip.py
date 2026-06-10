"""Tests for saving and loading workbooks (roundtrip)."""

import pytest
import rustypyxl

from conftest import FIXTURE_NAMES


class TestRoundtrip:
    """Test saving and loading workbooks."""

    def test_save_load_empty_sheet(self, temp_xlsx_path):
        """Save and load a workbook with an empty sheet."""
        wb1 = rustypyxl.Workbook()
        wb1.create_sheet("Empty")
        wb1.save(temp_xlsx_path)

        wb2 = rustypyxl.load_workbook(temp_xlsx_path)
        assert "Empty" in wb2.sheetnames

    def test_save_load_multiple_sheets(self, temp_xlsx_path):
        """Save and load a workbook with multiple sheets."""
        wb1 = rustypyxl.Workbook()
        wb1.create_sheet("Sheet1")
        wb1.create_sheet("Sheet2")
        wb1.create_sheet("Sheet3")
        wb1.save(temp_xlsx_path)

        wb2 = rustypyxl.load_workbook(temp_xlsx_path)
        assert len(wb2) == 3
        assert "Sheet1" in wb2.sheetnames
        assert "Sheet2" in wb2.sheetnames
        assert "Sheet3" in wb2.sheetnames

    def test_save_load_special_sheet_names(self, temp_xlsx_path):
        """Save and load sheets with special characters in names."""
        wb1 = rustypyxl.Workbook()
        wb1.create_sheet("My Sheet")
        wb1.create_sheet("Data (2024)")
        wb1.save(temp_xlsx_path)

        wb2 = rustypyxl.load_workbook(temp_xlsx_path)
        assert "My Sheet" in wb2.sheetnames
        assert "Data (2024)" in wb2.sheetnames


class TestLoadExistingFiles:
    """Load+save+reload every generated feature fixture file."""

    @pytest.mark.parametrize("filename", FIXTURE_NAMES)
    def test_load_existing_xlsx(self, filename, fixtures_dir, temp_xlsx_path):
        """Externally-authored files load, resave, and reload cleanly."""
        wb = rustypyxl.load_workbook(str(fixtures_dir / filename))
        assert len(wb) > 0
        first_sheet = wb.sheetnames[0]
        a1 = wb[first_sheet]["A1"].value

        wb.save(temp_xlsx_path)
        wb2 = rustypyxl.load_workbook(temp_xlsx_path)
        assert wb2.sheetnames == wb.sheetnames
        assert wb2[first_sheet]["A1"].value == a1, "cell content changed on round-trip"


class TestFileIntegrity:
    """Test that saved files are valid xlsx format."""

    def test_saved_file_is_zip(self, temp_xlsx_path):
        """Verify saved file is a valid ZIP archive."""
        import zipfile

        wb = rustypyxl.Workbook()
        wb.create_sheet("Test")
        wb.save(temp_xlsx_path)

        assert zipfile.is_zipfile(temp_xlsx_path)

    def test_saved_file_has_required_parts(self, temp_xlsx_path):
        """Verify saved file contains required xlsx parts."""
        import zipfile

        wb = rustypyxl.Workbook()
        wb.create_sheet("Test")
        wb.save(temp_xlsx_path)

        with zipfile.ZipFile(temp_xlsx_path, 'r') as z:
            names = z.namelist()
            assert "[Content_Types].xml" in names
            assert "xl/workbook.xml" in names
            assert "_rels/.rels" in names

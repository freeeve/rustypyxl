"""Tests for data validation support."""

import openpyxl

import rustypyxl


class TestValidationRoundtrip:
    """Test data validation survives save/load cycle."""

    def test_load_existing_validation(self, fixtures_dir):
        """Load an externally-authored file with data validation."""
        wb = rustypyxl.load_workbook(str(fixtures_dir / "validation.xlsx"))
        assert wb.sheetnames == ["Validated"]
        assert wb["Validated"]["A1"].value == "pick one"

    def test_validation_survives_roundtrip(self, fixtures_dir, temp_xlsx_path):
        """The validation rule must survive load+save."""
        wb = rustypyxl.load_workbook(str(fixtures_dir / "validation.xlsx"))
        wb.save(temp_xlsx_path)

        chk = openpyxl.load_workbook(temp_xlsx_path)
        rules = chk["Validated"].data_validations.dataValidation
        assert len(rules) == 1, "validation rule stripped on round-trip"
        assert rules[0].type == "list"
        assert rules[0].formula1 == '"Yes,No,Maybe"'
        assert str(rules[0].sqref) == "B1:B10"

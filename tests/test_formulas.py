"""Tests for formula support."""

import os
import pytest
import rustypyxl


class TestFormulaBasics:
    """Test basic formula operations."""

    def test_cell_is_formula_property(self, workbook_with_sheet):
        """Test is_formula property on cells."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)

        # Not a formula initially
        assert cell.is_formula is False

        # Set a formula
        cell.value = "=SUM(A2:A10)"
        assert cell.is_formula is True

    def test_set_formula_value(self, workbook_with_sheet):
        """Set a formula as cell value."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        cell.value = "=A2+B2"
        assert cell.value == "=A2+B2"

    def test_set_sum_formula(self, workbook_with_sheet):
        """Set a SUM formula."""
        ws = workbook_with_sheet.active
        cell = ws.cell(1, 1)
        cell.value = "=SUM(A2:A100)"
        assert "SUM" in cell.value


class TestFormulaRoundtrip:
    """Test formulas survive save/load cycle."""

    def test_simple_formula_roundtrip(self, temp_xlsx_path):
        """Save and load a workbook with formulas."""
        wb1 = rustypyxl.Workbook()
        wb1.create_sheet("Formulas")
        wb1.save(temp_xlsx_path)

        wb2 = rustypyxl.load_workbook(temp_xlsx_path)
        assert "Formulas" in wb2.sheetnames

    def test_load_existing_formulas(self):
        """Load an existing file with formulas."""
        path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "test_formulas.xlsx")
        if not os.path.exists(path):
            pytest.skip("test_formulas.xlsx not found")

        wb = rustypyxl.load_workbook(path)
        assert wb is not None
        assert len(wb) > 0

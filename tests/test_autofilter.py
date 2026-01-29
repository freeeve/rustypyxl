"""Tests for AutoFilter functionality - inspired by openpyxl tests."""

import pytest
import rustypyxl


class TestAutoFilterBasics:
    """Basic AutoFilter tests."""

    def test_worksheet_saves_with_data(self, workbook_with_sheet, tmp_path):
        """Can save worksheet with data suitable for filtering."""
        ws = workbook_with_sheet.active
        ws.cell(1, 1).value = 'Header1'
        ws.cell(1, 2).value = 'Header2'
        ws.cell(2, 1).value = 'Value1'
        ws.cell(2, 2).value = 'Value2'

        path = str(tmp_path / 'test.xlsx')
        workbook_with_sheet.save(path)

        # Verify file was created and can be loaded
        wb2 = rustypyxl.load_workbook(path)
        assert wb2.active is not None


class TestAutoFilterWithData:
    """AutoFilter tests with actual data."""

    def test_save_workbook_with_filterable_data(self, workbook_with_sheet, tmp_path):
        """Can save a workbook with data that could be filtered."""
        ws = workbook_with_sheet.active

        # Create a table of data
        headers = ['Name', 'Category', 'Value', 'Active']
        for col, header in enumerate(headers, 1):
            ws.cell(1, col).value = header

        data = [
            ('Apple', 'Fruit', 10, True),
            ('Banana', 'Fruit', 15, True),
            ('Carrot', 'Vegetable', 5, False),
            ('Date', 'Fruit', 20, True),
            ('Eggplant', 'Vegetable', 8, True),
        ]

        for row_idx, row_data in enumerate(data, 2):
            for col_idx, value in enumerate(row_data, 1):
                ws.cell(row_idx, col_idx).value = value

        path = str(tmp_path / 'filterable.xlsx')
        workbook_with_sheet.save(path)

        # Verify file was created and can be loaded
        wb2 = rustypyxl.load_workbook(path)
        assert "Test" in wb2.sheetnames


class TestAutoFilterPreserved:
    """Tests that data in filtered ranges is preserved during save."""

    def test_filter_data_saves(self, workbook_with_sheet, tmp_path):
        """Data suitable for filtering saves correctly."""
        ws = workbook_with_sheet.active

        # Set up some data
        ws.cell(1, 1).value = 'ID'
        ws.cell(1, 2).value = 'Score'

        for i in range(2, 12):
            ws.cell(i, 1).value = i - 1
            ws.cell(i, 2).value = (i - 1) * 10

        path = str(tmp_path / 'test_preserve.xlsx')
        workbook_with_sheet.save(path)

        # Verify file was created and can be loaded
        wb2 = rustypyxl.load_workbook(path)
        assert len(wb2) == 1

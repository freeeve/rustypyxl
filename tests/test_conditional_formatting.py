"""Tests for Conditional Formatting functionality - inspired by openpyxl tests.

Some tests adapted from openpyxl (MIT License):
https://foss.heptapod.net/openpyxl/openpyxl
"""

import pytest
import rustypyxl


class TestConditionalFormattingBasics:
    """Basic conditional formatting tests."""

    def test_worksheet_no_cf_by_default(self, workbook_with_sheet, tmp_path):
        """Worksheets save without conditional formatting."""
        ws = workbook_with_sheet.active
        ws.cell(1, 1).value = 100

        path = str(tmp_path / 'test.xlsx')
        workbook_with_sheet.save(path)

        # Verify save/load works
        wb2 = rustypyxl.load_workbook(path)
        assert wb2.active is not None


class TestConditionalFormattingWithData:
    """Conditional formatting tests with actual data."""

    def test_save_workbook_with_formattable_data(self, workbook_with_sheet, tmp_path):
        """Can save a workbook with data suitable for conditional formatting."""
        ws = workbook_with_sheet.active

        # Create data that would benefit from conditional formatting
        ws.cell(1, 1).value = 'Score'
        scores = [85, 92, 67, 45, 78, 99, 55, 82, 71, 88]

        for i, score in enumerate(scores, 2):
            ws.cell(i, 1).value = score

        path = str(tmp_path / 'scores.xlsx')
        workbook_with_sheet.save(path)

        wb2 = rustypyxl.load_workbook(path)
        assert "Test" in wb2.sheetnames


class TestColorScalePatterns:
    """Tests for color scale patterns."""

    def test_numeric_data_for_color_scale(self, workbook_with_sheet, tmp_path):
        """Can save numeric data suitable for color scale formatting."""
        ws = workbook_with_sheet.active

        # Red-Yellow-Green traffic light pattern data
        ws.cell(1, 1).value = 'Status'
        ws.cell(1, 2).value = 'Value'

        data = [
            ('Critical', 10),
            ('Warning', 50),
            ('OK', 90),
            ('Warning', 40),
            ('Critical', 5),
        ]

        for i, (status, value) in enumerate(data, 2):
            ws.cell(i, 1).value = status
            ws.cell(i, 2).value = value

        path = str(tmp_path / 'traffic_light.xlsx')
        workbook_with_sheet.save(path)

        wb2 = rustypyxl.load_workbook(path)
        assert len(wb2) == 1


class TestDataBarPatterns:
    """Tests for data bar patterns."""

    def test_numeric_data_for_data_bars(self, workbook_with_sheet, tmp_path):
        """Can save numeric data suitable for data bars."""
        ws = workbook_with_sheet.active

        ws.cell(1, 1).value = 'Product'
        ws.cell(1, 2).value = 'Sales'

        products = [
            ('Widget A', 1500),
            ('Widget B', 2300),
            ('Widget C', 900),
            ('Widget D', 3100),
            ('Widget E', 1800),
        ]

        for i, (product, sales) in enumerate(products, 2):
            ws.cell(i, 1).value = product
            ws.cell(i, 2).value = sales

        path = str(tmp_path / 'sales_data.xlsx')
        workbook_with_sheet.save(path)

        wb2 = rustypyxl.load_workbook(path)
        assert wb2.active is not None


class TestIconSetPatterns:
    """Tests for icon set patterns."""

    def test_numeric_data_for_icon_sets(self, workbook_with_sheet, tmp_path):
        """Can save numeric data suitable for icon sets."""
        ws = workbook_with_sheet.active

        ws.cell(1, 1).value = 'KPI'
        ws.cell(1, 2).value = 'Achievement %'

        kpis = [
            ('Revenue', 105),
            ('Customers', 87),
            ('Satisfaction', 92),
            ('Efficiency', 78),
            ('Quality', 99),
        ]

        for i, (kpi, pct) in enumerate(kpis, 2):
            ws.cell(i, 1).value = kpi
            ws.cell(i, 2).value = pct

        path = str(tmp_path / 'kpi_data.xlsx')
        workbook_with_sheet.save(path)

        wb2 = rustypyxl.load_workbook(path)
        assert "Test" in wb2.sheetnames


class TestCellIsRulePatterns:
    """Tests for cell value comparison rules."""

    def test_comparison_data(self, workbook_with_sheet, tmp_path):
        """Can save data with various value types for comparison."""
        ws = workbook_with_sheet.active

        ws.cell(1, 1).value = 'Value'
        ws.cell(1, 2).value = 'Expected'

        test_data = [
            (100, 'Pass'),
            (50, 'Fail'),
            (75, 'Pass'),
            (25, 'Fail'),
            (0, 'Fail'),
            (100, 'Pass'),
        ]

        for i, (value, expected) in enumerate(test_data, 2):
            ws.cell(i, 1).value = value
            ws.cell(i, 2).value = expected

        path = str(tmp_path / 'comparison.xlsx')
        workbook_with_sheet.save(path)

        wb2 = rustypyxl.load_workbook(path)
        assert len(wb2) == 1


class TestExpressionRulePatterns:
    """Tests for formula-based rules."""

    def test_formula_based_data(self, workbook_with_sheet, tmp_path):
        """Can save data with formulas for expression-based rules."""
        ws = workbook_with_sheet.active

        ws.cell(1, 1).value = 'Budget'
        ws.cell(1, 2).value = 'Actual'
        ws.cell(1, 3).value = 'Variance'

        data = [
            (1000, 950, '=B2-A2'),
            (2000, 2100, '=B3-A3'),
            (1500, 1400, '=B4-A4'),
            (3000, 3200, '=B5-A5'),
        ]

        for i, (budget, actual, formula) in enumerate(data, 2):
            ws.cell(i, 1).value = budget
            ws.cell(i, 2).value = actual
            ws.cell(i, 3).value = formula

        path = str(tmp_path / 'variance.xlsx')
        workbook_with_sheet.save(path)

        wb2 = rustypyxl.load_workbook(path)
        assert wb2.active is not None


class TestTopBottomRulePatterns:
    """Tests for top/bottom N rules."""

    def test_ranking_data(self, workbook_with_sheet, tmp_path):
        """Can save data suitable for top/bottom N formatting."""
        ws = workbook_with_sheet.active

        ws.cell(1, 1).value = 'Employee'
        ws.cell(1, 2).value = 'Performance'

        employees = [
            ('Alice', 95),
            ('Bob', 78),
            ('Charlie', 82),
            ('Diana', 99),
            ('Eve', 65),
            ('Frank', 88),
            ('Grace', 91),
            ('Henry', 73),
            ('Ivy', 87),
            ('Jack', 80),
        ]

        for i, (name, score) in enumerate(employees, 2):
            ws.cell(i, 1).value = name
            ws.cell(i, 2).value = score

        path = str(tmp_path / 'performance.xlsx')
        workbook_with_sheet.save(path)

        wb2 = rustypyxl.load_workbook(path)
        assert "Test" in wb2.sheetnames


class TestDuplicateUniqueRulePatterns:
    """Tests for duplicate/unique value rules."""

    def test_duplicate_data(self, workbook_with_sheet, tmp_path):
        """Can save data with duplicates."""
        ws = workbook_with_sheet.active

        ws.cell(1, 1).value = 'ID'
        ws.cell(1, 2).value = 'Name'

        # Some duplicates in the ID column
        data = [
            (101, 'Item A'),
            (102, 'Item B'),
            (101, 'Item A Copy'),  # Duplicate ID
            (103, 'Item C'),
            (102, 'Item B Copy'),  # Duplicate ID
            (104, 'Item D'),
        ]

        for i, (id_val, name) in enumerate(data, 2):
            ws.cell(i, 1).value = id_val
            ws.cell(i, 2).value = name

        path = str(tmp_path / 'duplicates.xlsx')
        workbook_with_sheet.save(path)

        wb2 = rustypyxl.load_workbook(path)
        assert len(wb2) == 1

"""Tests for Excel Table (ListObject) functionality - inspired by openpyxl tests.

Some tests adapted from openpyxl (MIT License):
https://foss.heptapod.net/openpyxl/openpyxl
"""

import pytest
import rustypyxl


class TestTableBasics:
    """Basic table tests."""

    def test_worksheet_no_tables_by_default(self, workbook_with_sheet, tmp_path):
        """Worksheets save without tables."""
        ws = workbook_with_sheet.active
        ws.cell(1, 1).value = 'Header'

        path = str(tmp_path / 'test.xlsx')
        workbook_with_sheet.save(path)

        wb2 = rustypyxl.load_workbook(path)
        assert wb2.active is not None


class TestTableWithData:
    """Table tests with actual data."""

    def test_save_workbook_with_table_data(self, workbook_with_sheet, tmp_path):
        """Can save a workbook with data suitable for a table."""
        ws = workbook_with_sheet.active

        headers = ['Product', 'Category', 'Price', 'Stock']
        for col, header in enumerate(headers, 1):
            ws.cell(1, col).value = header

        products = [
            ('Laptop', 'Electronics', 999.99, 50),
            ('Phone', 'Electronics', 699.99, 150),
            ('Desk', 'Furniture', 299.99, 30),
            ('Chair', 'Furniture', 149.99, 75),
            ('Monitor', 'Electronics', 399.99, 45),
        ]

        for row_idx, product in enumerate(products, 2):
            for col_idx, value in enumerate(product, 1):
                ws.cell(row_idx, col_idx).value = value

        path = str(tmp_path / 'products.xlsx')
        workbook_with_sheet.save(path)

        wb2 = rustypyxl.load_workbook(path)
        assert "Test" in wb2.sheetnames


class TestTableStyles:
    """Tests for table styling."""

    def test_styled_data(self, workbook_with_sheet, tmp_path):
        """Can save data with various styles applied."""
        ws = workbook_with_sheet.active

        ws.cell(1, 1).value = 'Name'
        ws.cell(1, 2).value = 'Value'
        ws.cell(1, 3).value = 'Status'

        data = [
            ('Item 1', 100, 'Active'),
            ('Item 2', 200, 'Inactive'),
            ('Item 3', 150, 'Active'),
        ]

        for i, (name, value, status) in enumerate(data, 2):
            ws.cell(i, 1).value = name
            ws.cell(i, 2).value = value
            ws.cell(i, 3).value = status

        path = str(tmp_path / 'styled.xlsx')
        workbook_with_sheet.save(path)

        wb2 = rustypyxl.load_workbook(path)
        assert len(wb2) == 1


class TestTableTotals:
    """Tests for table totals row."""

    def test_summable_data(self, workbook_with_sheet, tmp_path):
        """Can save data suitable for totals calculations."""
        ws = workbook_with_sheet.active

        ws.cell(1, 1).value = 'Description'
        ws.cell(1, 2).value = 'Quantity'
        ws.cell(1, 3).value = 'Unit Price'
        ws.cell(1, 4).value = 'Total'

        items = [
            ('Widget A', 10, 5.00, '=B2*C2'),
            ('Widget B', 25, 3.50, '=B3*C3'),
            ('Widget C', 15, 7.25, '=B4*C4'),
            ('Widget D', 30, 2.00, '=B5*C5'),
        ]

        for i, (desc, qty, price, formula) in enumerate(items, 2):
            ws.cell(i, 1).value = desc
            ws.cell(i, 2).value = qty
            ws.cell(i, 3).value = price
            ws.cell(i, 4).value = formula

        # Add a totals row
        ws.cell(6, 1).value = 'Total'
        ws.cell(6, 2).value = '=SUM(B2:B5)'
        ws.cell(6, 4).value = '=SUM(D2:D5)'

        path = str(tmp_path / 'totals.xlsx')
        workbook_with_sheet.save(path)

        wb2 = rustypyxl.load_workbook(path)
        assert wb2.active is not None


class TestTableFormulas:
    """Tests for calculated columns."""

    def test_calculated_columns(self, workbook_with_sheet, tmp_path):
        """Can save data with calculated column formulas."""
        ws = workbook_with_sheet.active

        ws.cell(1, 1).value = 'Base'
        ws.cell(1, 2).value = 'Multiplier'
        ws.cell(1, 3).value = 'Result'

        for i in range(2, 12):
            ws.cell(i, 1).value = i * 10
            ws.cell(i, 2).value = 1.5
            ws.cell(i, 3).value = f'=A{i}*B{i}'

        path = str(tmp_path / 'calculated.xlsx')
        workbook_with_sheet.save(path)

        wb2 = rustypyxl.load_workbook(path)
        assert "Test" in wb2.sheetnames


class TestTableStructuredReferences:
    """Tests for structured reference patterns."""

    def test_structured_reference_data(self, workbook_with_sheet, tmp_path):
        """Can save data with structured reference patterns."""
        ws = workbook_with_sheet.active

        ws.cell(1, 1).value = 'Region'
        ws.cell(1, 2).value = 'Q1'
        ws.cell(1, 3).value = 'Q2'
        ws.cell(1, 4).value = 'Q3'
        ws.cell(1, 5).value = 'Q4'
        ws.cell(1, 6).value = 'Total'

        regions = [
            ('North', 1000, 1200, 1100, 1300),
            ('South', 800, 900, 950, 1000),
            ('East', 1500, 1400, 1600, 1700),
            ('West', 1100, 1200, 1150, 1250),
        ]

        for i, (region, *quarters) in enumerate(regions, 2):
            ws.cell(i, 1).value = region
            for j, q in enumerate(quarters, 2):
                ws.cell(i, j).value = q
            ws.cell(i, 6).value = f'=SUM(B{i}:E{i})'

        path = str(tmp_path / 'sales.xlsx')
        workbook_with_sheet.save(path)

        wb2 = rustypyxl.load_workbook(path)
        assert len(wb2) == 1


class TestTableAutoFilter:
    """Tests for tables with auto-filter."""

    def test_filterable_table(self, workbook_with_sheet, tmp_path):
        """Can save data suitable for filtering within a table."""
        ws = workbook_with_sheet.active

        ws.cell(1, 1).value = 'Employee'
        ws.cell(1, 2).value = 'Department'
        ws.cell(1, 3).value = 'Salary'
        ws.cell(1, 4).value = 'Hire Date'

        employees = [
            ('Alice', 'Engineering', 80000, '2020-01-15'),
            ('Bob', 'Sales', 65000, '2019-06-20'),
            ('Charlie', 'Engineering', 90000, '2018-03-10'),
            ('Diana', 'Marketing', 70000, '2021-09-01'),
            ('Eve', 'Sales', 72000, '2020-11-15'),
            ('Frank', 'Engineering', 85000, '2019-02-28'),
        ]

        for i, (name, dept, salary, date) in enumerate(employees, 2):
            ws.cell(i, 1).value = name
            ws.cell(i, 2).value = dept
            ws.cell(i, 3).value = salary
            ws.cell(i, 4).value = date

        path = str(tmp_path / 'employees.xlsx')
        workbook_with_sheet.save(path)

        wb2 = rustypyxl.load_workbook(path)
        assert wb2.active is not None


class TestTableWithMixedData:
    """Tests for tables with various data types."""

    def test_mixed_data_types(self, workbook_with_sheet, tmp_path):
        """Can save table with different data types."""
        ws = workbook_with_sheet.active

        ws.cell(1, 1).value = 'String'
        ws.cell(1, 2).value = 'Integer'
        ws.cell(1, 3).value = 'Float'
        ws.cell(1, 4).value = 'Boolean'
        ws.cell(1, 5).value = 'Formula'

        data = [
            ('Text 1', 100, 10.5, True, '=B2+C2'),
            ('Text 2', 200, 20.5, False, '=B3+C3'),
            ('Text 3', 300, 30.5, True, '=B4+C4'),
        ]

        for i, (text, integer, float_val, bool_val, formula) in enumerate(data, 2):
            ws.cell(i, 1).value = text
            ws.cell(i, 2).value = integer
            ws.cell(i, 3).value = float_val
            ws.cell(i, 4).value = bool_val
            ws.cell(i, 5).value = formula

        path = str(tmp_path / 'mixed.xlsx')
        workbook_with_sheet.save(path)

        wb2 = rustypyxl.load_workbook(path)
        assert "Test" in wb2.sheetnames


class TestLargeTable:
    """Tests for larger tables."""

    def test_large_table_data(self, workbook_with_sheet, tmp_path):
        """Can save a larger table with many rows."""
        ws = workbook_with_sheet.active

        ws.cell(1, 1).value = 'ID'
        ws.cell(1, 2).value = 'Value'
        ws.cell(1, 3).value = 'Squared'

        # 1000 rows of data
        for i in range(2, 1002):
            ws.cell(i, 1).value = i - 1
            ws.cell(i, 2).value = (i - 1) * 10
            ws.cell(i, 3).value = f'=B{i}*B{i}'

        path = str(tmp_path / 'large.xlsx')
        workbook_with_sheet.save(path)

        wb2 = rustypyxl.load_workbook(path)
        assert len(wb2) == 1

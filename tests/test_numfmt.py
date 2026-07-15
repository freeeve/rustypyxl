"""format_value renders a value under an Excel number-format code, matching
what Excel would display -- and Cell.display_value applies the cell's own
format. This is a capability openpyxl lacks (it only stores the code).
"""

import datetime

import pytest
import rustypyxl


@pytest.mark.parametrize(
    "value,code,expected",
    [
        (1234.0, "0", "1234"),
        (1234.5, "0.00", "1234.50"),
        (1234567.0, "#,##0", "1,234,567"),
        (0.1234, "0.00%", "12.34%"),
        (1234.5, "$#,##0.00", "$1,234.50"),
        (-1234.0, "#,##0;(#,##0)", "(1,234)"),
        (1_500_000.0, "#,##0,", "1,500"),
        (0.5, "#.##", ".5"),
        ("hello", "@", "hello"),
        ("hi", '"<"@">"', "<hi>"),
        (True, "0", "TRUE"),
    ],
)
def test_format_value_numbers_and_text(value, code, expected):
    assert rustypyxl.format_value(value, code) == expected


def test_format_value_dates():
    d = datetime.date(2023, 1, 15)  # a Sunday
    assert rustypyxl.format_value(d, "yyyy-mm-dd") == "2023-01-15"
    assert rustypyxl.format_value(d, "d-mmm-yyyy") == "15-Jan-2023"
    assert rustypyxl.format_value(d, "dddd") == "Sunday"


def test_format_value_datetime_with_time():
    dt = datetime.datetime(2023, 1, 15, 14, 30, 0)
    assert rustypyxl.format_value(dt, "yyyy-mm-dd hh:mm") == "2023-01-15 14:30"
    assert rustypyxl.format_value(dt, "h:mm AM/PM") == "2:30 PM"


def test_cell_display_value_applies_cell_format():
    wb = rustypyxl.Workbook()
    ws = wb.create_sheet("S")
    ws["A1"] = 0.1234
    ws["A1"].number_format = "0.00%"
    assert ws["A1"].display_value == "12.34%"


def test_cell_display_value_defaults_to_general():
    wb = rustypyxl.Workbook()
    ws = wb.create_sheet("S")
    ws["A1"] = 42
    assert ws["A1"].display_value == "42"

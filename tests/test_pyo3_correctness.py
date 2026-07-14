"""Regression tests for openpyxl-compatibility divergences in the bindings."""

import datetime

import pytest
import rustypyxl


class TestActiveAfterRemove:
    """wb.active must keep pointing at the sheet it pointed at."""

    def test_removing_an_earlier_sheet_keeps_the_active_sheet(self):
        wb = rustypyxl.Workbook()
        for name in ("A", "B", "C"):
            wb.create_sheet(name)
        wb.active = 1
        assert wb.active.title == "B"

        wb.remove(wb["A"])

        assert wb.active.title == "B", "active must follow the sheet, not the index"

    def test_removing_the_active_sheet_lands_on_the_next_one(self):
        wb = rustypyxl.Workbook()
        for name in ("A", "B", "C"):
            wb.create_sheet(name)
        wb.active = 1

        wb.remove(wb["B"])

        # openpyxl leaves the index in place, so it now names the next sheet
        assert wb.active.title == "C"

    def test_removing_a_later_sheet_leaves_the_active_sheet_alone(self):
        wb = rustypyxl.Workbook()
        for name in ("A", "B", "C"):
            wb.create_sheet(name)
        wb.active = 1

        wb.remove(wb["C"])

        assert wb.active.title == "B"

    def test_removing_the_last_sheet_clamps_the_active_index(self):
        wb = rustypyxl.Workbook()
        for name in ("A", "B"):
            wb.create_sheet(name)
        wb.active = 1

        wb.remove(wb["B"])

        assert wb.active.title == "A"

    def test_writes_after_remove_land_on_the_right_sheet(self):
        """The bug that made this matter: active pointed at the wrong sheet."""
        wb = rustypyxl.Workbook()
        for name in ("A", "B", "C"):
            wb.create_sheet(name)
        wb.active = 1
        wb.remove(wb["A"])

        wb.active["A1"] = "written"

        assert wb["B"]["A1"].value == "written"
        assert wb["C"]["A1"].value is None


class TestColorCoercion:
    """Theme/indexed/tint colors are not representable; say so, don't drop them."""

    def test_theme_color_raises_instead_of_producing_a_colorless_font(self):
        with pytest.raises(ValueError, match="theme"):
            rustypyxl.Font(color=rustypyxl.Color(theme=1))

    def test_indexed_color_raises(self):
        with pytest.raises(ValueError, match="indexed"):
            rustypyxl.Font(color=rustypyxl.Color(indexed=3))

    def test_tint_raises(self):
        with pytest.raises(ValueError, match="tint"):
            rustypyxl.Font(color=rustypyxl.Color(rgb="FFFF0000", tint=0.4))

    def test_rgb_color_object_still_works(self):
        font = rustypyxl.Font(color=rustypyxl.Color(rgb="FFFF0000"))
        assert font.color == "FFFF0000"

    def test_rgb_string_still_works(self):
        assert rustypyxl.Font(color="FF00FF00").color == "FF00FF00"


class TestDataType:
    """openpyxl reports 'd' for datetime cells."""

    def test_datetime_cell_reports_d(self, workbook_with_sheet):
        ws = workbook_with_sheet["Test"]
        ws["A1"] = datetime.datetime(2024, 3, 1, 12, 30, 0)
        assert ws["A1"].data_type == "d"

    def test_date_cell_reports_d(self, workbook_with_sheet):
        ws = workbook_with_sheet["Test"]
        ws["A1"] = datetime.date(2024, 3, 1)
        assert ws["A1"].data_type == "d"

    def test_other_types_are_unchanged(self, workbook_with_sheet):
        ws = workbook_with_sheet["Test"]
        ws["A1"] = "text"
        ws["A2"] = 42
        ws["A3"] = True
        ws["A4"] = "=SUM(A2:A2)"
        assert ws["A1"].data_type == "s"
        assert ws["A2"].data_type == "n"
        assert ws["A3"].data_type == "b"
        assert ws["A4"].data_type == "f"


class Reentrant:
    """Its __str__ reads back from the workbook it is being written into."""

    def __init__(self, workbook, sheet):
        self.workbook = workbook
        self.sheet = sheet

    def __str__(self):
        # Touching the workbook here used to raise "Already borrowed", because
        # the conversion ran while the workbook was mutably borrowed.
        return f"seen:{self.workbook[self.sheet]['A1'].value}"


class TestReentrancy:
    """Converting a value runs __str__, which may re-enter the workbook."""

    def test_setitem_tolerates_a_reentrant_str(self, workbook_with_sheet):
        wb = workbook_with_sheet
        ws = wb["Test"]
        ws["A1"] = "anchor"

        ws["B1"] = Reentrant(wb, "Test")

        assert ws["B1"].value == "seen:anchor"

    def test_cell_value_setter_tolerates_a_reentrant_str(self, workbook_with_sheet):
        wb = workbook_with_sheet
        ws = wb["Test"]
        ws["A1"] = "anchor"

        ws.cell(row=2, column=2).value = Reentrant(wb, "Test")

        assert ws["B2"].value == "seen:anchor"

    def test_write_rows_tolerates_a_reentrant_str(self, workbook_with_sheet):
        wb = workbook_with_sheet
        wb["Test"]["A1"] = "anchor"

        wb.write_rows("Test", [[Reentrant(wb, "Test")]], start_row=3, start_col=1)

        assert wb["Test"]["A3"].value == "seen:anchor"

    def test_set_cell_value_tolerates_a_reentrant_str(self, workbook_with_sheet):
        wb = workbook_with_sheet
        wb["Test"]["A1"] = "anchor"

        wb.set_cell_value("Test", 4, 1, Reentrant(wb, "Test"))

        assert wb["Test"]["A4"].value == "seen:anchor"

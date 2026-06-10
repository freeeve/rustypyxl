"""Worksheet/cell handles must stay pinned to their sheet across workbook
mutations (remove, insert, move, rename, copy) — or fail loudly when the
sheet is gone. Previously half the API resolved handles by positional index
and half by cached name, so removing a sheet made existing handles silently
read and write a *different* sheet.
"""

import pytest

import rustypyxl


@pytest.fixture
def three_sheets():
    wb = rustypyxl.Workbook()
    for name in ("Sheet1", "Sheet2", "Sheet3"):
        wb.create_sheet(name)
    return wb


class TestHandleStability:
    def test_handle_survives_removal_of_earlier_sheet(self, three_sheets):
        wb = three_sheets
        ws2 = wb["Sheet2"]
        wb.remove(wb["Sheet1"])
        ws2.append(["x"])
        assert wb["Sheet2"]["A1"].value == "x"
        assert wb["Sheet3"]["A1"].value is None, "write landed on the wrong sheet"

    def test_stale_handle_raises_instead_of_hitting_neighbor(self, three_sheets):
        wb = three_sheets
        gone = wb["Sheet3"]
        wb.remove(wb["Sheet3"])
        with pytest.raises(ValueError):
            gone.append(["boom"])
        with pytest.raises(ValueError):
            _ = gone.max_row
        with pytest.raises(ValueError):
            gone["A1"] = 1

    def test_stale_cell_handle_raises(self, three_sheets):
        wb = three_sheets
        cell = wb["Sheet3"]["A1"]
        wb.remove(wb["Sheet3"])
        with pytest.raises(ValueError):
            _ = cell.value
        with pytest.raises(ValueError):
            cell.value = 1

    def test_rename_does_not_orphan_handles(self):
        wb = rustypyxl.Workbook()
        ws = wb.create_sheet("Old")
        cell = ws["A1"]
        ws.title = "New"
        cell.value = 42
        assert wb["New"]["A1"].value == 42
        ws["B1"] = "via setitem"
        assert wb["New"]["B1"].value == "via setitem"
        assert ws.title == "New"

    def test_rename_via_other_handle_is_visible(self):
        wb = rustypyxl.Workbook()
        wb.create_sheet("Old")
        h1 = wb["Old"]
        h2 = wb["Old"]
        h1.title = "Renamed"
        assert h2.title == "Renamed"
        h2["A1"] = "still works"
        assert wb["Renamed"]["A1"].value == "still works"

    def test_insert_and_move_keep_handles_pinned(self):
        wb = rustypyxl.Workbook()
        wb.create_sheet("A")
        b = wb.create_sheet("B")
        wb.create_sheet("First", index=0)
        b.append([1, 2])
        assert wb["B"]["A1"].value == 1

        wb.move_sheet(b, -2)
        b.append([3, 4])
        assert wb["B"]["A2"].value == 3
        assert wb.index(b) == 0

    def test_copy_gets_distinct_identity(self):
        wb = rustypyxl.Workbook()
        wb.create_sheet("A")
        copy = wb.copy_worksheet(wb["A"])
        copy.append(["copied"])
        assert wb["A"]["A1"].value is None
        assert wb[copy.title]["A1"].value == "copied"

    def test_cell_setters_propagate_errors(self, three_sheets):
        wb = three_sheets
        cell = wb["Sheet2"]["A1"]
        wb.remove(wb["Sheet2"])
        # every connected setter must raise, not silently no-op
        with pytest.raises(ValueError):
            cell.value = 1
        with pytest.raises(ValueError):
            cell.number_format = "0.00"
        with pytest.raises(ValueError):
            cell.hyperlink = "https://example.com"
        with pytest.raises(ValueError):
            cell.comment = "note"


class TestNumberFormatClear:
    def test_assigning_none_clears_format(self):
        wb = rustypyxl.Workbook()
        ws = wb.create_sheet("A")
        c = ws["A1"]
        c.value = 1.5
        c.number_format = "0.00%"
        assert c.number_format == "0.00%"
        c.number_format = None
        assert c.number_format is None

    def test_clearing_keeps_other_style_properties(self, tmp_path):
        from rustypyxl import Font

        wb = rustypyxl.Workbook()
        ws = wb.create_sheet("A")
        c = ws["A1"]
        c.value = 1.5
        c.font = Font(bold=True)
        c.number_format = "0.00%"
        c.number_format = None
        assert c.number_format is None
        assert c.font is not None and c.font.bold

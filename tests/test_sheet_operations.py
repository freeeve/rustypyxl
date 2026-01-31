"""Tests for sheet add/remove/rename operations."""

import io
import pytest
import random
import string
from rustypyxl import Workbook


class TestSheetAddRemove:
    """Test basic sheet add and remove operations."""

    def test_add_single_sheet(self):
        """Adding a single sheet works."""
        wb = Workbook()
        ws = wb.create_sheet("Test")
        assert "Test" in wb.sheetnames
        assert len(wb.sheetnames) == 1

    def test_add_multiple_sheets(self):
        """Adding multiple sheets works."""
        wb = Workbook()
        wb.create_sheet("Sheet1")
        wb.create_sheet("Sheet2")
        wb.create_sheet("Sheet3")
        assert wb.sheetnames == ["Sheet1", "Sheet2", "Sheet3"]

    def test_remove_sheet(self):
        """Removing a sheet works."""
        wb = Workbook()
        ws1 = wb.create_sheet("Keep")
        ws2 = wb.create_sheet("Remove")
        wb.remove(ws2)
        assert "Remove" not in wb.sheetnames
        assert "Keep" in wb.sheetnames
        assert len(wb.sheetnames) == 1

    def test_remove_and_readd_same_name(self):
        """Removing and re-adding a sheet with the same name works."""
        wb = Workbook()
        ws = wb.create_sheet("TestSheet")
        wb.set_cell_value("TestSheet", 1, 1, "Original")
        wb.remove(ws)

        # Re-add with same name
        ws2 = wb.create_sheet("TestSheet")
        wb.set_cell_value("TestSheet", 1, 1, "New")

        assert "TestSheet" in wb.sheetnames
        assert wb.get_cell_value("TestSheet", 1, 1) == "New"

    def test_add_remove_multiple_cycles(self):
        """Multiple add/remove cycles work correctly."""
        wb = Workbook()

        for i in range(5):
            ws = wb.create_sheet(f"Cycle{i}")
            wb.set_cell_value(f"Cycle{i}", 1, 1, f"Value{i}")
            wb.remove(ws)

        # All should be removed
        assert len(wb.sheetnames) == 0

        # Add them back
        for i in range(5):
            wb.create_sheet(f"Final{i}")

        assert len(wb.sheetnames) == 5

    def test_remove_middle_sheet(self):
        """Removing a sheet from the middle works."""
        wb = Workbook()
        wb.create_sheet("First")
        ws2 = wb.create_sheet("Middle")
        wb.create_sheet("Last")

        wb.remove(ws2)

        assert wb.sheetnames == ["First", "Last"]


class TestSheetOperationsRoundtrip:
    """Test sheet operations survive save/load roundtrip."""

    def test_roundtrip_after_add(self):
        """Sheets survive roundtrip after adding."""
        wb = Workbook()
        wb.create_sheet("Sheet1")
        wb.create_sheet("Sheet2")
        wb.set_cell_value("Sheet1", 1, 1, "Data1")
        wb.set_cell_value("Sheet2", 1, 1, "Data2")

        data = wb.save_to_bytes()
        wb2 = Workbook.load(io.BytesIO(data))

        assert wb2.sheetnames == ["Sheet1", "Sheet2"]
        assert wb2.get_cell_value("Sheet1", 1, 1) == "Data1"
        assert wb2.get_cell_value("Sheet2", 1, 1) == "Data2"

    def test_roundtrip_after_remove(self):
        """Removed sheets don't appear after roundtrip."""
        wb = Workbook()
        wb.create_sheet("Keep")
        ws_remove = wb.create_sheet("Remove")
        wb.set_cell_value("Keep", 1, 1, "Kept")
        wb.set_cell_value("Remove", 1, 1, "Removed")
        wb.remove(ws_remove)

        data = wb.save_to_bytes()
        wb2 = Workbook.load(io.BytesIO(data))

        assert "Keep" in wb2.sheetnames
        assert "Remove" not in wb2.sheetnames
        assert wb2.get_cell_value("Keep", 1, 1) == "Kept"

    def test_multiple_roundtrips(self):
        """Multiple roundtrips preserve sheet state."""
        wb = Workbook()

        for roundtrip in range(3):
            wb.create_sheet(f"Round{roundtrip}")
            wb.set_cell_value(f"Round{roundtrip}", 1, 1, f"Value{roundtrip}")

            data = wb.save_to_bytes()
            wb = Workbook.load(io.BytesIO(data))

        assert len(wb.sheetnames) == 3
        for i in range(3):
            assert wb.get_cell_value(f"Round{i}", 1, 1) == f"Value{i}"


class TestSheetNameEdgeCases:
    """Test edge cases with sheet names."""

    def test_sheet_with_spaces(self):
        """Sheet names with spaces work."""
        wb = Workbook()
        wb.create_sheet("My Sheet")
        assert "My Sheet" in wb.sheetnames

    def test_sheet_with_unicode(self):
        """Sheet names with unicode characters work."""
        wb = Workbook()
        wb.create_sheet("シート1")
        wb.create_sheet("Données")
        wb.create_sheet("表格")

        data = wb.save_to_bytes()
        wb2 = Workbook.load(io.BytesIO(data))

        assert "シート1" in wb2.sheetnames
        assert "Données" in wb2.sheetnames
        assert "表格" in wb2.sheetnames

    def test_sheet_with_numbers(self):
        """Sheet names with numbers work."""
        wb = Workbook()
        wb.create_sheet("123")
        wb.create_sheet("Sheet123")

        assert "123" in wb.sheetnames
        assert "Sheet123" in wb.sheetnames


class TestFuzzSheetOperations:
    """Fuzz testing for sheet operations."""

    @pytest.mark.parametrize("seed", range(10))
    def test_random_add_remove_sequence(self, seed):
        """Random sequence of add/remove operations."""
        random.seed(seed)
        wb = Workbook()
        expected_sheets = []

        # Perform random operations
        for _ in range(20):
            op = random.choice(["add", "add", "add", "remove"])  # Bias toward add

            if op == "add":
                name = f"Sheet_{random.randint(0, 1000)}"
                # Avoid duplicate names
                if name not in expected_sheets:
                    wb.create_sheet(name)
                    expected_sheets.append(name)
            elif op == "remove" and expected_sheets:
                idx = random.randint(0, len(expected_sheets) - 1)
                name = expected_sheets[idx]
                ws = wb[name]
                wb.remove(ws)
                expected_sheets.remove(name)

        # Verify state
        assert sorted(wb.sheetnames) == sorted(expected_sheets)

        # Verify roundtrip
        data = wb.save_to_bytes()
        wb2 = Workbook.load(io.BytesIO(data))
        assert sorted(wb2.sheetnames) == sorted(expected_sheets)

    @pytest.mark.parametrize("seed", range(10))
    def test_random_add_remove_with_data(self, seed):
        """Random add/remove with data verification."""
        random.seed(seed)
        wb = Workbook()
        sheet_data = {}

        for _ in range(15):
            op = random.choice(["add", "add", "remove"])

            if op == "add":
                name = f"Data_{random.randint(0, 1000)}"
                if name not in sheet_data:
                    wb.create_sheet(name)
                    value = f"Value_{random.randint(0, 10000)}"
                    wb.set_cell_value(name, 1, 1, value)
                    sheet_data[name] = value
            elif op == "remove" and sheet_data:
                name = random.choice(list(sheet_data.keys()))
                ws = wb[name]
                wb.remove(ws)
                del sheet_data[name]

        # Roundtrip and verify
        data = wb.save_to_bytes()
        wb2 = Workbook.load(io.BytesIO(data))

        for name, expected_value in sheet_data.items():
            assert wb2.get_cell_value(name, 1, 1) == expected_value

    @pytest.mark.parametrize("seed", range(5))
    def test_remove_all_and_recreate(self, seed):
        """Remove all sheets and recreate."""
        random.seed(seed)
        wb = Workbook()

        # Create some sheets
        sheets = [f"Sheet_{i}" for i in range(5)]
        for name in sheets:
            wb.create_sheet(name)
            wb.set_cell_value(name, 1, 1, f"Data_{name}")

        # Remove all in random order
        random.shuffle(sheets)
        for name in sheets:
            ws = wb[name]
            wb.remove(ws)

        assert len(wb.sheetnames) == 0

        # Recreate
        new_sheets = [f"New_{i}" for i in range(3)]
        for name in new_sheets:
            wb.create_sheet(name)
            wb.set_cell_value(name, 1, 1, f"New_Data_{name}")

        # Roundtrip
        data = wb.save_to_bytes()
        wb2 = Workbook.load(io.BytesIO(data))

        assert sorted(wb2.sheetnames) == sorted(new_sheets)
        for name in new_sheets:
            assert wb2.get_cell_value(name, 1, 1) == f"New_Data_{name}"


class TestSheetOrderPreservation:
    """Test that sheet order is preserved."""

    def test_order_after_add(self):
        """Sheet order is preserved after adding."""
        wb = Workbook()
        names = ["Alpha", "Beta", "Gamma", "Delta"]
        for name in names:
            wb.create_sheet(name)

        assert wb.sheetnames == names

    def test_order_after_remove_middle(self):
        """Sheet order is preserved after removing middle sheet."""
        wb = Workbook()
        for name in ["A", "B", "C", "D", "E"]:
            wb.create_sheet(name)

        ws_c = wb["C"]
        wb.remove(ws_c)

        assert wb.sheetnames == ["A", "B", "D", "E"]

    def test_order_after_roundtrip(self):
        """Sheet order is preserved after roundtrip."""
        wb = Workbook()
        names = ["First", "Second", "Third", "Fourth"]
        for name in names:
            wb.create_sheet(name)

        data = wb.save_to_bytes()
        wb2 = Workbook.load(io.BytesIO(data))

        assert wb2.sheetnames == names

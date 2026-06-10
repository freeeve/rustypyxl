"""Round-trip fidelity tests: structure must survive load+save.

These are differential tests: openpyxl produces a feature-rich file,
rustypyxl loads and re-saves it, and openpyxl verifies nothing was lost.
"""

import pathlib

import openpyxl
import pytest
from openpyxl.worksheet.datavalidation import DataValidation as OxDV
from openpyxl.worksheet.table import Table as OxTable, TableStyleInfo

import rustypyxl


@pytest.fixture
def rich_source(tmp_path) -> pathlib.Path:
    """A feature-rich workbook produced by openpyxl."""
    path = tmp_path / "source.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Main"
    ws["A1"] = "data"
    ws["A3"] = "link"
    ws["A3"].hyperlink = "https://example.com/x?a=1&b=2"
    ws.freeze_panes = "C3"
    ws.auto_filter.ref = "A1:D20"

    dv = OxDV(type="list", formula1='"Yes,No"', allowBlank=True)
    dv.add("E2:E10")
    ws.add_data_validation(dv)

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9  # A4
    ws.print_options.gridLines = True

    ws["H1"] = "H1"
    ws["I1"] = "H2"
    ws["H2"] = "a"
    ws["I2"] = 2
    ws["H3"] = "b"
    ws["I3"] = 3
    tab = OxTable(displayName="OxTable1", ref="H1:I3")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(tab)

    hidden = wb.create_sheet("Hidden")
    hidden["A1"] = "secret"
    hidden.sheet_state = "hidden"
    wb.save(path)
    return path


@pytest.fixture
def resaved(rich_source, tmp_path) -> openpyxl.Workbook:
    """The source after a rustypyxl load+save cycle, reopened with openpyxl."""
    out = tmp_path / "resaved.xlsx"
    rwb = rustypyxl.load_workbook(str(rich_source))
    rwb.save(str(out))
    return openpyxl.load_workbook(out)


class TestStructureSurvivesRoundtrip:
    def test_cell_data(self, resaved):
        assert resaved["Main"]["A1"].value == "data"
        assert resaved["Hidden"]["A1"].value == "secret"

    def test_external_hyperlink(self, resaved):
        link = resaved["Main"]["A3"].hyperlink
        assert link is not None, "external hyperlink stripped on round-trip"
        assert link.target == "https://example.com/x?a=1&b=2"

    def test_freeze_panes(self, resaved):
        assert resaved["Main"].freeze_panes == "C3"

    def test_autofilter(self, resaved):
        assert resaved["Main"].auto_filter.ref == "A1:D20"

    def test_data_validation(self, resaved):
        rules = resaved["Main"].data_validations.dataValidation
        assert len(rules) == 1
        rule = rules[0]
        assert rule.type == "list"
        assert rule.formula1 == '"Yes,No"'
        assert str(rule.sqref) == "E2:E10"

    def test_page_setup(self, resaved):
        ws = resaved["Main"]
        assert ws.page_setup.orientation == "landscape"
        assert int(ws.page_setup.paperSize) == 9
        assert ws.print_options.gridLines

    def test_table(self, resaved):
        tables = {t.displayName: t for t in resaved["Main"].tables.values()}
        assert "OxTable1" in tables, "table stripped on round-trip"
        assert tables["OxTable1"].ref == "H1:I3"
        assert tables["OxTable1"].tableStyleInfo.name == "TableStyleMedium9"

    def test_hidden_sheet_state(self, resaved):
        assert resaved["Hidden"].sheet_state == "hidden", (
            "hidden sheet became visible on round-trip"
        )
        assert resaved["Main"].sheet_state == "visible"


class TestStructureApi:
    def test_sheet_state_property(self, rich_source):
        wb = rustypyxl.load_workbook(str(rich_source))
        assert wb["Hidden"].sheet_state == "hidden"
        assert wb["Main"].sheet_state == "visible"

        wb["Hidden"].sheet_state = "visible"
        assert wb["Hidden"].sheet_state == "visible"
        with pytest.raises(ValueError):
            wb["Main"].sheet_state = "invisible"

    def test_active_sheet_respects_loaded_tab(self, tmp_path):
        path = tmp_path / "active.xlsx"
        owb = openpyxl.Workbook()
        owb.create_sheet("Second")
        owb.active = 1
        owb.save(path)

        wb = rustypyxl.load_workbook(str(path))
        assert wb.active.title == "Second"

    def test_rustypyxl_written_hyperlink_works_in_openpyxl(self, tmp_path):
        path = tmp_path / "links.xlsx"
        wb = rustypyxl.Workbook()
        ws = wb.create_sheet("S")
        ws["A1"] = "site"
        wb.set_cell_hyperlink("S", 1, 1, "https://example.org/")
        wb.save(str(path))

        chk = openpyxl.load_workbook(path)
        link = chk["S"]["A1"].hyperlink
        assert link is not None
        assert link.target == "https://example.org/"

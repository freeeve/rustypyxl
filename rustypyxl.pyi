"""Type stubs for rustypyxl, a Rust-powered Excel library with an
openpyxl-compatible API."""

import datetime
import os
from typing import Any, BinaryIO, Iterator

CellValue = str | int | float | bool | datetime.datetime | datetime.date | datetime.time | None
_ColorArg = str | Color | None
# A color reads back as the plain hex string when that is all it is, and as a
# Color when it carries a theme, a palette index, or a tint.
_ColorValue = str | Color | None

def load_workbook(
    source: str | os.PathLike[str] | bytes | BinaryIO, password: str | None = None
) -> Workbook: ...
def format_value(
    value: str | int | float | bool | datetime.datetime | datetime.date | datetime.time | None,
    number_format: str,
) -> str: ...

class Workbook:
    def __init__(self) -> None: ...
    @staticmethod
    def load(
        source: str | os.PathLike[str] | bytes | BinaryIO, password: str | None = None
    ) -> Workbook: ...
    @property
    def active(self) -> Worksheet: ...
    @active.setter
    def active(self, value: Worksheet | int) -> None: ...
    @property
    def sheetnames(self) -> list[str]: ...
    @property
    def worksheets(self) -> list[Worksheet]: ...
    @property
    def defined_names(self) -> list[tuple[str, str]]: ...
    def __getitem__(self, key: str) -> Worksheet: ...
    def __contains__(self, key: str) -> bool: ...
    def __len__(self) -> int: ...
    def __iter__(self) -> Iterator[str]: ...
    def create_sheet(self, title: str | None = None, index: int | None = None) -> Worksheet: ...
    def remove(self, worksheet: Worksheet) -> None: ...
    def copy_worksheet(self, source: Worksheet) -> Worksheet: ...
    def move_sheet(self, sheet: Worksheet, offset: int) -> None: ...
    def index(self, worksheet: Worksheet) -> int: ...
    def create_named_range(self, name: str, worksheet: Worksheet, range: str) -> None: ...
    def save(
        self, filename: str | os.PathLike[str], password: str | None = None
    ) -> None: ...
    def save_to_bytes(self, password: str | None = None) -> bytes: ...
    def close(self) -> None: ...
    def set_compression(self, level: str) -> None: ...
    def write_rows(
        self,
        sheet_name: str,
        data: list[list[CellValue]],
        start_row: int = 1,
        start_col: int = 1,
    ) -> None: ...
    def read_rows(
        self,
        sheet_name: str,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
    ) -> list[list[CellValue]]: ...
    def get_cell_value(self, sheet_name: str, row: int, column: int) -> CellValue: ...
    def evaluate_formula(self, sheet_name: str, formula: str) -> Any: ...
    def evaluate_cell(self, sheet_name: str, row: int, column: int) -> Any: ...
    def calculate_all(self) -> int: ...
    @property
    def pivot_tables(self) -> list[PivotTable]: ...
    def add_pivot_table(
        self,
        source_sheet: str,
        source_ref: str,
        target_sheet: str,
        anchor: str,
        rows: list[str] = ...,
        columns: list[str] = ...,
        values: list[tuple[str, str]] = ...,
        name: str | None = None,
    ) -> None: ...
    def set_cell_value(self, sheet_name: str, row: int, column: int, value: CellValue) -> None: ...
    def get_cell_font(self, sheet_name: str, row: int, column: int) -> Font | None: ...
    def set_cell_font(self, sheet_name: str, row: int, column: int, font: Font) -> None: ...
    def get_cell_alignment(self, sheet_name: str, row: int, column: int) -> Alignment | None: ...
    def set_cell_alignment(self, sheet_name: str, row: int, column: int, alignment: Alignment) -> None: ...
    def get_cell_fill(self, sheet_name: str, row: int, column: int) -> PatternFill | None: ...
    def set_cell_fill(self, sheet_name: str, row: int, column: int, fill: PatternFill) -> None: ...
    def get_cell_border(self, sheet_name: str, row: int, column: int) -> Border | None: ...
    def set_cell_border(self, sheet_name: str, row: int, column: int, border: Border) -> None: ...
    def get_cell_protection(self, sheet_name: str, row: int, column: int) -> Protection | None: ...
    def set_cell_protection(self, sheet_name: str, row: int, column: int, protection: Protection) -> None: ...
    def get_cell_hyperlink(self, sheet_name: str, row: int, column: int) -> str | None: ...
    def set_cell_hyperlink(self, sheet_name: str, row: int, column: int, url: str | None = None) -> None: ...
    def get_cell_comment(self, sheet_name: str, row: int, column: int) -> str | None: ...
    def set_cell_comment(self, sheet_name: str, row: int, column: int, comment: str | None = None) -> None: ...
    def get_cell_number_format(self, sheet_name: str, row: int, column: int) -> str | None: ...
    def set_cell_number_format(self, sheet_name: str, row: int, column: int, format: str) -> None: ...
    def clear_cell_number_format(self, sheet_name: str, row: int, column: int) -> None: ...
    def set_cell_style(
        self,
        sheet_name: str,
        row: int,
        column: int,
        font: Font | None = None,
        fill: PatternFill | None = None,
        border: Border | None = None,
        alignment: Alignment | None = None,
        number_format: str | None = None,
    ) -> None: ...
    def insert_from_parquet(
        self,
        sheet_name: str,
        path: str,
        start_row: int = 1,
        start_col: int = 1,
        include_headers: bool = True,
        column_renames: dict[str, str] | None = None,
        columns: list[str] | None = None,
    ) -> dict[str, Any]: ...
    def export_to_parquet(
        self,
        sheet_name: str,
        path: str,
        has_headers: bool = True,
        compression: str = "snappy",
        column_renames: dict[str, str] | None = None,
        column_types: dict[str, str] | None = None,
    ) -> dict[str, Any]: ...
    def export_range_to_parquet(
        self,
        sheet_name: str,
        path: str,
        min_row: int,
        min_col: int,
        max_row: int,
        max_col: int,
        has_headers: bool = True,
        compression: str = "snappy",
    ) -> dict[str, Any]: ...

class Worksheet:
    title: str
    sheet_state: str
    freeze_panes: str | None
    @property
    def dimensions(self) -> str: ...
    @property
    def max_row(self) -> int: ...
    @property
    def max_column(self) -> int: ...
    @property
    def min_row(self) -> int: ...
    @property
    def min_column(self) -> int: ...
    @property
    def merged_cells(self) -> list[str]: ...
    def __getitem__(self, key: str) -> Any: ...
    def __setitem__(self, key: str, value: CellValue) -> None: ...
    def cell(self, row: int, column: int | None = None) -> Cell: ...
    def append(
        self,
        iterable: list[CellValue] | tuple[CellValue, ...] | Iterator[CellValue] | dict[str | int, CellValue],
    ) -> None: ...
    def iter_rows(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        values_only: bool = False,
    ) -> CellRangeIterator: ...
    def iter_cols(
        self,
        min_col: int | None = None,
        max_col: int | None = None,
        min_row: int | None = None,
        max_row: int | None = None,
        values_only: bool = False,
    ) -> CellRangeIterator: ...
    def merge_cells(
        self,
        range_string: str | None = None,
        start_row: int | None = None,
        start_column: int | None = None,
        end_row: int | None = None,
        end_column: int | None = None,
    ) -> None: ...
    def unmerge_cells(
        self,
        range_string: str | None = None,
        start_row: int | None = None,
        start_column: int | None = None,
        end_row: int | None = None,
        end_column: int | None = None,
    ) -> None: ...
    def auto_fit_column(self, column: int) -> float | None: ...
    def auto_fit_all(self) -> None: ...
    @property
    def auto_filter(self) -> AutoFilter: ...
    @property
    def sheet_protected(self) -> bool: ...
    def protect_sheet(self, password: str | None = None) -> None: ...
    def unprotect_sheet(self) -> None: ...
    @property
    def column_dimensions(self) -> ColumnDimensions: ...
    @property
    def row_dimensions(self) -> RowDimensions: ...
    @property
    def tables(self) -> list[dict[str, str]]: ...
    @property
    def data_validations(self) -> list[dict[str, Any]]: ...
    print_area: str | None
    def set_page_setup(
        self,
        orientation: str | None = None,
        paper_size: str | None = None,
        scale: int | None = None,
        fit_to_width: int | None = None,
        fit_to_height: int | None = None,
        print_gridlines: bool | None = None,
        center: bool | None = None,
    ) -> None: ...
    def set_page_margins(
        self,
        left: float = 0.7,
        right: float = 0.7,
        top: float = 0.75,
        bottom: float = 0.75,
        header: float = 0.3,
        footer: float = 0.3,
    ) -> None: ...
    def set_header_footer(
        self,
        header_left: str | None = None,
        header_center: str | None = None,
        header_right: str | None = None,
        footer_left: str | None = None,
        footer_center: str | None = None,
        footer_right: str | None = None,
    ) -> None: ...
    def add_conditional_formatting(self, cells: str, rule: dict[str, Any]) -> None: ...
    def add_data_validation(
        self,
        cells: str,
        type: str,
        formula1: str | None = None,
        formula2: str | None = None,
        operator: str | None = None,
        allow_blank: bool = True,
        show_error: bool = True,
        error_title: str | None = None,
        error: str | None = None,
        show_input: bool = True,
        prompt_title: str | None = None,
        prompt: str | None = None,
    ) -> None: ...
    def add_table(
        self,
        name: str,
        ref: str,
        style: str | None = None,
        headers: list[str] | None = None,
        totals_row: bool = False,
        header_row: bool = True,
        first_column: bool = False,
        last_column: bool = False,
        row_stripes: bool = True,
        column_stripes: bool = False,
        auto_filter: bool = True,
    ) -> None: ...
    def insert_rows(self, idx: int, amount: int | None = None) -> None: ...
    def insert_cols(self, idx: int, amount: int | None = None) -> None: ...
    def delete_rows(self, idx: int, amount: int | None = None) -> None: ...
    def delete_cols(self, idx: int, amount: int | None = None) -> None: ...
    def add_chart(
        self,
        chart_type: str,
        series: str | dict[str, str] | list[str | dict[str, str]],
        anchor: str,
        title: str | None = None,
        categories: str | None = None,
        legend: str | None = "r",
    ) -> None: ...
    def add_image(
        self,
        image: str | os.PathLike[str] | bytes,
        anchor: str,
        to: str | None = None,
        width: int | None = None,
        height: int | None = None,
        name: str | None = None,
    ) -> None: ...

class Cell:
    def __init__(self, row: int, column: int) -> None: ...
    @property
    def row(self) -> int: ...
    @property
    def column(self) -> int: ...
    @property
    def coordinate(self) -> str: ...
    @property
    def column_letter(self) -> str: ...
    @property
    def data_type(self) -> str: ...
    @property
    def display_value(self) -> str: ...
    @property
    def rich_text(self) -> list[dict[str, Any]] | None: ...
    @property
    def is_formula(self) -> bool: ...
    value: CellValue
    font: Font | None
    alignment: Alignment | None
    fill: PatternFill | None
    border: Border | None
    protection: Protection | None
    hyperlink: str | None
    comment: str | None
    number_format: str | None
    def offset(self, row: int, column: int) -> Cell: ...

class CellRangeIterator:
    def __iter__(self) -> CellRangeIterator: ...
    def __next__(self) -> tuple[Any, ...]: ...

class AutoFilter:
    ref: str | None

class ColumnDimension:
    width: float | None
    @property
    def index(self) -> str: ...

class ColumnDimensions:
    def __getitem__(self, key: str) -> ColumnDimension: ...

class RowDimension:
    height: float | None
    @property
    def index(self) -> int: ...

class RowDimensions:
    def __getitem__(self, key: int) -> RowDimension: ...

class PivotTable:
    @property
    def name(self) -> str: ...
    @property
    def cache_id(self) -> int | None: ...
    @property
    def location(self) -> str | None: ...
    @property
    def source_sheet(self) -> str | None: ...
    @property
    def source_ref(self) -> str | None: ...
    @property
    def fields(self) -> list[str]: ...
    @property
    def row_fields(self) -> list[str]: ...
    @property
    def col_fields(self) -> list[str]: ...
    @property
    def page_fields(self) -> list[str]: ...
    @property
    def data_fields(self) -> list[dict[str, str]]: ...

class WriteOnlyWorkbook:
    def __init__(self, path: str) -> None: ...
    def create_sheet(self, name: str) -> None: ...
    def append_row(self, values: list[CellValue]) -> None: ...
    def append_rows(self, rows: list[list[CellValue]]) -> None: ...
    def close(self) -> None: ...

class Font:
    name: str | None
    size: float | None
    bold: bool
    italic: bool
    underline: str | None
    strike: bool
    color: _ColorValue
    vertAlign: str | None
    def __init__(
        self,
        name: str | None = None,
        size: float | None = None,
        bold: bool = False,
        italic: bool = False,
        underline: str | None = None,
        strike: bool = False,
        color: _ColorArg = None,
        vertAlign: str | None = None,
    ) -> None: ...
    def copy(self) -> Font: ...

class Alignment:
    horizontal: str | None
    vertical: str | None
    wrap_text: bool
    shrink_to_fit: bool
    indent: int
    text_rotation: int
    def __init__(
        self,
        horizontal: str | None = None,
        vertical: str | None = None,
        wrap_text: bool = False,
        shrink_to_fit: bool = False,
        indent: int = 0,
        text_rotation: int = 0,
    ) -> None: ...
    def copy(self) -> Alignment: ...

class PatternFill:
    fill_type: str | None
    fgColor: _ColorValue
    bgColor: _ColorValue
    patternType: str | None
    @property
    def start_color(self) -> _ColorValue: ...
    @property
    def end_color(self) -> _ColorValue: ...
    def __init__(
        self,
        fill_type: str | None = None,
        fgColor: _ColorArg = None,
        bgColor: _ColorArg = None,
        patternType: str | None = None,
        start_color: _ColorArg = None,
        end_color: _ColorArg = None,
    ) -> None: ...
    def copy(self) -> PatternFill: ...

class Side:
    style: str | None
    color: _ColorValue
    def __init__(self, style: str | None = None, color: _ColorArg = None) -> None: ...
    def copy(self) -> Side: ...

class Border:
    left: Side | None
    right: Side | None
    top: Side | None
    bottom: Side | None
    diagonal: Side | None
    diagonal_direction: str | None
    outline: bool
    def __init__(
        self,
        left: Side | None = None,
        right: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        diagonal: Side | None = None,
        diagonal_direction: str | None = None,
        outline: bool = True,
    ) -> None: ...
    def copy(self) -> Border: ...

class Protection:
    locked: bool
    hidden: bool
    def __init__(self, locked: bool = True, hidden: bool = False) -> None: ...
    def copy(self) -> Protection: ...

class Color:
    rgb: str | None
    theme: int | None
    tint: float
    indexed: int | None
    def __init__(
        self,
        rgb: str | None = None,
        theme: int | None = None,
        tint: float = 0.0,
        indexed: int | None = None,
    ) -> None: ...
    def copy(self) -> Color: ...

class GradientStop:
    position: float
    color: str | None
    def copy(self) -> GradientStop: ...

class GradientFill:
    fill_type: str | None
    degree: float
    left: float
    right: float
    top: float
    bottom: float
    stop: list[GradientStop]
    def copy(self) -> GradientFill: ...

//! Python bindings for Workbook.

#![allow(deprecated)]

use pyo3::prelude::*;
use pyo3::exceptions::{PyTypeError, PyValueError};
use pyo3::types::PyBytes;
use rustypyxl_core::{Workbook, CellValue, CompressionLevel, CellStyle, Font, Fill, Border, BorderStyle, Alignment, Protection};
use std::sync::Arc;

use crate::worksheet::PyWorksheet;
use crate::style::{PyFont, PyAlignment, PyPatternFill, PyBorder, PySide, PyProtection};

/// An Excel Workbook (openpyxl-compatible API).
#[pyclass(name = "Workbook")]
pub struct PyWorkbook {
    pub(crate) inner: Workbook,
}

#[pymethods]
impl PyWorkbook {
    /// Create a new empty workbook.
    #[new]
    fn new() -> Self {
        PyWorkbook {
            inner: Workbook::new(),
        }
    }

    /// Load a workbook from a file path, bytes, or file-like object.
    ///
    /// Args:
    ///     source: File path (str), bytes, or file-like object with .read() method
    ///
    /// Returns:
    ///     Workbook: The loaded workbook
    #[staticmethod]
    #[pyo3(signature = (source))]
    pub fn load(source: &Bound<'_, PyAny>) -> PyResult<Self> {
        // Check if source is a string (file path)
        if let Ok(path) = source.extract::<&str>() {
            let inner = Workbook::load(path)
                .map_err(|e| PyValueError::new_err(e.to_string()))?;
            return Ok(PyWorkbook { inner });
        }

        // Check if source is bytes
        if let Ok(bytes) = source.extract::<&[u8]>() {
            let inner = Workbook::load_from_bytes(bytes)
                .map_err(|e| PyValueError::new_err(e.to_string()))?;
            return Ok(PyWorkbook { inner });
        }

        // Check if source has .read() method (file-like object)
        if source.hasattr("read")? {
            let bytes_obj = source.call_method0("read")?;
            let bytes = bytes_obj.extract::<&[u8]>()?;
            let inner = Workbook::load_from_bytes(bytes)
                .map_err(|e| PyValueError::new_err(e.to_string()))?;
            return Ok(PyWorkbook { inner });
        }

        Err(PyTypeError::new_err(
            "Expected file path (str), bytes, or file-like object with .read() method"
        ))
    }

    /// Get the active worksheet.
    #[getter]
    fn active(self_: Py<Self>, py: Python<'_>) -> PyResult<PyWorksheet> {
        let this = self_.borrow(py);
        if this.inner.worksheets.is_empty() {
            return Err(PyValueError::new_err("No worksheets in workbook"));
        }
        let title = this.inner.sheet_names.get(0)
            .cloned()
            .unwrap_or_else(|| "Sheet1".to_string());
        Ok(PyWorksheet::connected(self_.clone_ref(py), 0, title))
    }

    /// Get all sheet names.
    #[getter]
    fn sheetnames(&self) -> Vec<String> {
        self.inner.sheet_names.clone()
    }

    /// Get all worksheets.
    #[getter]
    fn worksheets(self_: Py<Self>, py: Python<'_>) -> Vec<PyWorksheet> {
        let this = self_.borrow(py);
        (0..this.inner.worksheets.len())
            .map(|i| {
                let title = this.inner.sheet_names.get(i)
                    .cloned()
                    .unwrap_or_else(|| format!("Sheet{}", i + 1));
                PyWorksheet::connected(self_.clone_ref(py), i, title)
            })
            .collect()
    }

    /// Get a worksheet by name using subscript notation: wb['Sheet1'].
    fn __getitem__(self_: Py<Self>, key: &str, py: Python<'_>) -> PyResult<PyWorksheet> {
        let this = self_.borrow(py);
        for (idx, name) in this.inner.sheet_names.iter().enumerate() {
            if name == key {
                return Ok(PyWorksheet::connected(self_.clone_ref(py), idx, name.clone()));
            }
        }
        Err(PyValueError::new_err(format!(
            "Worksheet '{}' does not exist",
            key
        )))
    }

    /// Check if a worksheet exists: 'Sheet1' in wb.
    fn __contains__(&self, key: &str) -> bool {
        self.inner.sheet_names.contains(&key.to_string())
    }

    /// Get the number of worksheets.
    fn __len__(&self) -> usize {
        self.inner.worksheets.len()
    }

    /// Iterate over worksheet names.
    fn __iter__(&self) -> PyResult<PySheetNameIterator> {
        Ok(PySheetNameIterator {
            names: self.inner.sheet_names.clone(),
            index: 0,
        })
    }

    /// Create a new worksheet.
    ///
    /// Args:
    ///     title: Optional worksheet title
    ///     index: Optional position to insert the worksheet
    ///
    /// Returns:
    ///     Worksheet: The newly created worksheet
    #[pyo3(signature = (title=None, index=None))]
    fn create_sheet(self_: Py<Self>, title: Option<String>, index: Option<usize>, py: Python<'_>) -> PyResult<PyWorksheet> {
        // Note: index is currently ignored for simplicity
        let _ = index;

        {
            let mut this = self_.borrow_mut(py);
            this.inner
                .create_sheet(title)
                .map_err(|e| PyValueError::new_err(e.to_string()))?;
        }

        let this = self_.borrow(py);
        let idx = this.inner.worksheets.len() - 1;
        let sheet_title = this.inner.sheet_names.get(idx)
            .cloned()
            .unwrap_or_else(|| format!("Sheet{}", idx + 1));
        Ok(PyWorksheet::connected(self_.clone_ref(py), idx, sheet_title))
    }

    /// Remove a worksheet.
    ///
    /// Args:
    ///     worksheet: The worksheet to remove (by name or PyWorksheet)
    fn remove(&mut self, worksheet: &PyWorksheet) -> PyResult<()> {
        let name = worksheet.title();
        self.inner
            .remove_sheet(&name)
            .map_err(|e| PyValueError::new_err(e.to_string()))
    }

    /// Copy a worksheet.
    ///
    /// Args:
    ///     source: The worksheet to copy
    ///
    /// Returns:
    ///     Worksheet: The copied worksheet
    fn copy_worksheet(self_: Py<Self>, source: &PyWorksheet, py: Python<'_>) -> PyResult<PyWorksheet> {
        let new_name: String;
        let idx: usize;

        {
            let mut this = self_.borrow_mut(py);
            // Get the source worksheet's data
            let source_idx = source.index;
            if source_idx >= this.inner.worksheets.len() {
                return Err(PyValueError::new_err("Invalid worksheet index"));
            }

            // Clone the worksheet
            let src_ws = &this.inner.worksheets[source_idx];
            let mut new_ws = src_ws.clone();

            // Generate a new unique name
            let base_name = format!("{} Copy", src_ws.title);
            let mut counter = 1;
            new_name = base_name.clone();
            let mut temp_name = new_name.clone();
            while this.inner.sheet_names.contains(&temp_name) {
                temp_name = format!("{} {}", base_name, counter);
                counter += 1;
            }
            new_ws.set_title(&temp_name);

            this.inner.worksheets.push(new_ws);
            this.inner.sheet_names.push(temp_name.clone());

            idx = this.inner.worksheets.len() - 1;
            // Re-assign for return
            drop(this);
        }

        let this = self_.borrow(py);
        let sheet_title = this.inner.sheet_names.get(idx)
            .cloned()
            .unwrap_or_else(|| format!("Sheet{}", idx + 1));
        Ok(PyWorksheet::connected(self_.clone_ref(py), idx, sheet_title))
    }

    /// Move a worksheet within the workbook.
    fn move_sheet(&mut self, sheet: &PyWorksheet, offset: i32) -> PyResult<()> {
        let current_idx = sheet.index;
        if current_idx >= self.inner.worksheets.len() {
            return Err(PyValueError::new_err("Invalid worksheet index"));
        }

        let new_idx = (current_idx as i32 + offset).max(0) as usize;
        let new_idx = new_idx.min(self.inner.worksheets.len() - 1);

        if current_idx != new_idx {
            let ws = self.inner.worksheets.remove(current_idx);
            let name = self.inner.sheet_names.remove(current_idx);
            self.inner.worksheets.insert(new_idx, ws);
            self.inner.sheet_names.insert(new_idx, name);
        }

        Ok(())
    }

    /// Get the index of a worksheet.
    fn index(&self, worksheet: &PyWorksheet) -> usize {
        worksheet.index
    }

    /// Create a named range.
    fn create_named_range(&mut self, name: String, worksheet: &PyWorksheet, range: String) -> PyResult<()> {
        let ws_title = worksheet.title();
        let full_range = format!("'{}'!{}", ws_title, range);
        self.inner
            .create_named_range(name, full_range)
            .map_err(|e| PyValueError::new_err(e.to_string()))
    }

    /// Get all defined names (named ranges).
    #[getter]
    fn defined_names(&self) -> Vec<(String, String)> {
        self.inner
            .get_named_ranges()
            .iter()
            .map(|(n, r)| (n.to_string(), r.to_string()))
            .collect()
    }

    /// Save the workbook to a file.
    ///
    /// Args:
    ///     filename: Path to save the Excel file
    fn save(&self, filename: &str) -> PyResult<()> {
        self.inner
            .save(filename)
            .map_err(|e| PyValueError::new_err(e.to_string()))
    }

    /// Save the workbook to bytes.
    ///
    /// Returns:
    ///     bytes: The workbook as an xlsx file in memory
    fn save_to_bytes<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        let bytes = self.inner
            .save_to_bytes()
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(PyBytes::new(py, &bytes))
    }

    /// Set compression level for saving.
    ///
    /// Args:
    ///     level: Compression level - "none", "fast", "default", or "best"
    fn set_compression(&mut self, level: &str) -> PyResult<()> {
        self.inner.compression = match level.to_lowercase().as_str() {
            "none" | "stored" => CompressionLevel::None,
            "fast" | "1" => CompressionLevel::Fast,
            "default" | "6" => CompressionLevel::Default,
            "best" | "9" => CompressionLevel::Best,
            _ => return Err(PyValueError::new_err(
                "Invalid compression level. Use: 'none', 'fast', 'default', or 'best'"
            )),
        };
        Ok(())
    }

    /// Close the workbook (no-op for compatibility).
    fn close(&self) {
        // No-op - we don't hold file handles open
    }

    /// Set a cell value in a specific sheet.
    ///
    /// This is the primary method for setting cell values.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///     value: Value to set (string, number, boolean, or None)
    pub fn set_cell_value(&mut self, sheet_name: &str, row: u32, column: u32, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let cell_value = python_to_cell_value(value)?;
        self.inner
            .set_cell_value_in_sheet(sheet_name, row, column, cell_value)
            .map_err(|e| PyValueError::new_err(e.to_string()))
    }

    /// Get a cell value from a specific sheet.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///
    /// Returns:
    ///     The cell value, or None if empty
    pub fn get_cell_value(&self, sheet_name: &str, row: u32, column: u32, py: Python<'_>) -> PyResult<PyObject> {
        let ws = self.inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        if let Some(cell) = ws.get_cell(row, column) {
            Ok(cell_value_to_python(&cell.value, py))
        } else {
            Ok(py.None())
        }
    }

    /// Write multiple rows of data to a sheet (bulk operation for performance).
    ///
    /// This is significantly faster than setting cells one at a time.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     data: List of rows, where each row is a list of values
    ///     start_row: Starting row (1-indexed, default 1)
    ///     start_col: Starting column (1-indexed, default 1)
    #[pyo3(signature = (sheet_name, data, start_row=1, start_col=1))]
    fn write_rows(
        &mut self,
        sheet_name: &str,
        data: Vec<Vec<Bound<'_, PyAny>>>,
        start_row: u32,
        start_col: u32,
    ) -> PyResult<()> {
        // Get mutable reference to worksheet once (avoid repeated lookups)
        let ws = self.inner
            .get_sheet_by_name_mut(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        for (row_idx, row_data) in data.iter().enumerate() {
            let row = start_row + row_idx as u32;
            for (col_idx, value) in row_data.iter().enumerate() {
                let col = start_col + col_idx as u32;
                let cell_value = python_to_cell_value(value)?;
                ws.set_cell_value(row, col, cell_value);
            }
        }
        Ok(())
    }

    /// Read all values from a sheet as a 2D list (bulk operation for performance).
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     min_row: Minimum row (1-indexed, default 1)
    ///     max_row: Maximum row (default: last row with data)
    ///     min_col: Minimum column (1-indexed, default 1)
    ///     max_col: Maximum column (default: last column with data)
    ///
    /// Returns:
    ///     List of rows, where each row is a list of values
    #[pyo3(signature = (sheet_name, min_row=None, max_row=None, min_col=None, max_col=None))]
    fn read_rows(
        &self,
        sheet_name: &str,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        py: Python<'_>,
    ) -> PyResult<Vec<Vec<PyObject>>> {
        let ws = self.inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        // dimensions() returns (min_row, min_col, max_row, max_col)
        let (dims_min_row, dims_min_col, dims_max_row, dims_max_col) = ws.dimensions();

        let min_r = min_row.unwrap_or(dims_min_row);
        let max_r = max_row.unwrap_or(dims_max_row);
        let min_c = min_col.unwrap_or(dims_min_col);
        let max_c = max_col.unwrap_or(dims_max_col);

        let mut result = Vec::new();
        for row in min_r..=max_r {
            let mut row_data = Vec::new();
            for col in min_c..=max_c {
                if let Some(cell) = ws.get_cell(row, col) {
                    row_data.push(cell_value_to_python(&cell.value, py));
                } else {
                    row_data.push(py.None());
                }
            }
            result.push(row_data);
        }
        Ok(result)
    }

    /// Set a cell's font style.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///     font: Font style to apply
    pub fn set_cell_font(&mut self, sheet_name: &str, row: u32, column: u32, font: &PyFont) -> PyResult<()> {
        let rust_font = pyfont_to_font(font);
        let style = rustypyxl_core::CellStyle::new().with_font(rust_font);
        self.set_or_merge_cell_style(sheet_name, row, column, style)
    }

    /// Set a cell's fill (background color).
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///     fill: PatternFill style to apply
    pub fn set_cell_fill(&mut self, sheet_name: &str, row: u32, column: u32, fill: &PyPatternFill) -> PyResult<()> {
        let rust_fill = pyfill_to_fill(fill);
        let style = rustypyxl_core::CellStyle::new().with_fill(rust_fill);
        self.set_or_merge_cell_style(sheet_name, row, column, style)
    }

    /// Set a cell's border.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///     border: Border style to apply
    pub fn set_cell_border(&mut self, sheet_name: &str, row: u32, column: u32, border: &PyBorder) -> PyResult<()> {
        let rust_border = pyborder_to_border(border);
        let style = rustypyxl_core::CellStyle::new().with_border(rust_border);
        self.set_or_merge_cell_style(sheet_name, row, column, style)
    }

    /// Set a cell's alignment.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///     alignment: Alignment style to apply
    pub fn set_cell_alignment(&mut self, sheet_name: &str, row: u32, column: u32, alignment: &PyAlignment) -> PyResult<()> {
        let rust_align = pyalignment_to_alignment(alignment);
        let style = rustypyxl_core::CellStyle::new().with_alignment(rust_align);
        self.set_or_merge_cell_style(sheet_name, row, column, style)
    }

    /// Set a cell's number format.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///     format: Number format string (e.g., "#,##0.00", "0.00%")
    pub fn set_cell_number_format(&mut self, sheet_name: &str, row: u32, column: u32, format: &str) -> PyResult<()> {
        let style = rustypyxl_core::CellStyle::new().with_number_format(format);
        self.set_or_merge_cell_style(sheet_name, row, column, style)
    }

    /// Set a cell's protection.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///     protection: Protection style to apply
    pub fn set_cell_protection(&mut self, sheet_name: &str, row: u32, column: u32, protection: &PyProtection) -> PyResult<()> {
        let rust_protection = pyprotection_to_protection(protection);
        let style = rustypyxl_core::CellStyle::new().with_protection(rust_protection);
        self.set_or_merge_cell_style(sheet_name, row, column, style)
    }

    /// Set multiple style properties on a cell at once.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///     font: Optional font style
    ///     fill: Optional fill style
    ///     border: Optional border style
    ///     alignment: Optional alignment style
    ///     number_format: Optional number format string
    #[pyo3(signature = (sheet_name, row, column, font=None, fill=None, border=None, alignment=None, number_format=None))]
    fn set_cell_style(
        &mut self,
        sheet_name: &str,
        row: u32,
        column: u32,
        font: Option<&PyFont>,
        fill: Option<&PyPatternFill>,
        border: Option<&PyBorder>,
        alignment: Option<&PyAlignment>,
        number_format: Option<&str>,
    ) -> PyResult<()> {
        let mut style = rustypyxl_core::CellStyle::new();

        if let Some(f) = font {
            style = style.with_font(pyfont_to_font(f));
        }
        if let Some(f) = fill {
            style = style.with_fill(pyfill_to_fill(f));
        }
        if let Some(b) = border {
            style = style.with_border(pyborder_to_border(b));
        }
        if let Some(a) = alignment {
            style = style.with_alignment(pyalignment_to_alignment(a));
        }
        if let Some(nf) = number_format {
            style = style.with_number_format(nf);
        }

        self.set_or_merge_cell_style(sheet_name, row, column, style)
    }

    /// Get a cell's font style.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///
    /// Returns:
    ///     Font style or None if not set
    pub fn get_cell_font(&self, sheet_name: &str, row: u32, column: u32) -> PyResult<Option<PyFont>> {
        let ws = self.inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        if let Some(cell) = ws.get_cell(row, column) {
            if let Some(ref style) = cell.style {
                if let Some(ref font) = style.font {
                    return Ok(Some(font_to_pyfont(font)));
                }
            }
        }
        Ok(None)
    }

    /// Get a cell's fill style.
    pub fn get_cell_fill(&self, sheet_name: &str, row: u32, column: u32) -> PyResult<Option<PyPatternFill>> {
        let ws = self.inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        if let Some(cell) = ws.get_cell(row, column) {
            if let Some(ref style) = cell.style {
                if let Some(ref fill) = style.fill {
                    return Ok(Some(fill_to_pyfill(fill)));
                }
            }
        }
        Ok(None)
    }

    /// Get a cell's border style.
    pub fn get_cell_border(&self, sheet_name: &str, row: u32, column: u32) -> PyResult<Option<PyBorder>> {
        let ws = self.inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        if let Some(cell) = ws.get_cell(row, column) {
            if let Some(ref style) = cell.style {
                if let Some(ref border) = style.border {
                    return Ok(Some(border_to_pyborder(border)));
                }
            }
        }
        Ok(None)
    }

    /// Get a cell's alignment style.
    pub fn get_cell_alignment(&self, sheet_name: &str, row: u32, column: u32) -> PyResult<Option<PyAlignment>> {
        let ws = self.inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        if let Some(cell) = ws.get_cell(row, column) {
            if let Some(ref style) = cell.style {
                if let Some(ref align) = style.alignment {
                    return Ok(Some(alignment_to_pyalignment(align)));
                }
            }
        }
        Ok(None)
    }

    /// Get a cell's number format.
    pub fn get_cell_number_format(&self, sheet_name: &str, row: u32, column: u32) -> PyResult<Option<String>> {
        let ws = self.inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        if let Some(cell) = ws.get_cell(row, column) {
            if let Some(ref style) = cell.style {
                return Ok(style.number_format.clone());
            }
        }
        Ok(None)
    }

    /// Get a cell's protection style.
    pub fn get_cell_protection(&self, sheet_name: &str, row: u32, column: u32) -> PyResult<Option<PyProtection>> {
        let ws = self.inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        if let Some(cell) = ws.get_cell(row, column) {
            if let Some(ref style) = cell.style {
                if let Some(ref protection) = style.protection {
                    return Ok(Some(protection_to_pyprotection(protection)));
                }
            }
        }
        Ok(None)
    }

    /// Import data from a Parquet file directly into a worksheet.
    ///
    /// This is the fastest way to load large datasets, as it bypasses
    /// Python FFI entirely and reads directly from Parquet into cells.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet to insert into
    ///     path: Path to the Parquet file
    ///     start_row: Starting row (1-indexed, default 1)
    ///     start_col: Starting column (1-indexed, default 1)
    ///     include_headers: Include column headers (default True)
    ///     column_renames: Dict mapping original column names to new names
    ///     columns: List of column names to import (None = all columns)
    ///
    /// Returns:
    ///     Dict with import results: rows_imported, columns_imported,
    ///     range (e.g. "A1:Z1000"), header_range, data_range, column_names
    #[cfg(feature = "parquet")]
    #[pyo3(signature = (sheet_name, path, start_row=1, start_col=1, include_headers=true, column_renames=None, columns=None))]
    fn insert_from_parquet(
        &mut self,
        sheet_name: &str,
        path: &str,
        start_row: u32,
        start_col: u32,
        include_headers: bool,
        column_renames: Option<std::collections::HashMap<String, String>>,
        columns: Option<Vec<String>>,
        py: Python<'_>,
    ) -> PyResult<PyObject> {
        use rustypyxl_core::ParquetImportOptions;
        use pyo3::types::PyDict;

        let mut opts = ParquetImportOptions::new().with_headers(include_headers);

        if let Some(renames) = column_renames {
            opts.column_renames = renames;
        }

        if let Some(cols) = columns {
            opts.columns = cols;
        }

        let result = self.inner
            .insert_from_parquet(sheet_name, path, start_row, start_col, Some(opts))
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        // Build result dict
        let dict = PyDict::new(py);
        dict.set_item("rows_imported", result.rows_imported)?;
        dict.set_item("columns_imported", result.columns_imported)?;
        dict.set_item("start_row", result.start_row)?;
        dict.set_item("start_col", result.start_col)?;
        dict.set_item("end_row", result.end_row)?;
        dict.set_item("end_col", result.end_col)?;
        dict.set_item("range", result.range_with_headers())?;
        dict.set_item("header_range", result.header_range())?;
        dict.set_item("data_range", result.data_range())?;
        dict.set_item("column_names", result.column_names)?;

        Ok(dict.into())
    }

    /// Export a worksheet to a Parquet file.
    ///
    /// This exports cell data directly to Parquet format with automatic type inference.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet to export
    ///     path: Output path for the Parquet file
    ///     has_headers: Whether the first row contains headers (default True)
    ///     compression: Compression type: "snappy", "gzip", "zstd", "lz4", "none" (default "snappy")
    ///     column_renames: Dict mapping original column names to new names
    ///     column_types: Dict mapping column names to types: "string", "float64", "int64", "boolean", "date", "datetime"
    ///
    /// Returns:
    ///     Dict with export results: rows_exported, columns_exported, column_names, file_size
    #[cfg(feature = "parquet")]
    #[pyo3(signature = (sheet_name, path, has_headers=true, compression="snappy", column_renames=None, column_types=None))]
    fn export_to_parquet(
        &self,
        sheet_name: &str,
        path: &str,
        has_headers: bool,
        compression: &str,
        column_renames: Option<std::collections::HashMap<String, String>>,
        column_types: Option<std::collections::HashMap<String, String>>,
        py: Python<'_>,
    ) -> PyResult<PyObject> {
        use rustypyxl_core::{ParquetExportOptions, ParquetCompression, ColumnType};
        use pyo3::types::PyDict;

        let compression = match compression.to_lowercase().as_str() {
            "none" => ParquetCompression::None,
            "snappy" => ParquetCompression::Snappy,
            "gzip" => ParquetCompression::Gzip,
            "zstd" => ParquetCompression::Zstd,
            "lz4" => ParquetCompression::Lz4,
            _ => return Err(PyValueError::new_err(format!(
                "Invalid compression: {}. Use 'none', 'snappy', 'gzip', 'zstd', or 'lz4'",
                compression
            ))),
        };

        let mut opts = ParquetExportOptions::new()
            .with_headers(has_headers)
            .with_compression(compression);

        if let Some(renames) = column_renames {
            opts.column_renames = renames;
        }

        if let Some(types) = column_types {
            for (col_name, type_str) in types {
                let col_type = match type_str.to_lowercase().as_str() {
                    "string" | "str" => ColumnType::String,
                    "float64" | "float" | "double" => ColumnType::Float64,
                    "int64" | "int" | "integer" => ColumnType::Int64,
                    "boolean" | "bool" => ColumnType::Boolean,
                    "date" => ColumnType::Date,
                    "datetime" | "timestamp" => ColumnType::DateTime,
                    "auto" => ColumnType::Auto,
                    _ => return Err(PyValueError::new_err(format!(
                        "Invalid column type: {}. Use 'string', 'float64', 'int64', 'boolean', 'date', 'datetime', or 'auto'",
                        type_str
                    ))),
                };
                opts.column_types.insert(col_name, col_type);
            }
        }

        let result = self.inner
            .export_to_parquet(sheet_name, path, Some(opts))
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        // Build result dict
        let dict = PyDict::new(py);
        dict.set_item("rows_exported", result.rows_exported)?;
        dict.set_item("columns_exported", result.columns_exported)?;
        dict.set_item("column_names", result.column_names)?;
        dict.set_item("file_size", result.file_size)?;

        Ok(dict.into())
    }

    /// Export a specific range from a worksheet to a Parquet file.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet to export
    ///     path: Output path for the Parquet file
    ///     min_row: Starting row (1-indexed)
    ///     min_col: Starting column (1-indexed)
    ///     max_row: Ending row (1-indexed)
    ///     max_col: Ending column (1-indexed)
    ///     has_headers: Whether the first row contains headers (default True)
    ///     compression: Compression type (default "snappy")
    ///
    /// Returns:
    ///     Dict with export results
    #[cfg(feature = "parquet")]
    #[pyo3(signature = (sheet_name, path, min_row, min_col, max_row, max_col, has_headers=true, compression="snappy"))]
    fn export_range_to_parquet(
        &self,
        sheet_name: &str,
        path: &str,
        min_row: u32,
        min_col: u32,
        max_row: u32,
        max_col: u32,
        has_headers: bool,
        compression: &str,
        py: Python<'_>,
    ) -> PyResult<PyObject> {
        use rustypyxl_core::{ParquetExportOptions, ParquetCompression};
        use pyo3::types::PyDict;

        let compression = match compression.to_lowercase().as_str() {
            "none" => ParquetCompression::None,
            "snappy" => ParquetCompression::Snappy,
            "gzip" => ParquetCompression::Gzip,
            "zstd" => ParquetCompression::Zstd,
            "lz4" => ParquetCompression::Lz4,
            _ => return Err(PyValueError::new_err(format!(
                "Invalid compression: {}",
                compression
            ))),
        };

        let opts = ParquetExportOptions::new()
            .with_headers(has_headers)
            .with_compression(compression);

        let result = self.inner
            .export_range_to_parquet(sheet_name, path, min_row, min_col, max_row, max_col, Some(opts))
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        let dict = PyDict::new(py);
        dict.set_item("rows_exported", result.rows_exported)?;
        dict.set_item("columns_exported", result.columns_exported)?;
        dict.set_item("column_names", result.column_names)?;
        dict.set_item("file_size", result.file_size)?;

        Ok(dict.into())
    }

    /// Load a workbook from S3.
    ///
    /// Args:
    ///     bucket: S3 bucket name
    ///     key: S3 object key (path within the bucket)
    ///     region: Optional AWS region (e.g., "us-east-1")
    ///     endpoint_url: Optional custom endpoint URL (for S3-compatible services)
    ///
    /// Returns:
    ///     Workbook: The loaded workbook
    #[cfg(feature = "s3")]
    #[staticmethod]
    #[pyo3(signature = (bucket, key, region=None, endpoint_url=None))]
    pub fn load_from_s3(
        bucket: &str,
        key: &str,
        region: Option<&str>,
        endpoint_url: Option<&str>,
    ) -> PyResult<Self> {
        use rustypyxl_core::S3Config;

        let mut config = S3Config::new();
        if let Some(r) = region {
            config = config.with_region(r);
        }
        if let Some(url) = endpoint_url {
            config = config.with_endpoint_url(url);
        }

        let inner = Workbook::load_from_s3(bucket, key, Some(config))
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(PyWorkbook { inner })
    }

    /// Save the workbook to S3.
    ///
    /// Args:
    ///     bucket: S3 bucket name
    ///     key: S3 object key (path within the bucket)
    ///     region: Optional AWS region (e.g., "us-east-1")
    ///     endpoint_url: Optional custom endpoint URL (for S3-compatible services)
    #[cfg(feature = "s3")]
    #[pyo3(signature = (bucket, key, region=None, endpoint_url=None))]
    pub fn save_to_s3(
        &self,
        bucket: &str,
        key: &str,
        region: Option<&str>,
        endpoint_url: Option<&str>,
    ) -> PyResult<()> {
        use rustypyxl_core::S3Config;

        let mut config = S3Config::new();
        if let Some(r) = region {
            config = config.with_region(r);
        }
        if let Some(url) = endpoint_url {
            config = config.with_endpoint_url(url);
        }

        self.inner
            .save_to_s3(bucket, key, Some(config))
            .map_err(|e| PyValueError::new_err(e.to_string()))
    }

    fn __str__(&self) -> String {
        format!("<Workbook with {} sheet(s)>", self.inner.worksheets.len())
    }

    fn __repr__(&self) -> String {
        self.__str__()
    }
}

impl PyWorkbook {
    /// Helper to set or merge a cell style with the existing style.
    fn set_or_merge_cell_style(&mut self, sheet_name: &str, row: u32, column: u32, new_style: CellStyle) -> PyResult<()> {
        // First, compute merged style from existing cell (if any)
        let merged_style = {
            let ws = self.inner
                .get_sheet_by_name(sheet_name)
                .map_err(|e| PyValueError::new_err(e.to_string()))?;

            if let Some(cell) = ws.get_cell(row, column) {
                if let Some(ref existing) = cell.style {
                    let mut merged = (**existing).clone();
                    if new_style.font.is_some() {
                        merged.font = new_style.font.clone();
                    }
                    if new_style.fill.is_some() {
                        merged.fill = new_style.fill.clone();
                    }
                    if new_style.border.is_some() {
                        merged.border = new_style.border.clone();
                    }
                    if new_style.alignment.is_some() {
                        merged.alignment = new_style.alignment.clone();
                    }
                    if new_style.number_format.is_some() {
                        merged.number_format = new_style.number_format.clone();
                    }
                    merged
                } else {
                    new_style
                }
            } else {
                new_style
            }
        };

        // Get the style index for this style
        let style_index = self.inner.styles.get_or_add_cell_xf(&merged_style);

        // Now get mutable reference and set the style
        let ws = self.inner
            .get_sheet_by_name_mut(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        let cell = ws.get_or_create_cell_mut(row, column);
        cell.style = Some(Arc::new(merged_style));
        cell.style_index = Some(style_index as u32);

        Ok(())
    }
}

/// Iterator over worksheet names.
#[pyclass]
pub struct PySheetNameIterator {
    names: Vec<String>,
    index: usize,
}

#[pymethods]
impl PySheetNameIterator {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__(&mut self) -> Option<String> {
        if self.index < self.names.len() {
            let name = self.names[self.index].clone();
            self.index += 1;
            Some(name)
        } else {
            None
        }
    }
}

/// Convert a Python value to a CellValue.
fn python_to_cell_value(value: &Bound<'_, PyAny>) -> PyResult<CellValue> {
    if value.is_none() {
        Ok(CellValue::Empty)
    } else if let Ok(s) = value.extract::<String>() {
        if s.starts_with('=') {
            // Store formula WITHOUT the leading '=' (it will be added back when written)
            Ok(CellValue::Formula(s[1..].to_string()))
        } else {
            Ok(CellValue::from(s))
        }
    } else if let Ok(b) = value.extract::<bool>() {
        Ok(CellValue::Boolean(b))
    } else if let Ok(n) = value.extract::<f64>() {
        Ok(CellValue::Number(n))
    } else if let Ok(n) = value.extract::<i64>() {
        Ok(CellValue::Number(n as f64))
    } else {
        // Try to convert to string as fallback
        Ok(CellValue::from(value.str()?.to_string()))
    }
}

/// Convert a CellValue to a Python object.
fn cell_value_to_python(value: &CellValue, py: Python<'_>) -> PyObject {
    match value {
        CellValue::Empty => py.None(),
        CellValue::String(s) => s.as_ref().to_object(py),
        CellValue::Number(n) => n.to_object(py),
        CellValue::Boolean(b) => b.to_object(py),
        CellValue::Formula(f) => format!("={}", f).to_object(py),
        CellValue::Date(d) => d.to_object(py),
    }
}

// =====================
// Style conversion helpers
// =====================

/// Convert PyFont to Rust Font.
fn pyfont_to_font(pf: &PyFont) -> Font {
    Font {
        name: pf.name.clone(),
        size: pf.size,
        bold: pf.bold,
        italic: pf.italic,
        underline: pf.underline.is_some(),
        strike: pf.strike,
        color: pf.color.clone(),
        vert_align: pf.vertAlign.clone(),
    }
}

/// Convert Rust Font to PyFont.
fn font_to_pyfont(f: &Font) -> PyFont {
    PyFont {
        name: f.name.clone(),
        size: f.size,
        bold: f.bold,
        italic: f.italic,
        underline: if f.underline { Some("single".to_string()) } else { None },
        strike: f.strike,
        color: f.color.clone(),
        vertAlign: f.vert_align.clone(),
    }
}

/// Convert PyPatternFill to Rust Fill.
fn pyfill_to_fill(pf: &PyPatternFill) -> Fill {
    Fill {
        pattern_type: pf.fill_type.clone().or(pf.patternType.clone()),
        fg_color: pf.fgColor.clone(),
        bg_color: pf.bgColor.clone(),
    }
}

/// Convert Rust Fill to PyPatternFill.
fn fill_to_pyfill(f: &Fill) -> PyPatternFill {
    PyPatternFill {
        fill_type: f.pattern_type.clone(),
        fgColor: f.fg_color.clone(),
        bgColor: f.bg_color.clone(),
        patternType: f.pattern_type.clone(),
    }
}

/// Convert PySide to Rust BorderStyle.
fn pyside_to_borderstyle(ps: &PySide) -> Option<BorderStyle> {
    ps.style.as_ref().map(|s| BorderStyle {
        style: s.clone(),
        color: ps.color.clone(),
    })
}

/// Convert Rust BorderStyle to PySide.
fn borderstyle_to_pyside(bs: &BorderStyle) -> PySide {
    PySide {
        style: Some(bs.style.clone()),
        color: bs.color.clone(),
    }
}

/// Convert PyBorder to Rust Border.
fn pyborder_to_border(pb: &PyBorder) -> Border {
    Border {
        left: pb.left.as_ref().and_then(|s| pyside_to_borderstyle(s)),
        right: pb.right.as_ref().and_then(|s| pyside_to_borderstyle(s)),
        top: pb.top.as_ref().and_then(|s| pyside_to_borderstyle(s)),
        bottom: pb.bottom.as_ref().and_then(|s| pyside_to_borderstyle(s)),
        diagonal: pb.diagonal.as_ref().and_then(|s| pyside_to_borderstyle(s)),
    }
}

/// Convert Rust Border to PyBorder.
fn border_to_pyborder(b: &Border) -> PyBorder {
    PyBorder {
        left: b.left.as_ref().map(borderstyle_to_pyside),
        right: b.right.as_ref().map(borderstyle_to_pyside),
        top: b.top.as_ref().map(borderstyle_to_pyside),
        bottom: b.bottom.as_ref().map(borderstyle_to_pyside),
        diagonal: b.diagonal.as_ref().map(borderstyle_to_pyside),
        diagonal_direction: None,
        outline: true,
    }
}

/// Convert PyAlignment to Rust Alignment.
fn pyalignment_to_alignment(pa: &PyAlignment) -> Alignment {
    Alignment {
        horizontal: pa.horizontal.clone(),
        vertical: pa.vertical.clone(),
        wrap_text: pa.wrap_text,
        text_rotation: if pa.text_rotation != 0 { Some(pa.text_rotation) } else { None },
        indent: if pa.indent != 0 { Some(pa.indent) } else { None },
        shrink_to_fit: pa.shrink_to_fit,
    }
}

/// Convert Rust Alignment to PyAlignment.
fn alignment_to_pyalignment(a: &Alignment) -> PyAlignment {
    PyAlignment {
        horizontal: a.horizontal.clone(),
        vertical: a.vertical.clone(),
        wrap_text: a.wrap_text,
        shrink_to_fit: a.shrink_to_fit,
        indent: a.indent.unwrap_or(0),
        text_rotation: a.text_rotation.unwrap_or(0),
    }
}

/// Convert PyProtection to Rust Protection.
fn pyprotection_to_protection(pp: &PyProtection) -> Protection {
    Protection {
        locked: pp.locked,
        hidden: pp.hidden,
    }
}

/// Convert Rust Protection to PyProtection.
fn protection_to_pyprotection(p: &Protection) -> PyProtection {
    PyProtection {
        locked: p.locked,
        hidden: p.hidden,
    }
}

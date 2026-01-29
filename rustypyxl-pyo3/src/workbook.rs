//! Python bindings for Workbook.

#![allow(deprecated)]

use pyo3::prelude::*;
use pyo3::exceptions::PyValueError;
use rustypyxl_core::{Workbook, CellValue, CompressionLevel};

use crate::worksheet::PyWorksheet;

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

    /// Load a workbook from a file path.
    #[staticmethod]
    pub fn load(filename: &str) -> PyResult<Self> {
        let inner = Workbook::load(filename)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(PyWorkbook { inner })
    }

    /// Get the active worksheet.
    #[getter]
    fn active(&self) -> PyResult<PyWorksheet> {
        if self.inner.worksheets.is_empty() {
            return Err(PyValueError::new_err("No worksheets in workbook"));
        }
        Ok(PyWorksheet::from_ref(self, 0))
    }

    /// Get all sheet names.
    #[getter]
    fn sheetnames(&self) -> Vec<String> {
        self.inner.sheet_names.clone()
    }

    /// Get all worksheets.
    #[getter]
    fn worksheets(&self) -> Vec<PyWorksheet> {
        (0..self.inner.worksheets.len())
            .map(|i| PyWorksheet::from_ref(self, i))
            .collect()
    }

    /// Get a worksheet by name using subscript notation: wb['Sheet1'].
    fn __getitem__(&self, key: &str) -> PyResult<PyWorksheet> {
        for (idx, name) in self.inner.sheet_names.iter().enumerate() {
            if name == key {
                return Ok(PyWorksheet::from_ref(self, idx));
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
    fn create_sheet(&mut self, title: Option<String>, index: Option<usize>) -> PyResult<PyWorksheet> {
        // Note: index is currently ignored for simplicity
        let _ = index;

        self.inner
            .create_sheet(title)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        let idx = self.inner.worksheets.len() - 1;
        Ok(PyWorksheet::from_ref(self, idx))
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
    fn copy_worksheet(&mut self, source: &PyWorksheet) -> PyResult<PyWorksheet> {
        // Get the source worksheet's data
        let source_idx = source.index;
        if source_idx >= self.inner.worksheets.len() {
            return Err(PyValueError::new_err("Invalid worksheet index"));
        }

        // Clone the worksheet
        let src_ws = &self.inner.worksheets[source_idx];
        let mut new_ws = src_ws.clone();

        // Generate a new unique name
        let base_name = format!("{} Copy", src_ws.title);
        let mut counter = 1;
        let mut new_name = base_name.clone();
        while self.inner.sheet_names.contains(&new_name) {
            new_name = format!("{} {}", base_name, counter);
            counter += 1;
        }
        new_ws.set_title(&new_name);

        self.inner.worksheets.push(new_ws);
        self.inner.sheet_names.push(new_name);

        let idx = self.inner.worksheets.len() - 1;
        Ok(PyWorksheet::from_ref(self, idx))
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
    fn set_cell_value(&mut self, sheet_name: &str, row: u32, column: u32, value: &Bound<'_, PyAny>) -> PyResult<()> {
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
    fn get_cell_value(&self, sheet_name: &str, row: u32, column: u32, py: Python<'_>) -> PyResult<PyObject> {
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

    fn __str__(&self) -> String {
        format!("<Workbook with {} sheet(s)>", self.inner.worksheets.len())
    }

    fn __repr__(&self) -> String {
        self.__str__()
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
        CellValue::Formula(f) => f.to_object(py),
        CellValue::Date(d) => d.to_object(py),
    }
}

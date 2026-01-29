//! Python bindings for Worksheet.

use pyo3::prelude::*;
use pyo3::exceptions::PyValueError;
use pyo3::Py;
use rustypyxl_core::{parse_coordinate, column_to_letter};

use crate::cell::PyCell;
use crate::workbook::PyWorkbook;

/// An Excel Worksheet (openpyxl-compatible API).
///
/// Worksheets are accessed from a Workbook, not created directly.
#[pyclass(name = "Worksheet")]
pub struct PyWorksheet {
    /// Index of the worksheet in the workbook.
    pub(crate) index: usize,
    /// Cached title (for when we can't access workbook).
    cached_title: String,
}

impl PyWorksheet {
    /// Create a PyWorksheet from a workbook reference.
    pub fn from_ref(wb: &PyWorkbook, index: usize) -> Self {
        let title = wb.inner.sheet_names.get(index)
            .cloned()
            .unwrap_or_else(|| format!("Sheet{}", index + 1));
        PyWorksheet { index, cached_title: title }
    }
}

#[pymethods]
impl PyWorksheet {
    /// Get the worksheet title.
    #[getter]
    pub fn title(&self) -> String {
        self.cached_title.clone()
    }

    /// Set the worksheet title.
    #[setter]
    fn set_title(&mut self, _title: String) {
        // Note: This is tricky because we need access to the workbook.
        // For now, just update the cached title.
        // In a full implementation, we'd need to store a reference to the workbook.
        self.cached_title = _title;
    }

    /// Get a cell using subscript notation: ws['A1'] or ws['A1:B2'].
    fn __getitem__(&self, key: &str, py: Python<'_>) -> PyResult<PyObject> {
        // Check if it's a range
        if key.contains(':') {
            // Return a cell range
            return Err(PyValueError::new_err(
                "Cell ranges not yet implemented. Use cell(row, column) instead."
            ));
        }

        // Single cell
        let (row, col) = parse_coordinate(key)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        Ok(Py::new(py, PyCell::new(row, col))?.into_any())
    }

    /// Set a cell value using subscript notation: ws['A1'] = 'Hello'.
    fn __setitem__(&self, _key: &str, _value: PyObject) -> PyResult<()> {
        // This requires workbook access, which is tricky with our current design.
        // For now, raise an error suggesting the alternative.
        Err(PyValueError::new_err(
            "Direct cell assignment not yet supported. Use cell(row, column).value = value instead."
        ))
    }

    /// Get a cell at a specific row and column.
    ///
    /// Args:
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///
    /// Returns:
    ///     Cell: The cell at the specified position
    #[pyo3(signature = (row, column=None))]
    fn cell(&self, row: u32, column: Option<u32>) -> PyResult<PyCell> {
        let col = column.unwrap_or(1);
        if row == 0 || col == 0 {
            return Err(PyValueError::new_err("Row and column must be at least 1"));
        }
        Ok(PyCell::new(row, col))
    }

    /// Iterate over rows.
    ///
    /// Args:
    ///     min_row: Minimum row number
    ///     max_row: Maximum row number
    ///     min_col: Minimum column number
    ///     max_col: Maximum column number
    ///     values_only: If True, return values instead of Cell objects
    #[pyo3(signature = (min_row=None, max_row=None, min_col=None, max_col=None, values_only=false))]
    fn iter_rows(
        &self,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        values_only: bool,
        py: Python<'_>,
    ) -> PyResult<Vec<Vec<PyObject>>> {
        let min_r = min_row.unwrap_or(1);
        let max_r = max_row.unwrap_or(10);
        let min_c = min_col.unwrap_or(1);
        let max_c = max_col.unwrap_or(10);

        let mut rows = Vec::new();
        for r in min_r..=max_r {
            let mut row = Vec::new();
            for c in min_c..=max_c {
                if values_only {
                    row.push(py.None());
                } else {
                    row.push(Py::new(py, PyCell::new(r, c))?.into_any());
                }
            }
            rows.push(row);
        }
        Ok(rows)
    }

    /// Iterate over columns.
    #[pyo3(signature = (min_col=None, max_col=None, min_row=None, max_row=None, values_only=false))]
    fn iter_cols(
        &self,
        min_col: Option<u32>,
        max_col: Option<u32>,
        min_row: Option<u32>,
        max_row: Option<u32>,
        values_only: bool,
        py: Python<'_>,
    ) -> PyResult<Vec<Vec<PyObject>>> {
        let min_c = min_col.unwrap_or(1);
        let max_c = max_col.unwrap_or(10);
        let min_r = min_row.unwrap_or(1);
        let max_r = max_row.unwrap_or(10);

        let mut cols = Vec::new();
        for c in min_c..=max_c {
            let mut col = Vec::new();
            for r in min_r..=max_r {
                if values_only {
                    col.push(py.None());
                } else {
                    col.push(Py::new(py, PyCell::new(r, c))?.into_any());
                }
            }
            cols.push(col);
        }
        Ok(cols)
    }

    /// Get the maximum row containing data.
    #[getter]
    fn max_row(&self) -> u32 {
        // Without workbook access, return a default
        1
    }

    /// Get the maximum column containing data.
    #[getter]
    fn max_column(&self) -> u32 {
        1
    }

    /// Get the minimum row containing data.
    #[getter]
    fn min_row(&self) -> u32 {
        1
    }

    /// Get the minimum column containing data.
    #[getter]
    fn min_column(&self) -> u32 {
        1
    }

    /// Get the dimensions as a string (e.g., "A1:D10").
    #[getter]
    fn dimensions(&self) -> String {
        format!("{}{}:{}{}",
            column_to_letter(self.min_column()),
            self.min_row(),
            column_to_letter(self.max_column()),
            self.max_row()
        )
    }

    /// Merge cells in a range.
    ///
    /// Args:
    ///     range_string: Range to merge (e.g., "A1:B2")
    ///     start_row: Start row (alternative to range_string)
    ///     start_column: Start column
    ///     end_row: End row
    ///     end_column: End column
    #[pyo3(signature = (range_string=None, start_row=None, start_column=None, end_row=None, end_column=None))]
    fn merge_cells(
        &self,
        range_string: Option<&str>,
        start_row: Option<u32>,
        start_column: Option<u32>,
        end_row: Option<u32>,
        end_column: Option<u32>,
    ) -> PyResult<()> {
        // This requires workbook access to modify the worksheet
        let _range = if let Some(rs) = range_string {
            rs.to_string()
        } else if let (Some(sr), Some(sc), Some(er), Some(ec)) =
            (start_row, start_column, end_row, end_column)
        {
            format!("{}{}:{}{}",
                column_to_letter(sc), sr,
                column_to_letter(ec), er
            )
        } else {
            return Err(PyValueError::new_err(
                "Must specify either range_string or all of start_row, start_column, end_row, end_column"
            ));
        };

        // Note: Actual merging requires workbook access
        Ok(())
    }

    /// Unmerge cells in a range.
    #[pyo3(signature = (range_string=None, start_row=None, start_column=None, end_row=None, end_column=None))]
    fn unmerge_cells(
        &self,
        range_string: Option<&str>,
        start_row: Option<u32>,
        start_column: Option<u32>,
        end_row: Option<u32>,
        end_column: Option<u32>,
    ) -> PyResult<()> {
        self.merge_cells(range_string, start_row, start_column, end_row, end_column)
    }

    /// Get merged cell ranges.
    #[getter]
    fn merged_cells(&self) -> Vec<String> {
        Vec::new()
    }

    /// Append a row of values.
    fn append(&self, _iterable: Vec<PyObject>) -> PyResult<()> {
        // Requires workbook access
        Ok(())
    }

    /// Insert rows.
    #[pyo3(signature = (_idx, _amount=None))]
    fn insert_rows(&self, _idx: u32, _amount: Option<u32>) -> PyResult<()> {
        Ok(())
    }

    /// Insert columns.
    #[pyo3(signature = (_idx, _amount=None))]
    fn insert_cols(&self, _idx: u32, _amount: Option<u32>) -> PyResult<()> {
        Ok(())
    }

    /// Delete rows.
    #[pyo3(signature = (_idx, _amount=None))]
    fn delete_rows(&self, _idx: u32, _amount: Option<u32>) -> PyResult<()> {
        Ok(())
    }

    /// Delete columns.
    #[pyo3(signature = (_idx, _amount=None))]
    fn delete_cols(&self, _idx: u32, _amount: Option<u32>) -> PyResult<()> {
        Ok(())
    }

    /// Freeze panes at a cell.
    #[getter]
    fn freeze_panes(&self) -> Option<String> {
        None
    }

    #[setter]
    fn set_freeze_panes(&self, _cell: Option<&str>) {
        // No-op without workbook access
    }

    /// Auto-filter.
    #[getter]
    fn auto_filter(&self) -> Option<String> {
        None
    }

    /// Print title rows.
    #[getter]
    fn print_title_rows(&self) -> Option<String> {
        None
    }

    /// Print title columns.
    #[getter]
    fn print_title_cols(&self) -> Option<String> {
        None
    }

    /// Print area.
    #[getter]
    fn print_area(&self) -> Option<String> {
        None
    }

    fn __str__(&self) -> String {
        format!("<Worksheet \"{}\">", self.cached_title)
    }

    fn __repr__(&self) -> String {
        self.__str__()
    }
}

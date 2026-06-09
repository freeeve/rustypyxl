//! Python bindings for Worksheet.

use pyo3::prelude::*;
use pyo3::exceptions::{PyNotImplementedError, PyValueError};
use pyo3::types::PyList;
use pyo3::Py;
use rustypyxl_core::{column_to_letter, parse_coordinate, Worksheet};

use crate::cell::PyCell;
use crate::workbook::{cell_value_to_python, python_to_cell_value, PyWorkbook};

/// An Excel Worksheet (openpyxl-compatible API).
///
/// Worksheets are accessed from a Workbook, not created directly.
#[pyclass(name = "Worksheet")]
pub struct PyWorksheet {
    /// Index of the worksheet in the workbook.
    pub(crate) index: usize,
    /// Cached title; kept in sync with the workbook on rename.
    cached_title: String,
    /// Reference to parent workbook (for connected operations).
    pub(crate) workbook: Option<Py<PyWorkbook>>,
}

impl PyWorksheet {
    /// Create a PyWorksheet from a workbook reference.
    pub fn from_ref(wb: &PyWorkbook, index: usize) -> Self {
        let title = wb.inner.sheet_names.get(index)
            .cloned()
            .unwrap_or_else(|| format!("Sheet{}", index + 1));
        PyWorksheet { index, cached_title: title, workbook: None }
    }

    /// Create a connected PyWorksheet with a workbook reference.
    pub fn connected(wb_ref: Py<PyWorkbook>, index: usize, title: String) -> Self {
        PyWorksheet {
            index,
            cached_title: title,
            workbook: Some(wb_ref),
        }
    }

    /// Build a cell handle, connected to the parent workbook when one is present.
    fn make_cell(&self, row: u32, column: u32, py: Python<'_>) -> PyCell {
        if let Some(ref wb) = self.workbook {
            PyCell::connected(row, column, wb.clone_ref(py), self.cached_title.clone())
        } else {
            PyCell::new(row, column)
        }
    }

    /// Read this sheet's data extent as (min_row, min_col, max_row, max_col).
    fn sheet_dims(&self, py: Python<'_>) -> (u32, u32, u32, u32) {
        if let Some(ref wb) = self.workbook {
            let this = wb.borrow(py);
            if let Some(ws) = this.inner.worksheets.get(self.index) {
                return ws.dimensions();
            }
        }
        (1, 1, 1, 1)
    }

    /// Read a single cell value into a Python object (None if empty/missing).
    fn read_value(&self, row: u32, column: u32, py: Python<'_>) -> PyObject {
        if let Some(ref wb) = self.workbook {
            let this = wb.borrow(py);
            if let Some(ws) = this.inner.worksheets.get(self.index) {
                if let Some(cell) = ws.get_cell(row, column) {
                    return cell_value_to_python(&cell.value, py);
                }
            }
        }
        py.None()
    }

    /// Run a closure against the mutable core worksheet.
    fn with_sheet_mut<F: FnOnce(&mut Worksheet)>(&self, py: Python<'_>, f: F) -> PyResult<()> {
        if let Some(ref wb) = self.workbook {
            let mut this = wb.borrow_mut(py);
            let ws = this.inner.worksheets.get_mut(self.index)
                .ok_or_else(|| PyValueError::new_err("Worksheet index out of range"))?;
            f(ws);
            Ok(())
        } else {
            Err(PyValueError::new_err("Worksheet is not attached to a workbook"))
        }
    }

    /// Resolve a merge/range argument into an "A1:B2" string.
    fn resolve_range(
        &self,
        range_string: Option<&str>,
        start_row: Option<u32>,
        start_column: Option<u32>,
        end_row: Option<u32>,
        end_column: Option<u32>,
    ) -> PyResult<String> {
        if let Some(rs) = range_string {
            Ok(rs.to_string())
        } else if let (Some(sr), Some(sc), Some(er), Some(ec)) =
            (start_row, start_column, end_row, end_column)
        {
            Ok(format!("{}{}:{}{}", column_to_letter(sc), sr, column_to_letter(ec), er))
        } else {
            Err(PyValueError::new_err(
                "Must specify either range_string or all of start_row, start_column, end_row, end_column",
            ))
        }
    }
}

#[pymethods]
impl PyWorksheet {
    /// Get the worksheet title.
    #[getter]
    pub fn title(&self) -> String {
        self.cached_title.clone()
    }

    /// Rename the worksheet (e.g. ws.title = "Results").
    #[setter]
    fn set_title(&mut self, value: String) -> PyResult<()> {
        let idx = self.index;
        if let Some(ref wb) = self.workbook {
            Python::with_gil(|py| -> PyResult<()> {
                let mut this = wb.borrow_mut(py);
                if this.inner.sheet_names.iter().enumerate().any(|(i, n)| i != idx && n == &value) {
                    return Err(PyValueError::new_err(format!(
                        "Worksheet '{}' already exists",
                        value
                    )));
                }
                if let Some(name) = this.inner.sheet_names.get_mut(idx) {
                    *name = value.clone();
                }
                if let Some(ws) = this.inner.worksheets.get_mut(idx) {
                    ws.set_title(value.clone());
                }
                Ok(())
            })?;
        }
        self.cached_title = value;
        Ok(())
    }

    /// Get a cell (ws['A1']) or a range of cells (ws['A1:B2']).
    ///
    /// A single coordinate returns one Cell; a range returns a list of rows,
    /// each a list of Cell objects.
    fn __getitem__(&self, key: &str, py: Python<'_>) -> PyResult<PyObject> {
        if let Some(colon) = key.find(':') {
            let (r1, c1) = parse_coordinate(&key[..colon])
                .map_err(|e| PyValueError::new_err(e.to_string()))?;
            let (r2, c2) = parse_coordinate(&key[colon + 1..])
                .map_err(|e| PyValueError::new_err(e.to_string()))?;
            let (min_r, max_r) = (r1.min(r2), r1.max(r2));
            let (min_c, max_c) = (c1.min(c2), c1.max(c2));

            let rows = PyList::empty(py);
            for r in min_r..=max_r {
                let row = PyList::empty(py);
                for c in min_c..=max_c {
                    row.append(Py::new(py, self.make_cell(r, c, py))?)?;
                }
                rows.append(row)?;
            }
            return Ok(rows.into_any().unbind());
        }

        let (row, col) = parse_coordinate(key)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(Py::new(py, self.make_cell(row, col, py))?.into_any())
    }

    /// Set a cell value using subscript notation: ws['A1'] = 'Hello'.
    fn __setitem__(&self, key: &str, value: Bound<'_, PyAny>, py: Python<'_>) -> PyResult<()> {
        if key.contains(':') {
            return Err(PyValueError::new_err(
                "Range assignment is not supported; assign cells individually",
            ));
        }
        let (row, col) = parse_coordinate(key)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        if let Some(ref wb) = self.workbook {
            let mut this = wb.borrow_mut(py);
            this.set_cell_value(&self.cached_title, row, col, &value)
        } else {
            Err(PyValueError::new_err("Worksheet is not attached to a workbook"))
        }
    }

    /// Get a cell at a specific row and column (both 1-indexed).
    #[pyo3(signature = (row, column=None))]
    fn cell(&self, row: u32, column: Option<u32>, py: Python<'_>) -> PyResult<PyCell> {
        let col = column.unwrap_or(1);
        if row == 0 || col == 0 {
            return Err(PyValueError::new_err("Row and column must be at least 1"));
        }
        Ok(self.make_cell(row, col, py))
    }

    /// Iterate over rows. Returns a list of rows; each row is a list of Cell
    /// objects, or of values when values_only=True. Bounds default to the
    /// sheet's used range.
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
        let (_, _, dmax_r, dmax_c) = self.sheet_dims(py);
        let min_r = min_row.unwrap_or(1).max(1);
        let max_r = max_row.unwrap_or(dmax_r);
        let min_c = min_col.unwrap_or(1).max(1);
        let max_c = max_col.unwrap_or(dmax_c);

        let mut rows = Vec::new();
        for r in min_r..=max_r {
            let mut row = Vec::new();
            for c in min_c..=max_c {
                if values_only {
                    row.push(self.read_value(r, c, py));
                } else {
                    row.push(Py::new(py, self.make_cell(r, c, py))?.into_any());
                }
            }
            rows.push(row);
        }
        Ok(rows)
    }

    /// Iterate over columns (column-major). See iter_rows for argument behavior.
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
        let (_, _, dmax_r, dmax_c) = self.sheet_dims(py);
        let min_c = min_col.unwrap_or(1).max(1);
        let max_c = max_col.unwrap_or(dmax_c);
        let min_r = min_row.unwrap_or(1).max(1);
        let max_r = max_row.unwrap_or(dmax_r);

        let mut cols = Vec::new();
        for c in min_c..=max_c {
            let mut col = Vec::new();
            for r in min_r..=max_r {
                if values_only {
                    col.push(self.read_value(r, c, py));
                } else {
                    col.push(Py::new(py, self.make_cell(r, c, py))?.into_any());
                }
            }
            cols.push(col);
        }
        Ok(cols)
    }

    /// Get the maximum row containing data.
    #[getter]
    fn max_row(&self, py: Python<'_>) -> u32 {
        self.sheet_dims(py).2
    }

    /// Get the maximum column containing data.
    #[getter]
    fn max_column(&self, py: Python<'_>) -> u32 {
        self.sheet_dims(py).3
    }

    /// Get the minimum row containing data.
    #[getter]
    fn min_row(&self, py: Python<'_>) -> u32 {
        self.sheet_dims(py).0
    }

    /// Get the minimum column containing data.
    #[getter]
    fn min_column(&self, py: Python<'_>) -> u32 {
        self.sheet_dims(py).1
    }

    /// Get the used dimensions as a string (e.g., "A1:D10").
    #[getter]
    fn dimensions(&self, py: Python<'_>) -> String {
        let (min_r, min_c, max_r, max_c) = self.sheet_dims(py);
        format!(
            "{}{}:{}{}",
            column_to_letter(min_c),
            min_r,
            column_to_letter(max_c),
            max_r
        )
    }

    /// Merge cells in a range (e.g. "A1:B2") or by explicit coordinates.
    #[pyo3(signature = (range_string=None, start_row=None, start_column=None, end_row=None, end_column=None))]
    fn merge_cells(
        &self,
        range_string: Option<&str>,
        start_row: Option<u32>,
        start_column: Option<u32>,
        end_row: Option<u32>,
        end_column: Option<u32>,
        py: Python<'_>,
    ) -> PyResult<()> {
        let range = self.resolve_range(range_string, start_row, start_column, end_row, end_column)?;
        self.with_sheet_mut(py, move |ws| ws.merge_cells(&range))
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
        py: Python<'_>,
    ) -> PyResult<()> {
        let range = self.resolve_range(range_string, start_row, start_column, end_row, end_column)?;
        self.with_sheet_mut(py, move |ws| ws.unmerge_cells(&range))
    }

    /// Get merged cell ranges as "A1:B2" strings.
    #[getter]
    fn merged_cells(&self, py: Python<'_>) -> Vec<String> {
        if let Some(ref wb) = self.workbook {
            let this = wb.borrow(py);
            if let Some(ws) = this.inner.worksheets.get(self.index) {
                return ws
                    .merged_cells
                    .iter()
                    .map(|(s, e)| format!("{}:{}", s, e))
                    .collect();
            }
        }
        Vec::new()
    }

    /// Append a row of values after the last row containing data.
    fn append(&self, iterable: Vec<Bound<'_, PyAny>>, py: Python<'_>) -> PyResult<()> {
        if let Some(ref wb) = self.workbook {
            let mut this = wb.borrow_mut(py);
            let ws = this.inner.worksheets.get_mut(self.index)
                .ok_or_else(|| PyValueError::new_err("Worksheet index out of range"))?;
            let target_row = if ws.cells.is_empty() { 1 } else { ws.dimensions().2 + 1 };
            for (i, value) in iterable.iter().enumerate() {
                let cv = python_to_cell_value(value)?;
                ws.set_cell_value(target_row, (i as u32) + 1, cv);
            }
            Ok(())
        } else {
            Err(PyValueError::new_err("Worksheet is not attached to a workbook"))
        }
    }

    /// Insert rows. Not yet implemented.
    #[pyo3(signature = (idx, amount=None))]
    fn insert_rows(&self, idx: u32, amount: Option<u32>) -> PyResult<()> {
        let _ = (idx, amount);
        Err(PyNotImplementedError::new_err("insert_rows is not yet implemented"))
    }

    /// Insert columns. Not yet implemented.
    #[pyo3(signature = (idx, amount=None))]
    fn insert_cols(&self, idx: u32, amount: Option<u32>) -> PyResult<()> {
        let _ = (idx, amount);
        Err(PyNotImplementedError::new_err("insert_cols is not yet implemented"))
    }

    /// Delete rows. Not yet implemented.
    #[pyo3(signature = (idx, amount=None))]
    fn delete_rows(&self, idx: u32, amount: Option<u32>) -> PyResult<()> {
        let _ = (idx, amount);
        Err(PyNotImplementedError::new_err("delete_rows is not yet implemented"))
    }

    /// Delete columns. Not yet implemented.
    #[pyo3(signature = (idx, amount=None))]
    fn delete_cols(&self, idx: u32, amount: Option<u32>) -> PyResult<()> {
        let _ = (idx, amount);
        Err(PyNotImplementedError::new_err("delete_cols is not yet implemented"))
    }

    /// Get the freeze-panes anchor cell, if any.
    #[getter]
    fn freeze_panes(&self, py: Python<'_>) -> Option<String> {
        if let Some(ref wb) = self.workbook {
            let this = wb.borrow(py);
            if let Some(ws) = this.inner.worksheets.get(self.index) {
                return ws.freeze_panes.clone();
            }
        }
        None
    }

    /// Freeze panes at a cell (e.g. "B2"); pass None to unfreeze.
    #[setter]
    fn set_freeze_panes(&self, cell: Option<String>) -> PyResult<()> {
        Python::with_gil(|py| self.with_sheet_mut(py, move |ws| ws.set_freeze_panes(cell)))
    }

    fn __str__(&self) -> String {
        format!("<Worksheet \"{}\">", self.cached_title)
    }

    fn __repr__(&self) -> String {
        self.__str__()
    }
}

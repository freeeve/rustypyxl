//! Column and row dimension proxies, for openpyxl-style access:
//! `ws.column_dimensions['A'].width = 20` and `ws.row_dimensions[1].height = 15`.

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::Py;
use rustypyxl_core::{column_to_letter, letter_to_column};

use crate::workbook::PyWorkbook;

fn sheet_index(wb: &PyWorkbook, uid: u64) -> PyResult<usize> {
    wb.inner
        .sheet_index_by_uid(uid)
        .ok_or_else(|| PyValueError::new_err("Worksheet no longer exists in this workbook"))
}

/// The mapping returned by `ws.column_dimensions`; index by column letter.
#[pyclass(name = "ColumnDimensions")]
pub struct PyColumnDimensions {
    pub(crate) workbook: Py<PyWorkbook>,
    pub(crate) uid: u64,
}

#[pymethods]
impl PyColumnDimensions {
    fn __getitem__(&self, key: &str, py: Python<'_>) -> PyResult<PyColumnDimension> {
        let column = letter_to_column(key).map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(PyColumnDimension {
            workbook: self.workbook.clone_ref(py),
            uid: self.uid,
            column,
        })
    }
}

/// A single column's dimension (width). Setting `width` writes through to the
/// worksheet; reading returns the set width or None.
#[pyclass(name = "ColumnDimension")]
pub struct PyColumnDimension {
    workbook: Py<PyWorkbook>,
    uid: u64,
    column: u32,
}

#[pymethods]
impl PyColumnDimension {
    #[getter]
    fn width(&self, py: Python<'_>) -> PyResult<Option<f64>> {
        let this = self.workbook.borrow(py);
        let idx = sheet_index(&this, self.uid)?;
        Ok(this.inner.worksheets[idx].get_column_width(self.column))
    }

    #[setter]
    fn set_width(&self, py: Python<'_>, width: Option<f64>) -> PyResult<()> {
        let mut this = self.workbook.borrow_mut(py);
        let idx = sheet_index(&this, self.uid)?;
        if let Some(w) = width {
            this.inner.worksheets[idx].set_column_width(self.column, w);
        }
        Ok(())
    }

    /// The column letter this proxy addresses.
    #[getter]
    fn index(&self) -> String {
        column_to_letter(self.column)
    }
}

/// The mapping returned by `ws.row_dimensions`; index by 1-based row number.
#[pyclass(name = "RowDimensions")]
pub struct PyRowDimensions {
    pub(crate) workbook: Py<PyWorkbook>,
    pub(crate) uid: u64,
}

#[pymethods]
impl PyRowDimensions {
    fn __getitem__(&self, row: u32, py: Python<'_>) -> PyResult<PyRowDimension> {
        Ok(PyRowDimension {
            workbook: self.workbook.clone_ref(py),
            uid: self.uid,
            row,
        })
    }
}

/// A single row's dimension (height).
#[pyclass(name = "RowDimension")]
pub struct PyRowDimension {
    workbook: Py<PyWorkbook>,
    uid: u64,
    row: u32,
}

#[pymethods]
impl PyRowDimension {
    #[getter]
    fn height(&self, py: Python<'_>) -> PyResult<Option<f64>> {
        let this = self.workbook.borrow(py);
        let idx = sheet_index(&this, self.uid)?;
        Ok(this.inner.worksheets[idx].get_row_height(self.row))
    }

    #[setter]
    fn set_height(&self, py: Python<'_>, height: Option<f64>) -> PyResult<()> {
        let mut this = self.workbook.borrow_mut(py);
        let idx = sheet_index(&this, self.uid)?;
        if let Some(h) = height {
            this.inner.worksheets[idx].set_row_height(self.row, h);
        }
        Ok(())
    }

    /// The row number this proxy addresses.
    #[getter]
    fn index(&self) -> u32 {
        self.row
    }
}

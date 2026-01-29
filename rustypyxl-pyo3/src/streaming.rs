//! Python bindings for streaming workbook.

use pyo3::prelude::*;
use pyo3::exceptions::PyValueError;
use rustypyxl_core::streaming::{StreamingWorkbook, StreamingSheet};
use rustypyxl_core::CellValue;
use std::sync::Arc;

/// A write-only workbook that streams data directly to disk.
///
/// This uses minimal memory by writing rows immediately instead of
/// holding them in memory. Similar to openpyxl's write_only mode.
///
/// Example:
///     wb = WriteOnlyWorkbook("output.xlsx")
///     wb.create_sheet("Data")
///
///     for i in range(1_000_000):
///         wb.append_row([f"Row {i}", i, i * 1.5])
///
///     wb.close()
#[pyclass(name = "WriteOnlyWorkbook")]
pub struct PyStreamingWorkbook {
    inner: Option<StreamingWorkbook>,
    current_sheet: Option<StreamingSheet>,
}

#[pymethods]
impl PyStreamingWorkbook {
    /// Create a new write-only workbook.
    ///
    /// Args:
    ///     path: Path to save the Excel file
    #[new]
    fn new(path: &str) -> PyResult<Self> {
        let wb = StreamingWorkbook::new(path)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(PyStreamingWorkbook {
            inner: Some(wb),
            current_sheet: None,
        })
    }

    /// Create a new sheet.
    ///
    /// Note: Only one sheet can be open at a time. Creating a new sheet
    /// will finalize the previous one.
    ///
    /// Args:
    ///     name: Sheet name
    fn create_sheet(&mut self, name: &str) -> PyResult<()> {
        let wb = self.inner.as_mut()
            .ok_or_else(|| PyValueError::new_err("Workbook already closed"))?;

        let sheet = wb.create_sheet(name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        self.current_sheet = Some(sheet);
        Ok(())
    }

    /// Append a row to the current sheet.
    ///
    /// Args:
    ///     values: List of values (str, int, float, bool, or None)
    fn append_row(&mut self, values: Vec<PyObject>, py: Python<'_>) -> PyResult<()> {
        let wb = self.inner.as_mut()
            .ok_or_else(|| PyValueError::new_err("Workbook already closed"))?;

        let sheet = self.current_sheet.as_mut()
            .ok_or_else(|| PyValueError::new_err("No sheet created. Call create_sheet() first."))?;

        let cell_values: Vec<CellValue> = values
            .into_iter()
            .map(|v| python_to_cell_value(v, py))
            .collect();

        wb.append_row(sheet, cell_values)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        Ok(())
    }

    /// Close the workbook and finalize the file.
    ///
    /// This must be called to properly save the file.
    fn close(&mut self) -> PyResult<()> {
        let wb = self.inner.take()
            .ok_or_else(|| PyValueError::new_err("Workbook already closed"))?;

        let sheet = self.current_sheet.take()
            .ok_or_else(|| PyValueError::new_err("No sheet created"))?;

        wb.close(sheet)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        Ok(())
    }
}

fn python_to_cell_value(obj: PyObject, py: Python<'_>) -> CellValue {
    if obj.is_none(py) {
        return CellValue::Empty;
    }

    // Try bool first (before int, since bool is a subclass of int in Python)
    if let Ok(b) = obj.extract::<bool>(py) {
        return CellValue::Boolean(b);
    }

    // Try int
    if let Ok(i) = obj.extract::<i64>(py) {
        return CellValue::Number(i as f64);
    }

    // Try float
    if let Ok(f) = obj.extract::<f64>(py) {
        return CellValue::Number(f);
    }

    // Try string
    if let Ok(s) = obj.extract::<String>(py) {
        if s.starts_with('=') {
            return CellValue::Formula(s);
        }
        return CellValue::String(Arc::from(s.as_str()));
    }

    // Default to empty
    CellValue::Empty
}

/// Placeholder for sheet handle (not currently used).
#[pyclass(name = "WriteOnlySheet")]
pub struct PyStreamingSheet {}

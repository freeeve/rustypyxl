//! Python bindings for streaming workbook.

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use rustypyxl_core::streaming::{StreamingSheet, StreamingWorkbook};
use rustypyxl_core::CellValue;

/// A write-only workbook that streams data directly to disk.
///
/// This uses minimal memory by writing rows immediately instead of
/// holding them in memory. Similar to openpyxl's write_only mode.
///
/// Example:
///     with WriteOnlyWorkbook("output.xlsx") as wb:
///         wb.create_sheet("Data")
///         for i in range(1_000_000):
///             wb.append_row([f"Row {i}", i, i * 1.5])
///         wb.create_sheet("Summary")  # finalizes "Data" automatically
///         wb.append_row(["total", 1_000_000])
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
        let wb = StreamingWorkbook::new(path).map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(PyStreamingWorkbook {
            inner: Some(wb),
            current_sheet: None,
        })
    }

    /// Create a new sheet. Only one sheet is open at a time: creating a
    /// new sheet finalizes the previous one.
    ///
    /// Args:
    ///     name: Sheet name
    fn create_sheet(&mut self, name: &str) -> PyResult<()> {
        let wb = self
            .inner
            .as_mut()
            .ok_or_else(|| PyValueError::new_err("Workbook already closed"))?;

        let sheet = wb
            .create_sheet(name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        self.current_sheet = Some(sheet);
        Ok(())
    }

    /// Append a row to the current sheet.
    ///
    /// Args:
    ///     values: List of values (str, int, float, bool, or None)
    fn append_row(&mut self, values: Vec<PyObject>, py: Python<'_>) -> PyResult<()> {
        let wb = self
            .inner
            .as_mut()
            .ok_or_else(|| PyValueError::new_err("Workbook already closed"))?;

        let sheet = self
            .current_sheet
            .as_mut()
            .ok_or_else(|| PyValueError::new_err("No sheet created. Call create_sheet() first."))?;

        let cell_values: Vec<CellValue> = values
            .into_iter()
            .map(|v| crate::workbook::python_to_cell_value(v.bind(py)))
            .collect::<PyResult<Vec<_>>>()?;

        wb.append_row(sheet, cell_values)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        Ok(())
    }

    /// Close the workbook and finalize the file.
    ///
    /// This must be called (or the workbook used as a context manager) to
    /// produce a valid file; dropping without closing leaves it truncated.
    fn close(&mut self) -> PyResult<()> {
        if self.inner.is_none() {
            return Err(PyValueError::new_err("Workbook already closed"));
        }
        self.do_close()
    }

    /// Context-manager support: `with WriteOnlyWorkbook(path) as wb:`.
    fn __enter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    #[pyo3(signature = (exc_type=None, exc_value=None, traceback=None))]
    fn __exit__(
        &mut self,
        exc_type: Option<Bound<'_, PyAny>>,
        exc_value: Option<Bound<'_, PyAny>>,
        traceback: Option<Bound<'_, PyAny>>,
    ) -> PyResult<bool> {
        let _ = (exc_value, traceback);
        if self.inner.is_some() {
            if exc_type.is_none() {
                self.do_close()?;
            } else {
                // An exception is already propagating; finalize best-effort
                // without masking it
                let _ = self.do_close();
            }
        }
        Ok(false)
    }
}

impl PyStreamingWorkbook {
    /// Consume the inner workbook and finalize the file, with or without an
    /// open sheet.
    fn do_close(&mut self) -> PyResult<()> {
        let wb = self
            .inner
            .take()
            .ok_or_else(|| PyValueError::new_err("Workbook already closed"))?;
        let result = match self.current_sheet.take() {
            Some(sheet) => wb.close(sheet),
            None => wb.finish(),
        };
        result.map_err(|e| PyValueError::new_err(e.to_string()))
    }
}

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
    ///
    /// Holds the GIL for the duration. A single row is a few microseconds of
    /// Rust work, and releasing the GIL that often costs far more than it
    /// saves: each re-acquire has to wait out a competing thread's switch
    /// interval, which made a contended million-row write orders of magnitude
    /// slower. Use append_rows to hand a batch to Rust and release the GIL once.
    fn append_row(&mut self, values: Vec<PyObject>, py: Python<'_>) -> PyResult<()> {
        let cell_values: Vec<CellValue> = values
            .into_iter()
            .map(|v| crate::workbook::python_to_cell_value(v.bind(py)))
            .collect::<PyResult<Vec<_>>>()?;

        let (wb, sheet) = self.parts_mut()?;
        wb.append_row(sheet, cell_values)
            .map_err(|e| PyValueError::new_err(e.to_string()))
    }

    /// Append many rows at once.
    ///
    /// Converts the batch, then writes it with the GIL released, so other
    /// Python threads run while the rows are serialized, compressed, and
    /// written. This is the throughput path for large streams: prefer it over
    /// calling append_row in a loop.
    ///
    /// Args:
    ///     rows: Iterable of rows, each a list of values
    fn append_rows(&mut self, rows: Vec<Vec<PyObject>>, py: Python<'_>) -> PyResult<()> {
        let batch: Vec<Vec<CellValue>> = rows
            .into_iter()
            .map(|row| {
                row.into_iter()
                    .map(|v| crate::workbook::python_to_cell_value(v.bind(py)))
                    .collect::<PyResult<Vec<_>>>()
            })
            .collect::<PyResult<_>>()?;

        let (wb, sheet) = self.parts_mut()?;
        py.allow_threads(|| {
            for row in batch {
                wb.append_row(sheet, row)?;
            }
            Ok(())
        })
        .map_err(|e: rustypyxl_core::RustypyxlError| PyValueError::new_err(e.to_string()))
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
    /// The open workbook and its current sheet, or the matching error.
    fn parts_mut(&mut self) -> PyResult<(&mut StreamingWorkbook, &mut StreamingSheet)> {
        let wb = self
            .inner
            .as_mut()
            .ok_or_else(|| PyValueError::new_err("Workbook already closed"))?;
        let sheet = self
            .current_sheet
            .as_mut()
            .ok_or_else(|| PyValueError::new_err("No sheet created. Call create_sheet() first."))?;
        Ok((wb, sheet))
    }

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

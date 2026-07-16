//! Python bindings for rustypyxl - openpyxl-compatible Excel library.
//!
//! This crate provides Python bindings via PyO3 for the rustypyxl library,
//! offering an API compatible with openpyxl.

use pyo3::prelude::*;

mod cell;
mod dimensions;
mod streaming;
mod style;
mod workbook;
mod worksheet;

use cell::PyCell;
use streaming::PyStreamingWorkbook;
use style::{
    PyAlignment, PyBorder, PyColor, PyFont, PyGradientFill, PyGradientStop, PyPatternFill,
    PyProtection, PySide,
};
use workbook::{PyPivotTable, PyWorkbook};
use worksheet::{PyCellRangeIterator, PyWorksheet};

/// Load a workbook from a file path, bytes, or file-like object.
///
/// Args:
///     source: File path (str), bytes, or file-like object with .read() method
///     password: Password for a protected (encrypted) workbook, if any
///
/// Returns:
///     Workbook: The loaded workbook
///
/// Example:
///     wb = load_workbook('file.xlsx')
///     wb = load_workbook(file_bytes)
///     wb = load_workbook('protected.xlsx', password='secret')
#[pyfunction]
#[pyo3(signature = (source, password=None))]
fn load_workbook(source: &Bound<'_, PyAny>, password: Option<&str>) -> PyResult<PyWorkbook> {
    PyWorkbook::load(source, password)
}

/// Render a value the way Excel would display it under a number-format code.
///
/// Args:
///     value: A number, string, bool, or date/datetime/time.
///     number_format: An Excel format code (e.g. "0.00%", "yyyy-mm-dd",
///         "#,##0.00;[Red](#,##0.00)").
///
/// Returns:
///     str: The display string. Dates/datetimes are converted to their Excel
///     serial and rendered through the code's date tokens.
///
/// Example:
///     format_value(0.1234, "0.00%")      # "12.34%"
///     format_value(1234.5, "$#,##0.00")  # "$1,234.50"
#[pyfunction]
#[pyo3(signature = (value, number_format))]
fn format_value(value: &Bound<'_, PyAny>, number_format: &str) -> PyResult<String> {
    // Dates/datetimes/times become an Excel serial so the date tokens render.
    if workbook::is_datetime_like(value)? {
        let serial = datetime_to_serial(value)?;
        return Ok(rustypyxl_core::format_number(serial, number_format));
    }
    let cv = workbook::python_to_cell_value(value)?;
    Ok(rustypyxl_core::format_value(&cv, number_format))
}

/// Convert a Python date/datetime/time to an Excel serial (1900 date system).
fn datetime_to_serial(value: &Bound<'_, PyAny>) -> PyResult<f64> {
    // date and datetime expose toordinal; a bare time does not.
    let mut serial = match value.call_method0("toordinal") {
        Ok(ord) => (ord.extract::<i64>()? - 693594) as f64,
        Err(_) => 0.0,
    };
    let read = |name: &str| -> i64 {
        value
            .getattr(name)
            .ok()
            .and_then(|v| v.extract::<i64>().ok())
            .unwrap_or(0)
    };
    let seconds = read("hour") * 3600 + read("minute") * 60 + read("second");
    let micros = read("microsecond");
    serial += (seconds as f64 + micros as f64 / 1_000_000.0) / 86400.0;
    Ok(serial)
}

/// The rustypyxl Python module.
#[pymodule]
fn rustypyxl(m: &Bound<'_, PyModule>) -> PyResult<()> {
    // Core classes
    m.add_class::<PyWorkbook>()?;
    m.add_class::<PyPivotTable>()?;
    m.add_class::<PyWorksheet>()?;
    m.add_class::<dimensions::PyColumnDimensions>()?;
    m.add_class::<dimensions::PyColumnDimension>()?;
    m.add_class::<dimensions::PyRowDimensions>()?;
    m.add_class::<dimensions::PyRowDimension>()?;
    m.add_class::<dimensions::PyAutoFilter>()?;
    m.add_class::<PyCell>()?;
    m.add_class::<PyCellRangeIterator>()?;

    // Streaming (write-only) classes
    m.add_class::<PyStreamingWorkbook>()?;

    // Style classes
    m.add_class::<PyFont>()?;
    m.add_class::<PyAlignment>()?;
    m.add_class::<PyPatternFill>()?;
    m.add_class::<PyBorder>()?;
    m.add_class::<PySide>()?;
    m.add_class::<PyProtection>()?;
    m.add_class::<PyColor>()?;
    m.add_class::<PyGradientFill>()?;
    m.add_class::<PyGradientStop>()?;

    // Functions
    m.add_function(wrap_pyfunction!(load_workbook, m)?)?;
    m.add_function(wrap_pyfunction!(format_value, m)?)?;

    // Add submodule for styles (openpyxl compatibility)
    let styles = PyModule::new(m.py(), "styles")?;
    styles.add_class::<PyFont>()?;
    styles.add_class::<PyAlignment>()?;
    styles.add_class::<PyPatternFill>()?;
    styles.add_class::<PyBorder>()?;
    styles.add_class::<PySide>()?;
    styles.add_class::<PyProtection>()?;
    styles.add_class::<PyColor>()?;
    styles.add_class::<PyGradientFill>()?;
    styles.add_class::<PyGradientStop>()?;
    m.add_submodule(&styles)?;
    // add_submodule alone doesn't register the module with the import system,
    // so `from rustypyxl.styles import Font` would fail without this.
    m.py()
        .import("sys")?
        .getattr("modules")?
        .set_item("rustypyxl.styles", &styles)?;

    Ok(())
}

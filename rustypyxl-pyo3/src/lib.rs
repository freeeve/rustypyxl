//! Python bindings for rustypyxl - openpyxl-compatible Excel library.
//!
//! This crate provides Python bindings via PyO3 for the rustypyxl-core library,
//! offering an API compatible with openpyxl.

use pyo3::prelude::*;

mod cell;
mod streaming;
mod style;
mod workbook;
mod worksheet;

use cell::PyCell;
use streaming::PyStreamingWorkbook;
use style::{PyFont, PyAlignment, PyPatternFill, PyBorder, PySide, PyProtection, PyGradientFill, PyGradientStop};
use workbook::PyWorkbook;
use worksheet::PyWorksheet;

/// Load a workbook from a file path, bytes, or file-like object.
///
/// Args:
///     source: File path (str), bytes, or file-like object with .read() method
///
/// Returns:
///     Workbook: The loaded workbook
///
/// Example:
///     wb = load_workbook('file.xlsx')
///     wb = load_workbook(file_bytes)
///     wb = load_workbook(io.BytesIO(file_bytes))
#[pyfunction]
#[pyo3(signature = (source))]
fn load_workbook(source: &Bound<'_, PyAny>) -> PyResult<PyWorkbook> {
    PyWorkbook::load(source)
}

/// The rustypyxl Python module.
#[pymodule]
fn rustypyxl(m: &Bound<'_, PyModule>) -> PyResult<()> {
    // Core classes
    m.add_class::<PyWorkbook>()?;
    m.add_class::<PyWorksheet>()?;
    m.add_class::<PyCell>()?;

    // Streaming (write-only) classes
    m.add_class::<PyStreamingWorkbook>()?;

    // Style classes
    m.add_class::<PyFont>()?;
    m.add_class::<PyAlignment>()?;
    m.add_class::<PyPatternFill>()?;
    m.add_class::<PyBorder>()?;
    m.add_class::<PySide>()?;
    m.add_class::<PyProtection>()?;
    m.add_class::<PyGradientFill>()?;
    m.add_class::<PyGradientStop>()?;

    // Functions
    m.add_function(wrap_pyfunction!(load_workbook, m)?)?;

    // Add submodule for styles (openpyxl compatibility)
    let styles = PyModule::new(m.py(), "styles")?;
    styles.add_class::<PyFont>()?;
    styles.add_class::<PyAlignment>()?;
    styles.add_class::<PyPatternFill>()?;
    styles.add_class::<PyBorder>()?;
    styles.add_class::<PySide>()?;
    styles.add_class::<PyProtection>()?;
    styles.add_class::<PyGradientFill>()?;
    styles.add_class::<PyGradientStop>()?;
    m.add_submodule(&styles)?;

    Ok(())
}

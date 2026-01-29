//! Python bindings for rustypyxl - openpyxl-compatible Excel library.
//!
//! This crate provides Python bindings via PyO3 for the rustypyxl-core library,
//! offering an API compatible with openpyxl.

use pyo3::prelude::*;

mod cell;
mod style;
mod workbook;
mod worksheet;

use cell::PyCell;
use style::{PyFont, PyAlignment, PyPatternFill, PyBorder};
use workbook::PyWorkbook;
use worksheet::PyWorksheet;

/// Load a workbook from a file path.
///
/// Args:
///     filename: Path to the Excel file (.xlsx)
///
/// Returns:
///     Workbook: The loaded workbook
///
/// Example:
///     wb = load_workbook('file.xlsx')
#[pyfunction]
fn load_workbook(filename: &str) -> PyResult<PyWorkbook> {
    PyWorkbook::load(filename)
}

/// The rustypyxl Python module.
#[pymodule]
fn rustypyxl(m: &Bound<'_, PyModule>) -> PyResult<()> {
    // Core classes
    m.add_class::<PyWorkbook>()?;
    m.add_class::<PyWorksheet>()?;
    m.add_class::<PyCell>()?;

    // Style classes
    m.add_class::<PyFont>()?;
    m.add_class::<PyAlignment>()?;
    m.add_class::<PyPatternFill>()?;
    m.add_class::<PyBorder>()?;

    // Functions
    m.add_function(wrap_pyfunction!(load_workbook, m)?)?;

    // Add submodule for styles (openpyxl compatibility)
    let styles = PyModule::new(m.py(), "styles")?;
    styles.add_class::<PyFont>()?;
    styles.add_class::<PyAlignment>()?;
    styles.add_class::<PyPatternFill>()?;
    styles.add_class::<PyBorder>()?;
    m.add_submodule(&styles)?;

    Ok(())
}

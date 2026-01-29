//! Python bindings for Cell.

use pyo3::prelude::*;
use rustypyxl_core::column_to_letter;

use crate::style::{PyFont, PyAlignment};

/// An Excel Cell (openpyxl-compatible API).
#[pyclass(name = "Cell")]
pub struct PyCell {
    #[pyo3(get)]
    pub row: u32,
    #[pyo3(get)]
    pub column: u32,
    pub(crate) value_internal: Option<PyObject>,
    pub(crate) font_internal: Option<PyFont>,
    pub(crate) alignment_internal: Option<PyAlignment>,
    pub(crate) hyperlink_internal: Option<String>,
    pub(crate) comment_internal: Option<String>,
    pub(crate) number_format_internal: Option<String>,
}

impl PyCell {
    /// Create a new cell at the given position.
    pub fn new(row: u32, column: u32) -> Self {
        PyCell {
            row,
            column,
            value_internal: None,
            font_internal: None,
            alignment_internal: None,
            hyperlink_internal: None,
            comment_internal: None,
            number_format_internal: None,
        }
    }
}

#[pymethods]
impl PyCell {
    /// Create a new cell (for Python construction).
    #[new]
    fn py_new(row: u32, column: u32) -> Self {
        Self::new(row, column)
    }

    /// Get the cell value.
    #[getter]
    fn value(&self, py: Python<'_>) -> PyObject {
        match &self.value_internal {
            Some(val) => val.clone_ref(py),
            None => py.None(),
        }
    }

    /// Set the cell value.
    #[setter]
    fn set_value(&mut self, value: PyObject) {
        self.value_internal = Some(value);
    }

    /// Get the cell coordinate (e.g., "A1").
    #[getter]
    fn coordinate(&self) -> String {
        format!("{}{}", column_to_letter(self.column), self.row)
    }

    /// Get the column letter (e.g., "A").
    #[getter]
    fn column_letter(&self) -> String {
        column_to_letter(self.column)
    }

    /// Get the cell's font.
    #[getter]
    fn font(&self) -> Option<PyFont> {
        self.font_internal.clone()
    }

    /// Set the cell's font.
    #[setter]
    fn set_font(&mut self, font: PyFont) {
        self.font_internal = Some(font);
    }

    /// Get the cell's alignment.
    #[getter]
    fn alignment(&self) -> Option<PyAlignment> {
        self.alignment_internal.clone()
    }

    /// Set the cell's alignment.
    #[setter]
    fn set_alignment(&mut self, alignment: PyAlignment) {
        self.alignment_internal = Some(alignment);
    }

    /// Get the cell's hyperlink.
    #[getter]
    fn hyperlink(&self) -> Option<String> {
        self.hyperlink_internal.clone()
    }

    /// Set the cell's hyperlink.
    #[setter]
    fn set_hyperlink(&mut self, hyperlink: Option<String>) {
        self.hyperlink_internal = hyperlink;
    }

    /// Get the cell's comment.
    #[getter]
    fn comment(&self) -> Option<String> {
        self.comment_internal.clone()
    }

    /// Set the cell's comment.
    #[setter]
    fn set_comment(&mut self, comment: Option<String>) {
        self.comment_internal = comment;
    }

    /// Get the cell's number format.
    #[getter]
    fn number_format(&self) -> Option<String> {
        self.number_format_internal.clone()
    }

    /// Set the cell's number format.
    #[setter]
    fn set_number_format(&mut self, format: Option<String>) {
        self.number_format_internal = format;
    }

    /// Get the data type of the cell.
    #[getter]
    fn data_type(&self, py: Python<'_>) -> &str {
        if let Some(ref val) = self.value_internal {
            if val.is_none(py) {
                "n"
            } else if val.extract::<String>(py).is_ok() {
                "s"
            } else if val.extract::<f64>(py).is_ok() || val.extract::<i64>(py).is_ok() {
                "n"
            } else if val.extract::<bool>(py).is_ok() {
                "b"
            } else {
                "s"
            }
        } else {
            "n"
        }
    }

    /// Check if the cell contains a formula.
    #[getter]
    fn is_formula(&self, py: Python<'_>) -> bool {
        if let Some(ref val) = self.value_internal {
            if let Ok(s) = val.extract::<String>(py) {
                return s.starts_with('=');
            }
        }
        false
    }

    /// Offset returns a cell at a relative position.
    fn offset(&self, row: i32, column: i32) -> PyResult<PyCell> {
        let new_row = (self.row as i32 + row).max(1) as u32;
        let new_col = (self.column as i32 + column).max(1) as u32;
        Ok(PyCell::new(new_row, new_col))
    }

    fn __str__(&self, py: Python<'_>) -> String {
        let val_str = if let Some(ref v) = self.value_internal {
            if let Ok(s) = v.extract::<String>(py) {
                s
            } else if let Ok(n) = v.extract::<f64>(py) {
                n.to_string()
            } else if let Ok(b) = v.extract::<bool>(py) {
                b.to_string()
            } else {
                "None".to_string()
            }
        } else {
            "None".to_string()
        };
        format!("<Cell {}.{}>", self.coordinate(), val_str)
    }

    fn __repr__(&self, py: Python<'_>) -> String {
        self.__str__(py)
    }
}

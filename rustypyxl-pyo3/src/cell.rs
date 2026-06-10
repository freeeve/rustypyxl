//! Python bindings for Cell.

use pyo3::prelude::*;
use rustypyxl_core::column_to_letter;

use crate::style::{PyFont, PyAlignment, PyPatternFill, PyBorder, PyProtection};
use crate::workbook::PyWorkbook;

/// An Excel Cell (openpyxl-compatible API).
///
/// Cells can be either "connected" (with a reference back to the workbook) or
/// "disconnected" (standalone). Connected cells persist style changes to the workbook.
#[pyclass(name = "Cell")]
pub struct PyCell {
    #[pyo3(get)]
    pub row: u32,
    #[pyo3(get)]
    pub column: u32,
    pub(crate) value_internal: Option<PyObject>,
    pub(crate) font_internal: Option<PyFont>,
    pub(crate) alignment_internal: Option<PyAlignment>,
    pub(crate) fill_internal: Option<PyPatternFill>,
    pub(crate) border_internal: Option<PyBorder>,
    pub(crate) protection_internal: Option<PyProtection>,
    pub(crate) hyperlink_internal: Option<String>,
    pub(crate) comment_internal: Option<String>,
    pub(crate) number_format_internal: Option<String>,
    /// Reference to parent workbook (for connected cells).
    pub(crate) workbook: Option<Py<PyWorkbook>>,
    /// Stable uid of the owning sheet (for connected cells). Resolving by
    /// uid keeps the handle correct across sheet renames and reorders.
    pub(crate) sheet_uid: Option<u64>,
}

impl PyCell {
    /// Create a new disconnected cell at the given position.
    pub fn new(row: u32, column: u32) -> Self {
        PyCell {
            row,
            column,
            value_internal: None,
            font_internal: None,
            alignment_internal: None,
            fill_internal: None,
            border_internal: None,
            protection_internal: None,
            hyperlink_internal: None,
            comment_internal: None,
            number_format_internal: None,
            workbook: None,
            sheet_uid: None,
        }
    }

    /// Create a connected cell that persists changes to the workbook.
    pub fn connected(row: u32, column: u32, workbook: Py<PyWorkbook>, sheet_uid: u64) -> Self {
        PyCell {
            row,
            column,
            value_internal: None,
            font_internal: None,
            alignment_internal: None,
            fill_internal: None,
            border_internal: None,
            protection_internal: None,
            hyperlink_internal: None,
            comment_internal: None,
            number_format_internal: None,
            workbook: Some(workbook),
            sheet_uid: Some(sheet_uid),
        }
    }

    /// Check if this cell is connected to a workbook.
    pub fn is_connected(&self) -> bool {
        self.workbook.is_some() && self.sheet_uid.is_some()
    }

    /// Resolve the current name of this cell's sheet (None when detached).
    fn sheet_name(&self, py: Python<'_>) -> PyResult<Option<String>> {
        if let (Some(ref wb), Some(uid)) = (&self.workbook, self.sheet_uid) {
            let this = wb.borrow(py);
            let idx = this.inner.sheet_index_by_uid(uid).ok_or_else(|| {
                pyo3::exceptions::PyValueError::new_err(
                    "Worksheet no longer exists in this workbook",
                )
            })?;
            return Ok(Some(this.inner.sheet_names[idx].clone()));
        }
        Ok(None)
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
    fn value(&self, py: Python<'_>) -> PyResult<PyObject> {
        if let Some(sheet) = self.sheet_name(py)? {
            if let Some(ref wb) = self.workbook {
                let wb_ref = wb.borrow(py);
                return wb_ref.get_cell_value(&sheet, self.row, self.column, py);
            }
        }
        Ok(match &self.value_internal {
            Some(val) => val.clone_ref(py),
            None => py.None(),
        })
    }

    /// Set the cell value.
    #[setter]
    fn set_value(&mut self, py: Python<'_>, value: PyObject) -> PyResult<()> {
        if let Some(sheet) = self.sheet_name(py)? {
            if let Some(ref wb) = self.workbook {
                let mut wb_ref = wb.borrow_mut(py);
                let bound_value = value.bind(py);
                return wb_ref.set_cell_value(&sheet, self.row, self.column, bound_value);
            }
        }
        self.value_internal = Some(value);
        Ok(())
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
    fn font(&self, py: Python<'_>) -> PyResult<Option<PyFont>> {
        // If connected, get from workbook
        if let Some(sheet) = self.sheet_name(py)? {
            if let Some(ref wb) = self.workbook {
                let wb_ref = wb.borrow(py);
                return wb_ref.get_cell_font(&sheet, self.row, self.column);
            }
        }
        Ok(self.font_internal.clone())
    }

    /// Set the cell's font.
    #[setter]
    fn set_font(&mut self, py: Python<'_>, font: PyFont) -> PyResult<()> {
        // If connected, persist to workbook
        if let Some(sheet) = self.sheet_name(py)? {
            if let Some(ref wb) = self.workbook {
                let mut wb_ref = wb.borrow_mut(py);
                return wb_ref.set_cell_font(&sheet, self.row, self.column, &font);
            }
        }
        self.font_internal = Some(font);
        Ok(())
    }

    /// Get the cell's alignment.
    #[getter]
    fn alignment(&self, py: Python<'_>) -> PyResult<Option<PyAlignment>> {
        if let Some(sheet) = self.sheet_name(py)? {
            if let Some(ref wb) = self.workbook {
                let wb_ref = wb.borrow(py);
                return wb_ref.get_cell_alignment(&sheet, self.row, self.column);
            }
        }
        Ok(self.alignment_internal.clone())
    }

    /// Set the cell's alignment.
    #[setter]
    fn set_alignment(&mut self, py: Python<'_>, alignment: PyAlignment) -> PyResult<()> {
        if let Some(sheet) = self.sheet_name(py)? {
            if let Some(ref wb) = self.workbook {
                let mut wb_ref = wb.borrow_mut(py);
                return wb_ref.set_cell_alignment(&sheet, self.row, self.column, &alignment);
            }
        }
        self.alignment_internal = Some(alignment);
        Ok(())
    }

    /// Get the cell's fill (background color).
    #[getter]
    fn fill(&self, py: Python<'_>) -> PyResult<Option<PyPatternFill>> {
        if let Some(sheet) = self.sheet_name(py)? {
            if let Some(ref wb) = self.workbook {
                let wb_ref = wb.borrow(py);
                return wb_ref.get_cell_fill(&sheet, self.row, self.column);
            }
        }
        Ok(self.fill_internal.clone())
    }

    /// Set the cell's fill (background color).
    #[setter]
    fn set_fill(&mut self, py: Python<'_>, fill: PyPatternFill) -> PyResult<()> {
        if let Some(sheet) = self.sheet_name(py)? {
            if let Some(ref wb) = self.workbook {
                let mut wb_ref = wb.borrow_mut(py);
                return wb_ref.set_cell_fill(&sheet, self.row, self.column, &fill);
            }
        }
        self.fill_internal = Some(fill);
        Ok(())
    }

    /// Get the cell's border.
    #[getter]
    fn border(&self, py: Python<'_>) -> PyResult<Option<PyBorder>> {
        if let Some(sheet) = self.sheet_name(py)? {
            if let Some(ref wb) = self.workbook {
                let wb_ref = wb.borrow(py);
                return wb_ref.get_cell_border(&sheet, self.row, self.column);
            }
        }
        Ok(self.border_internal.clone())
    }

    /// Set the cell's border.
    #[setter]
    fn set_border(&mut self, py: Python<'_>, border: PyBorder) -> PyResult<()> {
        if let Some(sheet) = self.sheet_name(py)? {
            if let Some(ref wb) = self.workbook {
                let mut wb_ref = wb.borrow_mut(py);
                return wb_ref.set_cell_border(&sheet, self.row, self.column, &border);
            }
        }
        self.border_internal = Some(border);
        Ok(())
    }

    /// Get the cell's protection.
    #[getter]
    fn protection(&self, py: Python<'_>) -> PyResult<Option<PyProtection>> {
        if let Some(sheet) = self.sheet_name(py)? {
            if let Some(ref wb) = self.workbook {
                let wb_ref = wb.borrow(py);
                return wb_ref.get_cell_protection(&sheet, self.row, self.column);
            }
        }
        Ok(self.protection_internal.clone())
    }

    /// Set the cell's protection.
    #[setter]
    fn set_protection(&mut self, py: Python<'_>, protection: PyProtection) -> PyResult<()> {
        if let Some(sheet) = self.sheet_name(py)? {
            if let Some(ref wb) = self.workbook {
                let mut wb_ref = wb.borrow_mut(py);
                return wb_ref.set_cell_protection(&sheet, self.row, self.column, &protection);
            }
        }
        self.protection_internal = Some(protection);
        Ok(())
    }

    /// Get the cell's hyperlink.
    #[getter]
    fn hyperlink(&self, py: Python<'_>) -> PyResult<Option<String>> {
        if let Some(sheet) = self.sheet_name(py)? {
            if let Some(ref wb) = self.workbook {
                let wb_ref = wb.borrow(py);
                return wb_ref.get_cell_hyperlink(&sheet, self.row, self.column);
            }
        }
        Ok(self.hyperlink_internal.clone())
    }

    /// Set the cell's hyperlink.
    #[setter]
    fn set_hyperlink(&mut self, py: Python<'_>, hyperlink: Option<String>) -> PyResult<()> {
        if let Some(sheet) = self.sheet_name(py)? {
            if let Some(ref wb) = self.workbook {
                let mut wb_ref = wb.borrow_mut(py);
                return wb_ref.set_cell_hyperlink(&sheet, self.row, self.column, hyperlink);
            }
        }
        self.hyperlink_internal = hyperlink;
        Ok(())
    }

    /// Get the cell's comment.
    #[getter]
    fn comment(&self, py: Python<'_>) -> PyResult<Option<String>> {
        if let Some(sheet) = self.sheet_name(py)? {
            if let Some(ref wb) = self.workbook {
                let wb_ref = wb.borrow(py);
                return wb_ref.get_cell_comment(&sheet, self.row, self.column);
            }
        }
        Ok(self.comment_internal.clone())
    }

    /// Set the cell's comment.
    #[setter]
    fn set_comment(&mut self, py: Python<'_>, comment: Option<String>) -> PyResult<()> {
        if let Some(sheet) = self.sheet_name(py)? {
            if let Some(ref wb) = self.workbook {
                let mut wb_ref = wb.borrow_mut(py);
                return wb_ref.set_cell_comment(&sheet, self.row, self.column, comment);
            }
        }
        self.comment_internal = comment;
        Ok(())
    }

    /// Get the cell's number format.
    #[getter]
    fn number_format(&self, py: Python<'_>) -> PyResult<Option<String>> {
        if let Some(sheet) = self.sheet_name(py)? {
            if let Some(ref wb) = self.workbook {
                let wb_ref = wb.borrow(py);
                return wb_ref.get_cell_number_format(&sheet, self.row, self.column);
            }
        }
        Ok(self.number_format_internal.clone())
    }

    /// Set the cell's number format.
    #[setter]
    fn set_number_format(&mut self, py: Python<'_>, format: Option<String>) -> PyResult<()> {
        if let Some(sheet) = self.sheet_name(py)? {
            if let Some(ref wb) = self.workbook {
                let mut wb_ref = wb.borrow_mut(py);
                match format {
                    Some(ref fmt) => {
                        return wb_ref.set_cell_number_format(&sheet, self.row, self.column, fmt);
                    }
                    None => {
                        // Assigning None clears the workbook-side format
                        return wb_ref.clear_cell_number_format(&sheet, self.row, self.column);
                    }
                }
            }
        }
        self.number_format_internal = format;
        Ok(())
    }

    /// Get the data type of the cell: 'n' number, 's' string, 'b' bool, 'f' formula.
    #[getter]
    fn data_type(&self, py: Python<'_>) -> PyResult<&'static str> {
        let val = self.value(py)?;
        let bound = val.bind(py);
        if bound.is_none() {
            return Ok("n");
        }
        if let Ok(s) = bound.extract::<String>() {
            return Ok(if s.starts_with('=') { "f" } else { "s" });
        }
        if bound.extract::<bool>().is_ok() {
            return Ok("b");
        }
        if bound.extract::<f64>().is_ok() {
            return Ok("n");
        }
        Ok("s")
    }

    /// Check if the cell contains a formula.
    #[getter]
    fn is_formula(&self, py: Python<'_>) -> PyResult<bool> {
        let val = self.value(py)?;
        let bound = val.bind(py);
        if let Ok(s) = bound.extract::<String>() {
            return Ok(s.starts_with('='));
        }
        Ok(false)
    }

    /// Offset returns a cell at a relative position, preserving the workbook link.
    fn offset(&self, row: i32, column: i32, py: Python<'_>) -> PyResult<PyCell> {
        let new_row = (self.row as i32 + row).max(1) as u32;
        let new_col = (self.column as i32 + column).max(1) as u32;
        if let (Some(ref wb), Some(uid)) = (&self.workbook, self.sheet_uid) {
            Ok(PyCell::connected(new_row, new_col, wb.clone_ref(py), uid))
        } else {
            Ok(PyCell::new(new_row, new_col))
        }
    }

    fn __str__(&self, py: Python<'_>) -> String {
        match self.sheet_name(py) {
            Ok(Some(sheet)) => format!("<Cell '{}'.{}>", sheet, self.coordinate()),
            _ => format!("<Cell {}>", self.coordinate()),
        }
    }

    fn __repr__(&self, py: Python<'_>) -> String {
        self.__str__(py)
    }
}

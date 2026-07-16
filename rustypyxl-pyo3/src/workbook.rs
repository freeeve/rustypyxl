//! Python bindings for Workbook.

#![allow(deprecated)]

use pyo3::exceptions::{PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyBytes;
use rustypyxl_core::{
    Alignment, Border, BorderStyle, CellStyle, CellValue, CompressionLevel, Fill, Font, Protection,
    Workbook,
};
use std::sync::Arc;

use crate::style::{PyAlignment, PyBorder, PyFont, PyPatternFill, PyProtection, PySide};
use crate::worksheet::PyWorksheet;

/// An Excel Workbook (openpyxl-compatible API).
#[pyclass(name = "Workbook")]
pub struct PyWorkbook {
    pub(crate) inner: Workbook,
}

#[pymethods]
impl PyWorkbook {
    /// Create a new empty workbook.
    #[new]
    fn new() -> Self {
        PyWorkbook {
            inner: Workbook::new(),
        }
    }

    /// Load a workbook from a file path, bytes, or file-like object.
    ///
    /// Args:
    ///     source: File path (str or os.PathLike), bytes, or file-like object
    ///             with .read() method
    ///
    /// Returns:
    ///     Workbook: The loaded workbook
    #[staticmethod]
    #[pyo3(signature = (source, password=None))]
    pub fn load(source: &Bound<'_, PyAny>, password: Option<&str>) -> PyResult<Self> {
        let py = source.py();

        // A password opens an encrypted (or plain) workbook: resolve the source
        // to bytes and decrypt as needed.
        if let Some(pw) = password {
            let bytes = read_source_bytes(source)?;
            let inner = py
                .allow_threads(|| Workbook::load_from_bytes_with_password(&bytes, pw))
                .map_err(|e| PyValueError::new_err(e.to_string()))?;
            return Ok(PyWorkbook { inner });
        }

        // Check if source is bytes (before PathBuf, which str also satisfies)
        if let Ok(bytes) = source.extract::<Vec<u8>>() {
            let inner = py
                .allow_threads(|| Workbook::load_from_bytes(&bytes))
                .map_err(|e| PyValueError::new_err(e.to_string()))?;
            return Ok(PyWorkbook { inner });
        }

        // Check if source is a file path (str or os.PathLike, e.g. pathlib.Path)
        if let Ok(path) = source.extract::<std::path::PathBuf>() {
            let inner = py
                .allow_threads(|| Workbook::load(&path.to_string_lossy()))
                .map_err(|e| PyValueError::new_err(e.to_string()))?;
            return Ok(PyWorkbook { inner });
        }

        // Check if source has .read() method (file-like object)
        if source.hasattr("read")? {
            let bytes_obj = source.call_method0("read")?;
            let bytes = bytes_obj.extract::<Vec<u8>>()?;
            let inner = py
                .allow_threads(|| Workbook::load_from_bytes(&bytes))
                .map_err(|e| PyValueError::new_err(e.to_string()))?;
            return Ok(PyWorkbook { inner });
        }

        Err(PyTypeError::new_err(
            "Expected file path (str or os.PathLike), bytes, or file-like object with .read() method"
        ))
    }

    /// Get the active worksheet (the active tab from the loaded file, or
    /// the first sheet for new workbooks).
    #[getter]
    fn active(self_: Py<Self>, py: Python<'_>) -> PyResult<PyWorksheet> {
        let this = self_.borrow(py);
        if this.inner.worksheets.is_empty() {
            return Err(PyValueError::new_err("No worksheets in workbook"));
        }
        let idx = this.inner.active_sheet.min(this.inner.worksheets.len() - 1);
        let title = this
            .inner
            .sheet_names
            .get(idx)
            .cloned()
            .unwrap_or_else(|| "Sheet1".to_string());
        let uid = this.inner.worksheets[idx].uid;
        Ok(PyWorksheet::connected(self_.clone_ref(py), uid, title))
    }

    /// Set the active worksheet, by index or by worksheet, as openpyxl allows.
    #[setter]
    fn set_active(&mut self, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let index = if let Ok(ws) = value.extract::<PyRef<'_, PyWorksheet>>() {
            self.inner
                .sheet_index_by_uid(ws.uid)
                .ok_or_else(|| PyValueError::new_err("Worksheet is not in this workbook"))?
        } else if let Ok(index) = value.extract::<usize>() {
            index
        } else {
            return Err(PyTypeError::new_err(
                "active must be a worksheet or a sheet index",
            ));
        };

        if index >= self.inner.worksheets.len() {
            return Err(PyValueError::new_err(format!(
                "Sheet index {} is out of range ({} sheets)",
                index,
                self.inner.worksheets.len()
            )));
        }
        self.inner.active_sheet = index;
        Ok(())
    }

    /// Get all sheet names.
    #[getter]
    fn sheetnames(&self) -> Vec<String> {
        self.inner.sheet_names.clone()
    }

    /// Get all worksheets.
    #[getter]
    fn worksheets(self_: Py<Self>, py: Python<'_>) -> Vec<PyWorksheet> {
        let this = self_.borrow(py);
        (0..this.inner.worksheets.len())
            .map(|i| {
                let title = this
                    .inner
                    .sheet_names
                    .get(i)
                    .cloned()
                    .unwrap_or_else(|| format!("Sheet{}", i + 1));
                let uid = this.inner.worksheets[i].uid;
                PyWorksheet::connected(self_.clone_ref(py), uid, title)
            })
            .collect()
    }

    /// Get a worksheet by name using subscript notation: wb['Sheet1'].
    fn __getitem__(self_: Py<Self>, key: &str, py: Python<'_>) -> PyResult<PyWorksheet> {
        let this = self_.borrow(py);
        for (idx, name) in this.inner.sheet_names.iter().enumerate() {
            if name == key {
                let uid = this.inner.worksheets[idx].uid;
                return Ok(PyWorksheet::connected(
                    self_.clone_ref(py),
                    uid,
                    name.clone(),
                ));
            }
        }
        // openpyxl raises KeyError for unknown sheet names
        Err(pyo3::exceptions::PyKeyError::new_err(format!(
            "Worksheet '{}' does not exist",
            key
        )))
    }

    /// Check if a worksheet exists: 'Sheet1' in wb.
    fn __contains__(&self, key: &str) -> bool {
        self.inner.sheet_names.contains(&key.to_string())
    }

    /// Get the number of worksheets.
    fn __len__(&self) -> usize {
        self.inner.worksheets.len()
    }

    /// Iterate over worksheet names.
    fn __iter__(&self) -> PyResult<PySheetNameIterator> {
        Ok(PySheetNameIterator {
            names: self.inner.sheet_names.clone(),
            index: 0,
        })
    }

    /// Create a new worksheet.
    ///
    /// Args:
    ///     title: Optional worksheet title
    ///     index: Optional position to insert the worksheet
    ///
    /// Returns:
    ///     Worksheet: The newly created worksheet
    #[pyo3(signature = (title=None, index=None))]
    fn create_sheet(
        self_: Py<Self>,
        title: Option<String>,
        index: Option<usize>,
        py: Python<'_>,
    ) -> PyResult<PyWorksheet> {
        let final_idx;
        let sheet_title;
        let sheet_uid;
        {
            let mut this = self_.borrow_mut(py);
            this.inner
                .create_sheet(title)
                .map_err(|e| PyValueError::new_err(e.to_string()))?;

            // The sheet was appended at the end; move it to `index` if requested.
            let last = this.inner.worksheets.len() - 1;
            final_idx = match index {
                Some(i) if i < last => {
                    let ws = this.inner.worksheets.remove(last);
                    let name = this.inner.sheet_names.remove(last);
                    this.inner.worksheets.insert(i, ws);
                    this.inner.sheet_names.insert(i, name);
                    i
                }
                _ => last,
            };
            sheet_title = this.inner.sheet_names[final_idx].clone();
            sheet_uid = this.inner.worksheets[final_idx].uid;
        }
        Ok(PyWorksheet::connected(
            self_.clone_ref(py),
            sheet_uid,
            sheet_title,
        ))
    }

    /// Remove a worksheet.
    ///
    /// Args:
    ///     worksheet: The worksheet to remove (by name or PyWorksheet)
    fn remove(&mut self, worksheet: &PyWorksheet) -> PyResult<()> {
        let idx = worksheet.resolve_index(self)?;
        let name = self.inner.sheet_names[idx].clone();
        self.inner
            .remove_sheet(&name)
            .map_err(|e| PyValueError::new_err(e.to_string()))
    }

    /// Copy a worksheet.
    ///
    /// Args:
    ///     source: The worksheet to copy
    ///
    /// Returns:
    ///     Worksheet: The copied worksheet
    fn copy_worksheet(
        self_: Py<Self>,
        source: &PyWorksheet,
        py: Python<'_>,
    ) -> PyResult<PyWorksheet> {
        let new_name: String;
        let idx: usize;

        let new_uid;
        {
            let mut this = self_.borrow_mut(py);
            // Get the source worksheet's data
            let source_idx = source.resolve_index(&this)?;

            // Clone the worksheet under a fresh stable uid (a cloned uid
            // would make two handles resolve to the same identity)
            let src_ws = &this.inner.worksheets[source_idx];
            let mut new_ws = src_ws.clone();
            let base_name = format!("{} Copy", src_ws.title);
            new_uid = this.inner.allocate_sheet_uid();
            new_ws.uid = new_uid;

            // Generate a new unique name
            let mut counter = 1;
            new_name = base_name.clone();
            let mut temp_name = new_name.clone();
            while this.inner.sheet_names.contains(&temp_name) {
                temp_name = format!("{} {}", base_name, counter);
                counter += 1;
            }
            new_ws.set_title(&temp_name);

            this.inner.worksheets.push(new_ws);
            this.inner.sheet_names.push(temp_name.clone());

            idx = this.inner.worksheets.len() - 1;
            // Re-assign for return
            drop(this);
        }

        let this = self_.borrow(py);
        let sheet_title = this
            .inner
            .sheet_names
            .get(idx)
            .cloned()
            .unwrap_or_else(|| format!("Sheet{}", idx + 1));
        Ok(PyWorksheet::connected(
            self_.clone_ref(py),
            new_uid,
            sheet_title,
        ))
    }

    /// Move a worksheet within the workbook.
    fn move_sheet(&mut self, sheet: &PyWorksheet, offset: i32) -> PyResult<()> {
        let current_idx = sheet.resolve_index(self)?;

        let new_idx = (current_idx as i32 + offset).max(0) as usize;
        let new_idx = new_idx.min(self.inner.worksheets.len() - 1);

        if current_idx != new_idx {
            let ws = self.inner.worksheets.remove(current_idx);
            let name = self.inner.sheet_names.remove(current_idx);
            self.inner.worksheets.insert(new_idx, ws);
            self.inner.sheet_names.insert(new_idx, name);
        }

        Ok(())
    }

    /// Get the index of a worksheet.
    fn index(&self, worksheet: &PyWorksheet) -> PyResult<usize> {
        worksheet.resolve_index(self)
    }

    /// Create a named range.
    fn create_named_range(
        &mut self,
        name: String,
        worksheet: &PyWorksheet,
        range: String,
    ) -> PyResult<()> {
        let idx = worksheet.resolve_index(self)?;
        let ws_title = self.inner.sheet_names[idx].clone();
        let full_range = format!("'{}'!{}", ws_title, range);
        self.inner
            .create_named_range(name, full_range)
            .map_err(|e| PyValueError::new_err(e.to_string()))
    }

    /// Get all defined names (named ranges).
    #[getter]
    fn defined_names(&self) -> Vec<(String, String)> {
        self.inner
            .get_named_ranges()
            .iter()
            .map(|(n, r)| (n.to_string(), r.to_string()))
            .collect()
    }

    /// Save the workbook to a file.
    ///
    /// Args:
    ///     filename: Path to save the Excel file (str or os.PathLike)
    ///     password: Encrypt the file with this password (agile encryption)
    #[pyo3(signature = (filename, password=None))]
    fn save(
        &self,
        filename: std::path::PathBuf,
        password: Option<&str>,
        py: Python<'_>,
    ) -> PyResult<()> {
        let path = filename.to_string_lossy();
        py.allow_threads(|| match password {
            Some(pw) => self.inner.save_with_password(&path, pw),
            None => self.inner.save(&path),
        })
        .map_err(|e| PyValueError::new_err(e.to_string()))
    }

    /// Save the workbook to bytes.
    ///
    /// Args:
    ///     password: Encrypt the bytes with this password (agile encryption)
    ///
    /// Returns:
    ///     bytes: The workbook as an xlsx file in memory
    #[pyo3(signature = (password=None))]
    fn save_to_bytes<'py>(
        &self,
        password: Option<&str>,
        py: Python<'py>,
    ) -> PyResult<Bound<'py, PyBytes>> {
        let bytes = py
            .allow_threads(|| match password {
                Some(pw) => self.inner.save_to_bytes_with_password(pw),
                None => self.inner.save_to_bytes(),
            })
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(PyBytes::new(py, &bytes))
    }

    /// Set compression level for saving.
    ///
    /// Args:
    ///     level: Compression level - "none", "fast", "default", or "best"
    fn set_compression(&mut self, level: &str) -> PyResult<()> {
        self.inner.compression = match level.to_lowercase().as_str() {
            "none" | "stored" => CompressionLevel::None,
            "fast" | "1" => CompressionLevel::Fast,
            "default" | "6" => CompressionLevel::Default,
            "best" | "9" => CompressionLevel::Best,
            _ => {
                return Err(PyValueError::new_err(
                    "Invalid compression level. Use: 'none', 'fast', 'default', or 'best'",
                ))
            }
        };
        Ok(())
    }

    /// Close the workbook (no-op for compatibility).
    fn close(&self) {
        // No-op - we don't hold file handles open
    }

    /// Set a cell value in a specific sheet.
    ///
    /// This is the primary method for setting cell values.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///     value: Value to set (string, number, boolean, or None)
    pub fn set_cell_value(
        self_: Py<Self>,
        py: Python<'_>,
        sheet_name: &str,
        row: u32,
        column: u32,
        value: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        // Convert before borrowing: see write_rows
        let cell_value = python_to_cell_value(value)?;
        self_
            .borrow_mut(py)
            .set_converted_cell_value(sheet_name, row, column, cell_value)
    }

    /// Get a cell value from a specific sheet.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///
    /// Returns:
    ///     The cell value, or None if empty
    pub fn get_cell_value(
        &self,
        sheet_name: &str,
        row: u32,
        column: u32,
        py: Python<'_>,
    ) -> PyResult<PyObject> {
        let ws = self
            .inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        if let Some(cell) = ws.get_cell(row, column) {
            Ok(cell_value_to_python(&cell.value, py))
        } else {
            Ok(py.None())
        }
    }

    /// Evaluate a formula string in the context of a sheet and return the
    /// computed value (a number, string, bool, None for blank, or an Excel error
    /// string like "#DIV/0!"). See the formula engine's documented subset.
    pub fn evaluate_formula(
        &self,
        sheet_name: &str,
        formula: &str,
        py: Python<'_>,
    ) -> PyResult<PyObject> {
        let value = self
            .inner
            .evaluate_formula(sheet_name, formula)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(formula_value_to_python(value, py))
    }

    /// Evaluate the cell at 1-based (row, column): a formula cell is computed,
    /// any other cell yields its value, and a blank cell yields None.
    pub fn evaluate_cell(
        &self,
        sheet_name: &str,
        row: u32,
        column: u32,
        py: Python<'_>,
    ) -> PyResult<PyObject> {
        let value = self
            .inner
            .evaluate_cell(sheet_name, row, column)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(formula_value_to_python(value, py))
    }

    /// Evaluate every formula cell and store its computed result as the cell's
    /// cached value, so a saved file shows results without Excel recalculating.
    /// Returns the number of formula cells calculated.
    pub fn calculate_all(&mut self) -> usize {
        self.inner.calculate_all()
    }

    /// Create a pivot table from a source range and add it to a target sheet.
    ///
    /// `source_ref` is a range like "A1:C100" whose first row holds the field
    /// headers. `rows` and `columns` name fields for those areas; `values`
    /// is a list of (field, aggregation) pairs (aggregation e.g. "sum",
    /// "count", "average"). The pivot is written on save and Excel rebuilds its
    /// cache from the source on open.
    #[pyo3(signature = (source_sheet, source_ref, target_sheet, anchor, rows=Vec::new(), columns=Vec::new(), values=Vec::new(), name=None))]
    #[allow(clippy::too_many_arguments)]
    pub fn add_pivot_table(
        &mut self,
        source_sheet: &str,
        source_ref: &str,
        target_sheet: &str,
        anchor: &str,
        rows: Vec<String>,
        columns: Vec<String>,
        values: Vec<(String, String)>,
        name: Option<&str>,
    ) -> PyResult<()> {
        self.inner
            .add_pivot_table(
                source_sheet,
                source_ref,
                target_sheet,
                anchor,
                &rows,
                &columns,
                &values,
                name,
            )
            .map_err(|e| PyValueError::new_err(e.to_string()))
    }

    /// The pivot tables in this workbook, read-only (source range, cache
    /// fields, and row/column/data/page field placements). Pivot tables are
    /// preserved on save but not editable through this API.
    #[getter]
    pub fn pivot_tables(&self) -> Vec<PyPivotTable> {
        self.inner
            .pivot_tables()
            .into_iter()
            .map(PyPivotTable::from_info)
            .collect()
    }

    /// Rich-text runs of a cell as a list of dicts (text + font attributes), or
    /// None if the cell is not rich text.
    pub fn get_cell_rich_text(
        &self,
        sheet_name: &str,
        row: u32,
        column: u32,
        py: Python<'_>,
    ) -> PyResult<PyObject> {
        use pyo3::types::{PyDict, PyList};
        let ws = self
            .inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        let Some(cell) = ws.get_cell(row, column) else {
            return Ok(py.None());
        };
        let Some(rich) = &cell.rich_text else {
            return Ok(py.None());
        };
        let runs = PyList::empty(py);
        for run in &rich.runs {
            let d = PyDict::new(py);
            d.set_item("text", &run.text)?;
            if let Some(font) = &run.font {
                d.set_item("bold", font.bold)?;
                d.set_item("italic", font.italic)?;
                d.set_item("underline", font.underline.clone())?;
                d.set_item("strike", font.strike)?;
                d.set_item("size", font.size)?;
                d.set_item("color", font.color.as_ref().and_then(|c| c.rgb.clone()))?;
                d.set_item("name", font.name.clone())?;
                d.set_item("vert_align", font.vert_align.clone())?;
            }
            runs.append(d)?;
        }
        Ok(runs.into_any().unbind())
    }

    /// Write multiple rows of data to a sheet (bulk operation for performance).
    ///
    /// This is significantly faster than setting cells one at a time.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     data: List of rows, where each row is a list of values
    ///     start_row: Starting row (1-indexed, default 1)
    ///     start_col: Starting column (1-indexed, default 1)
    #[pyo3(signature = (sheet_name, data, start_row=1, start_col=1))]
    fn write_rows(
        self_: Py<Self>,
        py: Python<'_>,
        sheet_name: &str,
        data: Vec<Vec<Bound<'_, PyAny>>>,
        start_row: u32,
        start_col: u32,
    ) -> PyResult<()> {
        // Convert every value before borrowing the workbook: the conversion
        // falls back to __str__, which is arbitrary Python and may touch this
        // same workbook -- doing that under borrow_mut raises "Already borrowed".
        let rows: Vec<Vec<CellValue>> = data
            .iter()
            .map(|row| row.iter().map(python_to_cell_value).collect())
            .collect::<PyResult<_>>()?;

        let mut this = self_.borrow_mut(py);
        // Get mutable reference to worksheet once (avoid repeated lookups)
        let ws = this
            .inner
            .get_sheet_by_name_mut(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        for (row_idx, row_data) in rows.into_iter().enumerate() {
            let row = start_row + row_idx as u32;
            for (col_idx, cell_value) in row_data.into_iter().enumerate() {
                ws.set_cell_value(row, start_col + col_idx as u32, cell_value);
            }
        }
        Ok(())
    }

    /// Read all values from a sheet as a 2D list (bulk operation for performance).
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     min_row: Minimum row (1-indexed, default 1)
    ///     max_row: Maximum row (default: last row with data)
    ///     min_col: Minimum column (1-indexed, default 1)
    ///     max_col: Maximum column (default: last column with data)
    ///
    /// Returns:
    ///     List of rows, where each row is a list of values
    #[pyo3(signature = (sheet_name, min_row=None, max_row=None, min_col=None, max_col=None))]
    fn read_rows(
        &self,
        sheet_name: &str,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        py: Python<'_>,
    ) -> PyResult<Vec<Vec<PyObject>>> {
        let ws = self
            .inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        // dimensions() returns (min_row, min_col, max_row, max_col)
        let (dims_min_row, dims_min_col, dims_max_row, dims_max_col) = ws.dimensions();

        let min_r = min_row.unwrap_or(dims_min_row);
        let max_r = max_row.unwrap_or(dims_max_row);
        let min_c = min_col.unwrap_or(dims_min_col);
        let max_c = max_col.unwrap_or(dims_max_col);

        let mut result = Vec::new();
        for row in min_r..=max_r {
            let mut row_data = Vec::new();
            for col in min_c..=max_c {
                if let Some(cell) = ws.get_cell(row, col) {
                    row_data.push(cell_value_to_python(&cell.value, py));
                } else {
                    row_data.push(py.None());
                }
            }
            result.push(row_data);
        }
        Ok(result)
    }

    /// Set a cell's font style.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///     font: Font style to apply
    pub fn set_cell_font(
        &mut self,
        sheet_name: &str,
        row: u32,
        column: u32,
        font: &PyFont,
    ) -> PyResult<()> {
        let rust_font = pyfont_to_font(font);
        let style = rustypyxl_core::CellStyle::new().with_font(rust_font);
        self.set_or_merge_cell_style(sheet_name, row, column, style)
    }

    /// Set a cell's fill (background color).
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///     fill: PatternFill style to apply
    pub fn set_cell_fill(
        &mut self,
        sheet_name: &str,
        row: u32,
        column: u32,
        fill: &PyPatternFill,
    ) -> PyResult<()> {
        let rust_fill = pyfill_to_fill(fill);
        let style = rustypyxl_core::CellStyle::new().with_fill(rust_fill);
        self.set_or_merge_cell_style(sheet_name, row, column, style)
    }

    /// Set a cell's border.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///     border: Border style to apply
    pub fn set_cell_border(
        &mut self,
        sheet_name: &str,
        row: u32,
        column: u32,
        border: &PyBorder,
    ) -> PyResult<()> {
        let rust_border = pyborder_to_border(border);
        let style = rustypyxl_core::CellStyle::new().with_border(rust_border);
        self.set_or_merge_cell_style(sheet_name, row, column, style)
    }

    /// Set a cell's alignment.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///     alignment: Alignment style to apply
    pub fn set_cell_alignment(
        &mut self,
        sheet_name: &str,
        row: u32,
        column: u32,
        alignment: &PyAlignment,
    ) -> PyResult<()> {
        let rust_align = pyalignment_to_alignment(alignment);
        let style = rustypyxl_core::CellStyle::new().with_alignment(rust_align);
        self.set_or_merge_cell_style(sheet_name, row, column, style)
    }

    /// Set a cell's number format.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///     format: Number format string (e.g., "#,##0.00", "0.00%")
    pub fn set_cell_number_format(
        &mut self,
        sheet_name: &str,
        row: u32,
        column: u32,
        format: &str,
    ) -> PyResult<()> {
        let style = rustypyxl_core::CellStyle::new().with_number_format(format);
        self.set_or_merge_cell_style(sheet_name, row, column, style)
    }

    /// Set a cell's protection.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///     protection: Protection style to apply
    /// Remove a cell's number format while keeping its other style
    /// properties (the format lives on the style xf, so the style and its
    /// index must be re-resolved, not just the per-cell field).
    pub fn clear_cell_number_format(
        &mut self,
        sheet_name: &str,
        row: u32,
        column: u32,
    ) -> PyResult<()> {
        let cleared_style = {
            let ws = self
                .inner
                .get_sheet_by_name(sheet_name)
                .map_err(|e| PyValueError::new_err(e.to_string()))?;
            match ws.get_cell(row, column).and_then(|c| c.style.clone()) {
                Some(existing) if existing.number_format.is_some() => {
                    let mut cleared = (*existing).clone();
                    cleared.number_format = None;
                    Some(cleared)
                }
                _ => None,
            }
        };

        if let Some(style) = cleared_style {
            let style_index = self.inner.styles.get_or_add_cell_xf(&style);
            let ws = self
                .inner
                .get_sheet_by_name_mut(sheet_name)
                .map_err(|e| PyValueError::new_err(e.to_string()))?;
            let cell = ws.get_or_create_cell_mut(row, column);
            cell.style = Some(Arc::new(style));
            cell.style_index = Some(style_index as u32);
            cell.number_format = None;
        } else {
            let ws = self
                .inner
                .get_sheet_by_name_mut(sheet_name)
                .map_err(|e| PyValueError::new_err(e.to_string()))?;
            if ws.get_cell(row, column).is_some() {
                ws.get_or_create_cell_mut(row, column).number_format = None;
            }
        }
        Ok(())
    }

    pub fn set_cell_protection(
        &mut self,
        sheet_name: &str,
        row: u32,
        column: u32,
        protection: &PyProtection,
    ) -> PyResult<()> {
        let rust_protection = pyprotection_to_protection(protection);
        let style = rustypyxl_core::CellStyle::new().with_protection(rust_protection);
        self.set_or_merge_cell_style(sheet_name, row, column, style)
    }

    /// Set multiple style properties on a cell at once.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///     font: Optional font style
    ///     fill: Optional fill style
    ///     border: Optional border style
    ///     alignment: Optional alignment style
    ///     number_format: Optional number format string
    #[pyo3(signature = (sheet_name, row, column, font=None, fill=None, border=None, alignment=None, number_format=None))]
    // Mirrors a Python keyword-argument API
    #[allow(clippy::too_many_arguments)]
    fn set_cell_style(
        &mut self,
        sheet_name: &str,
        row: u32,
        column: u32,
        font: Option<&PyFont>,
        fill: Option<&PyPatternFill>,
        border: Option<&PyBorder>,
        alignment: Option<&PyAlignment>,
        number_format: Option<&str>,
    ) -> PyResult<()> {
        let mut style = rustypyxl_core::CellStyle::new();

        if let Some(f) = font {
            style = style.with_font(pyfont_to_font(f));
        }
        if let Some(f) = fill {
            style = style.with_fill(pyfill_to_fill(f));
        }
        if let Some(b) = border {
            style = style.with_border(pyborder_to_border(b));
        }
        if let Some(a) = alignment {
            style = style.with_alignment(pyalignment_to_alignment(a));
        }
        if let Some(nf) = number_format {
            style = style.with_number_format(nf);
        }

        self.set_or_merge_cell_style(sheet_name, row, column, style)
    }

    /// Get a cell's font style.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet
    ///     row: Row number (1-indexed)
    ///     column: Column number (1-indexed)
    ///
    /// Returns:
    ///     Font style or None if not set
    pub fn get_cell_font(
        &self,
        sheet_name: &str,
        row: u32,
        column: u32,
    ) -> PyResult<Option<PyFont>> {
        let ws = self
            .inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        if let Some(cell) = ws.get_cell(row, column) {
            if let Some(ref style) = cell.style {
                if let Some(ref font) = style.font {
                    return Ok(Some(font_to_pyfont(font)));
                }
            }
        }
        Ok(None)
    }

    /// Get a cell's fill style.
    pub fn get_cell_fill(
        &self,
        sheet_name: &str,
        row: u32,
        column: u32,
    ) -> PyResult<Option<PyPatternFill>> {
        let ws = self
            .inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        if let Some(cell) = ws.get_cell(row, column) {
            if let Some(ref style) = cell.style {
                if let Some(ref fill) = style.fill {
                    return Ok(Some(fill_to_pyfill(fill)));
                }
            }
        }
        Ok(None)
    }

    /// Get a cell's border style.
    pub fn get_cell_border(
        &self,
        sheet_name: &str,
        row: u32,
        column: u32,
    ) -> PyResult<Option<PyBorder>> {
        let ws = self
            .inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        if let Some(cell) = ws.get_cell(row, column) {
            if let Some(ref style) = cell.style {
                if let Some(ref border) = style.border {
                    return Ok(Some(border_to_pyborder(border)));
                }
            }
        }
        Ok(None)
    }

    /// Get a cell's alignment style.
    pub fn get_cell_alignment(
        &self,
        sheet_name: &str,
        row: u32,
        column: u32,
    ) -> PyResult<Option<PyAlignment>> {
        let ws = self
            .inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        if let Some(cell) = ws.get_cell(row, column) {
            if let Some(ref style) = cell.style {
                if let Some(ref align) = style.alignment {
                    return Ok(Some(alignment_to_pyalignment(align)));
                }
            }
        }
        Ok(None)
    }

    /// Get a cell's number format.
    pub fn get_cell_number_format(
        &self,
        sheet_name: &str,
        row: u32,
        column: u32,
    ) -> PyResult<Option<String>> {
        let ws = self
            .inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        if let Some(cell) = ws.get_cell(row, column) {
            if let Some(ref style) = cell.style {
                return Ok(style.number_format.as_deref().map(str::to_string));
            }
        }
        Ok(None)
    }

    /// Get a cell's protection style.
    pub fn get_cell_protection(
        &self,
        sheet_name: &str,
        row: u32,
        column: u32,
    ) -> PyResult<Option<PyProtection>> {
        let ws = self
            .inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        if let Some(cell) = ws.get_cell(row, column) {
            if let Some(ref style) = cell.style {
                if let Some(ref protection) = style.protection {
                    return Ok(Some(protection_to_pyprotection(protection)));
                }
            }
        }
        Ok(None)
    }

    /// Set a cell's hyperlink URL.
    pub fn set_cell_hyperlink(
        &mut self,
        sheet_name: &str,
        row: u32,
        column: u32,
        url: Option<String>,
    ) -> PyResult<()> {
        let ws = self
            .inner
            .get_sheet_by_name_mut(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        match url {
            Some(u) => ws.set_cell_hyperlink(row, column, u),
            None => {
                if let Some(cell) = ws.get_cell_mut(row, column) {
                    cell.hyperlink = None;
                }
            }
        }
        Ok(())
    }

    /// Get a cell's hyperlink URL, or None.
    pub fn get_cell_hyperlink(
        &self,
        sheet_name: &str,
        row: u32,
        column: u32,
    ) -> PyResult<Option<String>> {
        let ws = self
            .inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(ws.get_cell(row, column).and_then(|c| c.hyperlink.clone()))
    }

    /// Set a cell's comment text.
    pub fn set_cell_comment(
        &mut self,
        sheet_name: &str,
        row: u32,
        column: u32,
        comment: Option<String>,
    ) -> PyResult<()> {
        let ws = self
            .inner
            .get_sheet_by_name_mut(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        match comment {
            Some(c) => ws.set_cell_comment(row, column, c),
            None => {
                if let Some(cell) = ws.get_cell_mut(row, column) {
                    cell.comment = None;
                }
            }
        }
        Ok(())
    }

    /// Get a cell's comment text, or None.
    pub fn get_cell_comment(
        &self,
        sheet_name: &str,
        row: u32,
        column: u32,
    ) -> PyResult<Option<String>> {
        let ws = self
            .inner
            .get_sheet_by_name(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(ws.get_cell(row, column).and_then(|c| c.comment.clone()))
    }

    /// Import data from a Parquet file directly into a worksheet.
    ///
    /// This is the fastest way to load large datasets, as it bypasses
    /// Python FFI entirely and reads directly from Parquet into cells.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet to insert into
    ///     path: Path to the Parquet file
    ///     start_row: Starting row (1-indexed, default 1)
    ///     start_col: Starting column (1-indexed, default 1)
    ///     include_headers: Include column headers (default True)
    ///     column_renames: Dict mapping original column names to new names
    ///     columns: List of column names to import (None = all columns)
    ///
    /// Returns:
    ///     Dict with import results: rows_imported, columns_imported,
    ///     range (e.g. "A1:Z1000"), header_range, data_range, column_names
    #[cfg(feature = "parquet")]
    #[pyo3(signature = (sheet_name, path, start_row=1, start_col=1, include_headers=true, column_renames=None, columns=None))]
    // Mirrors a Python keyword-argument API
    #[allow(clippy::too_many_arguments)]
    fn insert_from_parquet(
        &mut self,
        sheet_name: &str,
        path: &str,
        start_row: u32,
        start_col: u32,
        include_headers: bool,
        column_renames: Option<std::collections::HashMap<String, String>>,
        columns: Option<Vec<String>>,
        py: Python<'_>,
    ) -> PyResult<PyObject> {
        use pyo3::types::PyDict;
        use rustypyxl_core::ParquetImportOptions;

        let mut opts = ParquetImportOptions::new().with_headers(include_headers);

        if let Some(renames) = column_renames {
            opts.column_renames = renames;
        }

        if let Some(cols) = columns {
            opts.columns = cols;
        }

        let inner = &mut self.inner;
        let result = py
            .allow_threads(|| {
                inner.insert_from_parquet(sheet_name, path, start_row, start_col, Some(opts))
            })
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        // Build result dict
        let dict = PyDict::new(py);
        dict.set_item("rows_imported", result.rows_imported)?;
        dict.set_item("columns_imported", result.columns_imported)?;
        dict.set_item("start_row", result.start_row)?;
        dict.set_item("start_col", result.start_col)?;
        dict.set_item("end_row", result.end_row)?;
        dict.set_item("end_col", result.end_col)?;
        dict.set_item("range", result.range_with_headers())?;
        dict.set_item("header_range", result.header_range())?;
        dict.set_item("data_range", result.data_range())?;
        dict.set_item("column_names", result.column_names)?;

        Ok(dict.into())
    }

    /// Export a worksheet to a Parquet file.
    ///
    /// This exports cell data directly to Parquet format with automatic type inference.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet to export
    ///     path: Output path for the Parquet file
    ///     has_headers: Whether the first row contains headers (default True)
    ///     compression: Compression type: "snappy", "gzip", "zstd", "lz4", "none" (default "snappy")
    ///     column_renames: Dict mapping original column names to new names
    ///     column_types: Dict mapping column names to types: "string", "float64", "int64", "boolean", "date", "datetime"
    ///
    /// Returns:
    ///     Dict with export results: rows_exported, columns_exported, column_names, file_size
    #[cfg(feature = "parquet")]
    #[pyo3(signature = (sheet_name, path, has_headers=true, compression="snappy", column_renames=None, column_types=None))]
    // Mirrors a Python keyword-argument API
    #[allow(clippy::too_many_arguments)]
    fn export_to_parquet(
        &self,
        sheet_name: &str,
        path: &str,
        has_headers: bool,
        compression: &str,
        column_renames: Option<std::collections::HashMap<String, String>>,
        column_types: Option<std::collections::HashMap<String, String>>,
        py: Python<'_>,
    ) -> PyResult<PyObject> {
        use pyo3::types::PyDict;
        use rustypyxl_core::{ColumnType, ParquetCompression, ParquetExportOptions};

        let compression = match compression.to_lowercase().as_str() {
            "none" => ParquetCompression::None,
            "snappy" => ParquetCompression::Snappy,
            "gzip" => ParquetCompression::Gzip,
            "zstd" => ParquetCompression::Zstd,
            "lz4" => ParquetCompression::Lz4,
            _ => {
                return Err(PyValueError::new_err(format!(
                    "Invalid compression: {}. Use 'none', 'snappy', 'gzip', 'zstd', or 'lz4'",
                    compression
                )))
            }
        };

        let mut opts = ParquetExportOptions::new()
            .with_headers(has_headers)
            .with_compression(compression);

        if let Some(renames) = column_renames {
            opts.column_renames = renames;
        }

        if let Some(types) = column_types {
            for (col_name, type_str) in types {
                let col_type = match type_str.to_lowercase().as_str() {
                    "string" | "str" => ColumnType::String,
                    "float64" | "float" | "double" => ColumnType::Float64,
                    "int64" | "int" | "integer" => ColumnType::Int64,
                    "boolean" | "bool" => ColumnType::Boolean,
                    "date" => ColumnType::Date,
                    "datetime" | "timestamp" => ColumnType::DateTime,
                    "auto" => ColumnType::Auto,
                    _ => return Err(PyValueError::new_err(format!(
                        "Invalid column type: {}. Use 'string', 'float64', 'int64', 'boolean', 'date', 'datetime', or 'auto'",
                        type_str
                    ))),
                };
                opts.column_types.insert(col_name, col_type);
            }
        }

        let result = py
            .allow_threads(|| self.inner.export_to_parquet(sheet_name, path, Some(opts)))
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        // Build result dict
        let dict = PyDict::new(py);
        dict.set_item("rows_exported", result.rows_exported)?;
        dict.set_item("columns_exported", result.columns_exported)?;
        dict.set_item("column_names", result.column_names)?;
        dict.set_item("file_size", result.file_size)?;

        Ok(dict.into())
    }

    /// Export a specific range from a worksheet to a Parquet file.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet to export
    ///     path: Output path for the Parquet file
    ///     min_row: Starting row (1-indexed)
    ///     min_col: Starting column (1-indexed)
    ///     max_row: Ending row (1-indexed)
    ///     max_col: Ending column (1-indexed)
    ///     has_headers: Whether the first row contains headers (default True)
    ///     compression: Compression type (default "snappy")
    ///
    /// Returns:
    ///     Dict with export results
    #[cfg(feature = "parquet")]
    #[pyo3(signature = (sheet_name, path, min_row, min_col, max_row, max_col, has_headers=true, compression="snappy"))]
    // Mirrors a Python keyword-argument API
    #[allow(clippy::too_many_arguments)]
    fn export_range_to_parquet(
        &self,
        sheet_name: &str,
        path: &str,
        min_row: u32,
        min_col: u32,
        max_row: u32,
        max_col: u32,
        has_headers: bool,
        compression: &str,
        py: Python<'_>,
    ) -> PyResult<PyObject> {
        use pyo3::types::PyDict;
        use rustypyxl_core::{ParquetCompression, ParquetExportOptions};

        let compression = match compression.to_lowercase().as_str() {
            "none" => ParquetCompression::None,
            "snappy" => ParquetCompression::Snappy,
            "gzip" => ParquetCompression::Gzip,
            "zstd" => ParquetCompression::Zstd,
            "lz4" => ParquetCompression::Lz4,
            _ => {
                return Err(PyValueError::new_err(format!(
                    "Invalid compression: {}",
                    compression
                )))
            }
        };

        let opts = ParquetExportOptions::new()
            .with_headers(has_headers)
            .with_compression(compression);

        let result = py
            .allow_threads(|| {
                self.inner.export_range_to_parquet(
                    sheet_name,
                    path,
                    min_row,
                    min_col,
                    max_row,
                    max_col,
                    Some(opts),
                )
            })
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        let dict = PyDict::new(py);
        dict.set_item("rows_exported", result.rows_exported)?;
        dict.set_item("columns_exported", result.columns_exported)?;
        dict.set_item("column_names", result.column_names)?;
        dict.set_item("file_size", result.file_size)?;

        Ok(dict.into())
    }

    /// Load a workbook from S3.
    ///
    /// Args:
    ///     bucket: S3 bucket name
    ///     key: S3 object key (path within the bucket)
    ///     region: Optional AWS region (e.g., "us-east-1")
    ///     endpoint_url: Optional custom endpoint URL (for S3-compatible services)
    ///
    /// Returns:
    ///     Workbook: The loaded workbook
    #[cfg(feature = "s3")]
    #[staticmethod]
    #[pyo3(signature = (bucket, key, region=None, endpoint_url=None))]
    pub fn load_from_s3(
        bucket: &str,
        key: &str,
        region: Option<&str>,
        endpoint_url: Option<&str>,
        py: Python<'_>,
    ) -> PyResult<Self> {
        use rustypyxl_core::S3Config;

        let mut config = S3Config::new();
        if let Some(r) = region {
            config = config.with_region(r);
        }
        if let Some(url) = endpoint_url {
            config = config.with_endpoint_url(url);
        }

        let inner = py
            .allow_threads(|| Workbook::load_from_s3(bucket, key, Some(config)))
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(PyWorkbook { inner })
    }

    /// Save the workbook to S3.
    ///
    /// Args:
    ///     bucket: S3 bucket name
    ///     key: S3 object key (path within the bucket)
    ///     region: Optional AWS region (e.g., "us-east-1")
    ///     endpoint_url: Optional custom endpoint URL (for S3-compatible services)
    #[cfg(feature = "s3")]
    #[pyo3(signature = (bucket, key, region=None, endpoint_url=None))]
    pub fn save_to_s3(
        &self,
        bucket: &str,
        key: &str,
        region: Option<&str>,
        endpoint_url: Option<&str>,
        py: Python<'_>,
    ) -> PyResult<()> {
        use rustypyxl_core::S3Config;

        let mut config = S3Config::new();
        if let Some(r) = region {
            config = config.with_region(r);
        }
        if let Some(url) = endpoint_url {
            config = config.with_endpoint_url(url);
        }

        py.allow_threads(|| self.inner.save_to_s3(bucket, key, Some(config)))
            .map_err(|e| PyValueError::new_err(e.to_string()))
    }

    fn __str__(&self) -> String {
        format!("<Workbook with {} sheet(s)>", self.inner.worksheets.len())
    }

    fn __repr__(&self) -> String {
        self.__str__()
    }
}

impl PyWorkbook {
    /// Store an already-converted value. Callers convert from Python first, so
    /// no arbitrary Python runs while the workbook is mutably borrowed.
    pub(crate) fn set_converted_cell_value(
        &mut self,
        sheet_name: &str,
        row: u32,
        column: u32,
        value: CellValue,
    ) -> PyResult<()> {
        self.inner
            .set_cell_value_in_sheet(sheet_name, row, column, value)
            .map_err(|e| PyValueError::new_err(e.to_string()))
    }

    /// Helper to set or merge a cell style with the existing style.
    fn set_or_merge_cell_style(
        &mut self,
        sheet_name: &str,
        row: u32,
        column: u32,
        new_style: CellStyle,
    ) -> PyResult<()> {
        // First, compute merged style from existing cell (if any)
        let merged_style = {
            let ws = self
                .inner
                .get_sheet_by_name(sheet_name)
                .map_err(|e| PyValueError::new_err(e.to_string()))?;

            if let Some(cell) = ws.get_cell(row, column) {
                if let Some(ref existing) = cell.style {
                    let mut merged = (**existing).clone();
                    if new_style.font.is_some() {
                        merged.font = new_style.font.clone();
                    }
                    if new_style.fill.is_some() {
                        merged.fill = new_style.fill.clone();
                    }
                    if new_style.border.is_some() {
                        merged.border = new_style.border.clone();
                    }
                    if new_style.alignment.is_some() {
                        merged.alignment = new_style.alignment.clone();
                    }
                    if new_style.number_format.is_some() {
                        merged.number_format = new_style.number_format.clone();
                    }
                    merged
                } else {
                    new_style
                }
            } else {
                new_style
            }
        };

        // Get the style index for this style
        let style_index = self.inner.styles.get_or_add_cell_xf(&merged_style);

        // Now get mutable reference and set the style
        let ws = self
            .inner
            .get_sheet_by_name_mut(sheet_name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        let cell = ws.get_or_create_cell_mut(row, column);
        cell.style = Some(Arc::new(merged_style));
        cell.style_index = Some(style_index as u32);

        Ok(())
    }
}

/// Iterator over worksheet names.
#[pyclass]
pub struct PySheetNameIterator {
    names: Vec<String>,
    index: usize,
}

#[pymethods]
impl PySheetNameIterator {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__(&mut self) -> Option<String> {
        if self.index < self.names.len() {
            let name = self.names[self.index].clone();
            self.index += 1;
            Some(name)
        } else {
            None
        }
    }
}

/// Convert a Python value to a CellValue.
pub(crate) fn python_to_cell_value(value: &Bound<'_, PyAny>) -> PyResult<CellValue> {
    if value.is_none() {
        return Ok(CellValue::Empty);
    }
    if let Ok(s) = value.extract::<String>() {
        // Store formula WITHOUT the leading '=' (it will be added back when written)
        return Ok(match s.strip_prefix('=') {
            Some(formula) => CellValue::Formula(formula.to_string()),
            None => CellValue::from(s),
        });
    }
    // bool before the numeric branches: bool is a subclass of int in Python
    if let Ok(b) = value.extract::<bool>() {
        return Ok(CellValue::Boolean(b));
    }
    if let Ok(n) = value.extract::<i64>() {
        return Ok(CellValue::Number(n as f64));
    }
    if let Ok(n) = value.extract::<f64>() {
        return Ok(CellValue::Number(n));
    }
    // datetime/date/time become ISO-8601 date cells (t="d"). The check goes
    // through the Python datetime module rather than pyo3's PyDateTime types,
    // which don't exist under abi3-forward-compatibility builds (the wheels
    // for Python versions newer than pyo3's tested range, e.g. 3.13/3.14).
    if is_datetime_like(value)? {
        let iso = value.call_method0("isoformat")?.extract::<String>()?;
        return Ok(CellValue::Date(iso));
    }
    // Try to convert to string as fallback
    Ok(CellValue::from(value.str()?.to_string()))
}

/// True when the value is a datetime.datetime, datetime.date, or
/// datetime.time instance (datetime subclasses date, so two checks suffice).
pub(crate) fn is_datetime_like(value: &Bound<'_, PyAny>) -> PyResult<bool> {
    let module = value.py().import("datetime")?;
    Ok(value.is_instance(&module.getattr("date")?)?
        || value.is_instance(&module.getattr("time")?)?)
}

/// Parse an ISO-8601 date/time string into the matching Python datetime,
/// date, or time object. Returns None if the string is not parseable.
fn iso_string_to_python(py: Python<'_>, s: &str) -> Option<PyObject> {
    let module = py.import("datetime").ok()?;
    let class = if s.contains('T') {
        "datetime"
    } else if s.contains(':') {
        "time"
    } else {
        "date"
    };
    module
        .getattr(class)
        .ok()?
        .call_method1("fromisoformat", (s,))
        .ok()
        .map(|obj| obj.unbind())
}

/// Convert a CellValue to a Python object.
/// Read a load source (bytes, a file path, or a file-like object) into bytes.
fn read_source_bytes(source: &Bound<'_, PyAny>) -> PyResult<Vec<u8>> {
    if let Ok(bytes) = source.extract::<Vec<u8>>() {
        return Ok(bytes);
    }
    if let Ok(path) = source.extract::<std::path::PathBuf>() {
        return std::fs::read(&path)
            .map_err(|e| PyValueError::new_err(format!("could not read {}: {e}", path.display())));
    }
    if source.hasattr("read")? {
        return source.call_method0("read")?.extract::<Vec<u8>>();
    }
    Err(PyTypeError::new_err(
        "Expected file path (str or os.PathLike), bytes, or file-like object with .read() method",
    ))
}

/// A read-only view of a pivot table (openpyxl-level read support).
#[pyclass(name = "PivotTable", frozen)]
pub struct PyPivotTable {
    /// Pivot table name.
    #[pyo3(get)]
    pub name: String,
    /// The cache id it draws from, if declared.
    #[pyo3(get)]
    pub cache_id: Option<u32>,
    /// The range the pivot occupies on its sheet.
    #[pyo3(get)]
    pub location: Option<String>,
    /// Source data sheet name.
    #[pyo3(get)]
    pub source_sheet: Option<String>,
    /// Source data range.
    #[pyo3(get)]
    pub source_ref: Option<String>,
    /// The cache field names, in source-column order.
    #[pyo3(get)]
    pub fields: Vec<String>,
    /// Row-area field names.
    #[pyo3(get)]
    pub row_fields: Vec<String>,
    /// Column-area field names.
    #[pyo3(get)]
    pub col_fields: Vec<String>,
    /// Report-filter field names.
    #[pyo3(get)]
    pub page_fields: Vec<String>,
    /// (display name, source field, subtotal) for each data field.
    data_fields: Vec<(String, String, String)>,
}

impl PyPivotTable {
    fn from_info(info: rustypyxl_core::pivot::PivotTableInfo) -> Self {
        PyPivotTable {
            name: info.name,
            cache_id: info.cache_id,
            location: info.location,
            source_sheet: info.source_sheet,
            source_ref: info.source_ref,
            fields: info.cache_fields,
            row_fields: info.row_fields,
            col_fields: info.col_fields,
            page_fields: info.page_fields,
            data_fields: info
                .data_fields
                .into_iter()
                .map(|d| (d.name, d.source_field, d.subtotal))
                .collect(),
        }
    }
}

#[pymethods]
impl PyPivotTable {
    /// Data (values) fields as a list of dicts with keys name, source_field,
    /// and subtotal (the aggregation, e.g. "sum").
    #[getter]
    fn data_fields(&self, py: Python<'_>) -> PyResult<PyObject> {
        use pyo3::types::{PyDict, PyList};
        let list = PyList::empty(py);
        for (name, source_field, subtotal) in &self.data_fields {
            let d = PyDict::new(py);
            d.set_item("name", name)?;
            d.set_item("source_field", source_field)?;
            d.set_item("subtotal", subtotal)?;
            list.append(d)?;
        }
        Ok(list.into_any().unbind())
    }

    fn __repr__(&self) -> String {
        format!(
            "PivotTable(name={:?}, source={:?}!{:?})",
            self.name, self.source_sheet, self.source_ref
        )
    }
}

pub(crate) fn cell_value_to_python(value: &CellValue, py: Python<'_>) -> PyObject {
    match value {
        CellValue::Empty => py.None(),
        CellValue::String(s) => s.as_ref().to_object(py),
        CellValue::Number(n) => {
            // openpyxl returns int for integral cells; xlsx numbers are f64,
            // so this is exact within the 2^53 integer range
            if n.fract() == 0.0 && n.is_finite() && n.abs() < 9.007_199_254_740_992e15 {
                (*n as i64).to_object(py)
            } else {
                n.to_object(py)
            }
        }
        CellValue::Boolean(b) => b.to_object(py),
        CellValue::Formula(f) => format!("={}", f).to_object(py),
        CellValue::Date(d) => iso_string_to_python(py, d).unwrap_or_else(|| d.to_object(py)),
    }
}

/// Convert an evaluated formula value to a Python object: numbers become
/// int/float, text a str, booleans a bool, blanks None, and Excel error values
/// their string form (e.g. "#DIV/0!").
pub(crate) fn formula_value_to_python(
    value: rustypyxl_core::FormulaValue,
    py: Python<'_>,
) -> PyObject {
    use rustypyxl_core::FormulaValue;
    match value {
        FormulaValue::Empty => py.None(),
        FormulaValue::Text(s) => s.to_object(py),
        FormulaValue::Bool(b) => b.to_object(py),
        FormulaValue::Error(e) => e.to_object(py),
        FormulaValue::Number(n) => {
            if n.fract() == 0.0 && n.is_finite() && n.abs() < 9.007_199_254_740_992e15 {
                (n as i64).to_object(py)
            } else {
                n.to_object(py)
            }
        }
    }
}

// =====================
// Style conversion helpers
// =====================

/// Convert PyFont to Rust Font.
fn pyfont_to_font(pf: &PyFont) -> Font {
    Font {
        name: pf.name.clone(),
        size: pf.size,
        bold: pf.bold,
        italic: pf.italic,
        underline: pf.underline.clone(),
        strike: pf.strike,
        color: pf.color.clone(),
        vert_align: pf.vertAlign.clone(),
    }
}

/// Convert Rust Font to PyFont.
fn font_to_pyfont(f: &Font) -> PyFont {
    PyFont {
        name: f.name.clone(),
        size: f.size,
        bold: f.bold,
        italic: f.italic,
        underline: f.underline.clone(),
        strike: f.strike,
        color: f.color.clone(),
        vertAlign: f.vert_align.clone(),
    }
}

/// Convert PyPatternFill to Rust Fill.
fn pyfill_to_fill(pf: &PyPatternFill) -> Fill {
    Fill {
        pattern_type: pf.fill_type.clone().or(pf.patternType.clone()),
        fg_color: pf.fgColor.clone(),
        bg_color: pf.bgColor.clone(),
    }
}

/// Convert Rust Fill to PyPatternFill.
fn fill_to_pyfill(f: &Fill) -> PyPatternFill {
    PyPatternFill {
        fill_type: f.pattern_type.clone(),
        fgColor: f.fg_color.clone(),
        bgColor: f.bg_color.clone(),
        patternType: f.pattern_type.clone(),
    }
}

/// Convert PySide to Rust BorderStyle.
fn pyside_to_borderstyle(ps: &PySide) -> Option<BorderStyle> {
    ps.style.as_ref().map(|s| BorderStyle {
        style: s.clone(),
        color: ps.color.clone(),
    })
}

/// Convert Rust BorderStyle to PySide.
fn borderstyle_to_pyside(bs: &BorderStyle) -> PySide {
    PySide {
        style: Some(bs.style.clone()),
        color: bs.color.clone(),
    }
}

/// Convert PyBorder to Rust Border.
fn pyborder_to_border(pb: &PyBorder) -> Border {
    Border {
        left: pb.left.as_ref().and_then(pyside_to_borderstyle),
        right: pb.right.as_ref().and_then(pyside_to_borderstyle),
        top: pb.top.as_ref().and_then(pyside_to_borderstyle),
        bottom: pb.bottom.as_ref().and_then(pyside_to_borderstyle),
        diagonal: pb.diagonal.as_ref().and_then(pyside_to_borderstyle),
    }
}

/// Convert Rust Border to PyBorder.
fn border_to_pyborder(b: &Border) -> PyBorder {
    PyBorder {
        left: b.left.as_ref().map(borderstyle_to_pyside),
        right: b.right.as_ref().map(borderstyle_to_pyside),
        top: b.top.as_ref().map(borderstyle_to_pyside),
        bottom: b.bottom.as_ref().map(borderstyle_to_pyside),
        diagonal: b.diagonal.as_ref().map(borderstyle_to_pyside),
        diagonal_direction: None,
        outline: true,
    }
}

/// Convert PyAlignment to Rust Alignment.
fn pyalignment_to_alignment(pa: &PyAlignment) -> Alignment {
    Alignment {
        horizontal: pa.horizontal.clone(),
        vertical: pa.vertical.clone(),
        wrap_text: pa.wrap_text,
        text_rotation: if pa.text_rotation != 0 {
            Some(pa.text_rotation)
        } else {
            None
        },
        indent: if pa.indent != 0 {
            Some(pa.indent)
        } else {
            None
        },
        shrink_to_fit: pa.shrink_to_fit,
    }
}

/// Convert Rust Alignment to PyAlignment.
fn alignment_to_pyalignment(a: &Alignment) -> PyAlignment {
    PyAlignment {
        horizontal: a.horizontal.clone(),
        vertical: a.vertical.clone(),
        wrap_text: a.wrap_text,
        shrink_to_fit: a.shrink_to_fit,
        indent: a.indent.unwrap_or(0),
        text_rotation: a.text_rotation.unwrap_or(0),
    }
}

/// Convert PyProtection to Rust Protection.
fn pyprotection_to_protection(pp: &PyProtection) -> Protection {
    Protection {
        locked: pp.locked,
        hidden: pp.hidden,
    }
}

/// Convert Rust Protection to PyProtection.
fn protection_to_pyprotection(p: &Protection) -> PyProtection {
    PyProtection {
        locked: p.locked,
        hidden: p.hidden,
    }
}

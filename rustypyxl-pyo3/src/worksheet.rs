//! Python bindings for Worksheet.

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};
use pyo3::Py;
use rustypyxl_core::{
    column_to_letter, coordinate_from_row_col, parse_coordinate, CellValue, Worksheet,
};

use crate::cell::PyCell;
use crate::workbook::{cell_value_to_python, python_to_cell_value, PyWorkbook};

/// An Excel Worksheet (openpyxl-compatible API).
///
/// Worksheets are accessed from a Workbook, not created directly.
#[pyclass(name = "Worksheet")]
pub struct PyWorksheet {
    /// Stable sheet uid; the single source of truth for resolving this
    /// handle. Survives sheet removal, reordering, and renames; a removed
    /// sheet makes the handle error instead of silently targeting whatever
    /// sheet now occupies its old position.
    pub(crate) uid: u64,
    /// Title at handle creation; only used for repr and error messages.
    cached_title: String,
    /// Reference to parent workbook (for connected operations).
    pub(crate) workbook: Option<Py<PyWorkbook>>,
}

impl PyWorksheet {
    /// Create a connected PyWorksheet with a workbook reference.
    pub fn connected(wb_ref: Py<PyWorkbook>, uid: u64, title: String) -> Self {
        PyWorksheet {
            uid,
            cached_title: title,
            workbook: Some(wb_ref),
        }
    }

    /// Resolve this handle's current position in the workbook.
    pub(crate) fn resolve_index(&self, this: &PyWorkbook) -> PyResult<usize> {
        this.inner.sheet_index_by_uid(self.uid).ok_or_else(|| {
            PyValueError::new_err(format!(
                "Worksheet '{}' no longer exists in this workbook",
                self.cached_title
            ))
        })
    }

    /// Resolve this handle's current sheet name.
    fn resolve_title(&self, py: Python<'_>) -> PyResult<String> {
        if let Some(ref wb) = self.workbook {
            let this = wb.borrow(py);
            let idx = self.resolve_index(&this)?;
            return Ok(this.inner.sheet_names[idx].clone());
        }
        Ok(self.cached_title.clone())
    }

    /// Build a cell handle, connected to the parent workbook when one is present.
    fn make_cell(&self, row: u32, column: u32, py: Python<'_>) -> PyCell {
        if let Some(ref wb) = self.workbook {
            PyCell::connected(row, column, wb.clone_ref(py), self.uid)
        } else {
            PyCell::new(row, column)
        }
    }

    /// Read this sheet's data extent as (min_row, min_col, max_row, max_col).
    fn sheet_dims(&self, py: Python<'_>) -> PyResult<(u32, u32, u32, u32)> {
        if let Some(ref wb) = self.workbook {
            let this = wb.borrow(py);
            let idx = self.resolve_index(&this)?;
            return Ok(this.inner.worksheets[idx].dimensions());
        }
        Ok((1, 1, 1, 1))
    }

    /// Run a closure against the immutable core worksheet, returning its result.
    fn with_sheet_ref<R, F: FnOnce(&Worksheet) -> R>(&self, py: Python<'_>, f: F) -> PyResult<R> {
        if let Some(ref wb) = self.workbook {
            let this = wb.borrow(py);
            let idx = self.resolve_index(&this)?;
            Ok(f(&this.inner.worksheets[idx]))
        } else {
            Err(PyValueError::new_err(
                "Worksheet is not attached to a workbook",
            ))
        }
    }

    /// Run a closure against the mutable core worksheet.
    fn with_sheet_mut<F: FnOnce(&mut Worksheet)>(&self, py: Python<'_>, f: F) -> PyResult<()> {
        if let Some(ref wb) = self.workbook {
            let mut this = wb.borrow_mut(py);
            let idx = self.resolve_index(&this)?;
            f(&mut this.inner.worksheets[idx]);
            Ok(())
        } else {
            Err(PyValueError::new_err(
                "Worksheet is not attached to a workbook",
            ))
        }
    }

    /// Resolve a merge/range argument into an "A1:B2" string.
    fn resolve_range(
        &self,
        range_string: Option<&str>,
        start_row: Option<u32>,
        start_column: Option<u32>,
        end_row: Option<u32>,
        end_column: Option<u32>,
    ) -> PyResult<String> {
        if let Some(rs) = range_string {
            Ok(rs.to_string())
        } else if let (Some(sr), Some(sc), Some(er), Some(ec)) =
            (start_row, start_column, end_row, end_column)
        {
            Ok(format!(
                "{}{}:{}{}",
                column_to_letter(sc),
                sr,
                column_to_letter(ec),
                er
            ))
        } else {
            Err(PyValueError::new_err(
                "Must specify either range_string or all of start_row, start_column, end_row, end_column",
            ))
        }
    }
}

#[pymethods]
impl PyWorksheet {
    /// Get the worksheet title (always the current name, even after the
    /// sheet was renamed through another handle).
    #[getter]
    pub fn title(&self, py: Python<'_>) -> String {
        self.resolve_title(py)
            .unwrap_or_else(|_| self.cached_title.clone())
    }

    /// Sheet visibility: "visible", "hidden", or "veryHidden" (openpyxl-compatible).
    #[getter]
    fn sheet_state(&self, py: Python<'_>) -> PyResult<String> {
        if let Some(ref wb) = self.workbook {
            let this = wb.borrow(py);
            let idx = self.resolve_index(&this)?;
            return Ok(this.inner.worksheets[idx].visibility.as_str().to_string());
        }
        Ok("visible".to_string())
    }

    /// Set sheet visibility: "visible", "hidden", or "veryHidden".
    #[setter]
    fn set_sheet_state(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        if !matches!(value, "visible" | "hidden" | "veryHidden") {
            return Err(PyValueError::new_err(
                "sheet_state must be 'visible', 'hidden', or 'veryHidden'",
            ));
        }
        let state = rustypyxl_core::SheetVisibility::from_attr(value);
        self.with_sheet_mut(py, |ws| ws.visibility = state)
    }

    /// Rename the worksheet (e.g. ws.title = "Results").
    #[setter]
    fn set_title(&mut self, value: String) -> PyResult<()> {
        if let Some(ref wb) = self.workbook {
            Python::with_gil(|py| -> PyResult<()> {
                let mut this = wb.borrow_mut(py);
                let idx = self.resolve_index(&this)?;
                if this
                    .inner
                    .sheet_names
                    .iter()
                    .enumerate()
                    .any(|(i, n)| i != idx && n == &value)
                {
                    return Err(PyValueError::new_err(format!(
                        "Worksheet '{}' already exists",
                        value
                    )));
                }
                this.inner.sheet_names[idx] = value.clone();
                this.inner.worksheets[idx].set_title(value.clone());
                Ok(())
            })?;
        }
        self.cached_title = value;
        Ok(())
    }

    /// Get a cell (ws['A1']) or a range of cells (ws['A1:B2']).
    ///
    /// A single coordinate returns one Cell; a range returns a list of rows,
    /// each a list of Cell objects.
    fn __getitem__(&self, key: &str, py: Python<'_>) -> PyResult<PyObject> {
        if let Some(colon) = key.find(':') {
            let (r1, c1) = parse_coordinate(&key[..colon])
                .map_err(|e| PyValueError::new_err(e.to_string()))?;
            let (r2, c2) = parse_coordinate(&key[colon + 1..])
                .map_err(|e| PyValueError::new_err(e.to_string()))?;
            let (min_r, max_r) = (r1.min(r2), r1.max(r2));
            let (min_c, max_c) = (c1.min(c2), c1.max(c2));

            let rows = PyList::empty(py);
            for r in min_r..=max_r {
                let row = PyList::empty(py);
                for c in min_c..=max_c {
                    row.append(Py::new(py, self.make_cell(r, c, py))?)?;
                }
                rows.append(row)?;
            }
            return Ok(rows.into_any().unbind());
        }

        let (row, col) = parse_coordinate(key).map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(Py::new(py, self.make_cell(row, col, py))?.into_any())
    }

    /// Set a cell value using subscript notation: ws['A1'] = 'Hello'.
    fn __setitem__(&self, key: &str, value: Bound<'_, PyAny>, py: Python<'_>) -> PyResult<()> {
        if key.contains(':') {
            return Err(PyValueError::new_err(
                "Range assignment is not supported; assign cells individually",
            ));
        }
        let (row, col) = parse_coordinate(key).map_err(|e| PyValueError::new_err(e.to_string()))?;
        // Convert before borrowing the workbook: the conversion can run
        // arbitrary Python (__str__), which may re-enter this workbook.
        let cell_value = python_to_cell_value(&value)?;
        if let Some(ref wb) = self.workbook {
            let mut this = wb.borrow_mut(py);
            let idx = self.resolve_index(&this)?;
            let name = this.inner.sheet_names[idx].clone();
            this.set_converted_cell_value(&name, row, col, cell_value)
        } else {
            Err(PyValueError::new_err(
                "Worksheet is not attached to a workbook",
            ))
        }
    }

    /// Get a cell at a specific row and column (both 1-indexed).
    #[pyo3(signature = (row, column=None))]
    fn cell(&self, row: u32, column: Option<u32>, py: Python<'_>) -> PyResult<PyCell> {
        let col = column.unwrap_or(1);
        if row == 0 || col == 0 {
            return Err(PyValueError::new_err("Row and column must be at least 1"));
        }
        Ok(self.make_cell(row, col, py))
    }

    /// Iterate over rows lazily, like openpyxl: yields one tuple per row,
    /// of Cell objects (or raw values when values_only=True). Bounds default
    /// to the sheet's used range.
    #[pyo3(signature = (min_row=None, max_row=None, min_col=None, max_col=None, values_only=false))]
    fn iter_rows(
        &self,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        values_only: bool,
        py: Python<'_>,
    ) -> PyResult<PyCellRangeIterator> {
        let (_, _, dmax_r, dmax_c) = self.sheet_dims(py)?;
        Ok(PyCellRangeIterator {
            workbook: self.workbook.as_ref().map(|wb| wb.clone_ref(py)),
            sheet_uid: self.uid,
            min_row: min_row.unwrap_or(1).max(1),
            max_row: max_row.unwrap_or(dmax_r),
            min_col: min_col.unwrap_or(1).max(1),
            max_col: max_col.unwrap_or(dmax_c),
            values_only,
            by_columns: false,
            position: min_row.unwrap_or(1).max(1),
        })
    }

    /// Iterate over columns lazily (one tuple per column). See iter_rows.
    #[pyo3(signature = (min_col=None, max_col=None, min_row=None, max_row=None, values_only=false))]
    fn iter_cols(
        &self,
        min_col: Option<u32>,
        max_col: Option<u32>,
        min_row: Option<u32>,
        max_row: Option<u32>,
        values_only: bool,
        py: Python<'_>,
    ) -> PyResult<PyCellRangeIterator> {
        let (_, _, dmax_r, dmax_c) = self.sheet_dims(py)?;
        Ok(PyCellRangeIterator {
            workbook: self.workbook.as_ref().map(|wb| wb.clone_ref(py)),
            sheet_uid: self.uid,
            min_row: min_row.unwrap_or(1).max(1),
            max_row: max_row.unwrap_or(dmax_r),
            min_col: min_col.unwrap_or(1).max(1),
            max_col: max_col.unwrap_or(dmax_c),
            values_only,
            by_columns: true,
            position: min_col.unwrap_or(1).max(1),
        })
    }

    /// Get the maximum row containing data.
    #[getter]
    fn max_row(&self, py: Python<'_>) -> PyResult<u32> {
        Ok(self.sheet_dims(py)?.2)
    }

    /// Get the maximum column containing data.
    #[getter]
    fn max_column(&self, py: Python<'_>) -> PyResult<u32> {
        Ok(self.sheet_dims(py)?.3)
    }

    /// Get the minimum row containing data.
    #[getter]
    fn min_row(&self, py: Python<'_>) -> PyResult<u32> {
        Ok(self.sheet_dims(py)?.0)
    }

    /// Get the minimum column containing data.
    #[getter]
    fn min_column(&self, py: Python<'_>) -> PyResult<u32> {
        Ok(self.sheet_dims(py)?.1)
    }

    /// Get the used dimensions as a string (e.g., "A1:D10").
    #[getter]
    fn dimensions(&self, py: Python<'_>) -> PyResult<String> {
        let (min_r, min_c, max_r, max_c) = self.sheet_dims(py)?;
        Ok(format!(
            "{}{}:{}{}",
            column_to_letter(min_c),
            min_r,
            column_to_letter(max_c),
            max_r
        ))
    }

    /// Merge cells in a range (e.g. "A1:B2") or by explicit coordinates.
    #[pyo3(signature = (range_string=None, start_row=None, start_column=None, end_row=None, end_column=None))]
    fn merge_cells(
        &self,
        range_string: Option<&str>,
        start_row: Option<u32>,
        start_column: Option<u32>,
        end_row: Option<u32>,
        end_column: Option<u32>,
        py: Python<'_>,
    ) -> PyResult<()> {
        let range =
            self.resolve_range(range_string, start_row, start_column, end_row, end_column)?;
        self.with_sheet_mut(py, move |ws| ws.merge_cells(&range))
    }

    /// Unmerge cells in a range.
    #[pyo3(signature = (range_string=None, start_row=None, start_column=None, end_row=None, end_column=None))]
    fn unmerge_cells(
        &self,
        range_string: Option<&str>,
        start_row: Option<u32>,
        start_column: Option<u32>,
        end_row: Option<u32>,
        end_column: Option<u32>,
        py: Python<'_>,
    ) -> PyResult<()> {
        let range =
            self.resolve_range(range_string, start_row, start_column, end_row, end_column)?;
        self.with_sheet_mut(py, move |ws| ws.unmerge_cells(&range))
    }

    /// Get merged cell ranges as "A1:B2" strings.
    #[getter]
    fn merged_cells(&self, py: Python<'_>) -> PyResult<Vec<String>> {
        if let Some(ref wb) = self.workbook {
            let this = wb.borrow(py);
            let idx = self.resolve_index(&this)?;
            return Ok(this.inner.worksheets[idx]
                .merged_cells
                .iter()
                .map(|(s, e)| format!("{}:{}", s, e))
                .collect());
        }
        Ok(Vec::new())
    }

    /// Append a row after the last row containing data. Accepts any
    /// iterable of values (list, tuple, generator), or a dict mapping
    /// column letters or 1-based indices to values, like openpyxl.
    fn append(&self, iterable: Bound<'_, PyAny>, py: Python<'_>) -> PyResult<()> {
        // Collect (column, value) pairs before borrowing the workbook, since
        // evaluating a generator can run arbitrary Python code
        let mut cells: Vec<(u32, rustypyxl_core::CellValue)> = Vec::new();
        if let Ok(dict) = iterable.downcast::<pyo3::types::PyDict>() {
            for (key, value) in dict.iter() {
                let column = if let Ok(idx) = key.extract::<u32>() {
                    idx
                } else if let Ok(letter) = key.extract::<String>() {
                    let (_, col) = parse_coordinate(&format!("{}1", letter)).map_err(|_| {
                        PyValueError::new_err(format!("Invalid column key '{}'", letter))
                    })?;
                    col
                } else {
                    return Err(PyValueError::new_err(
                        "dict keys must be column letters or 1-based indices",
                    ));
                };
                if column == 0 {
                    return Err(PyValueError::new_err("Column index must be at least 1"));
                }
                cells.push((column, python_to_cell_value(&value)?));
            }
        } else {
            for (i, item) in iterable.try_iter()?.enumerate() {
                cells.push(((i as u32) + 1, python_to_cell_value(&item?)?));
            }
        }

        if let Some(ref wb) = self.workbook {
            let mut this = wb.borrow_mut(py);
            let idx = self.resolve_index(&this)?;
            let ws = &mut this.inner.worksheets[idx];
            let target_row = if ws.cells.is_empty() {
                1
            } else {
                ws.dimensions().2 + 1
            };
            for (column, cv) in cells {
                ws.set_cell_value(target_row, column, cv);
            }
            Ok(())
        } else {
            Err(PyValueError::new_err(
                "Worksheet is not attached to a workbook",
            ))
        }
    }

    /// Insert `amount` blank rows before row `idx` (1-based; openpyxl semantics).
    #[pyo3(signature = (idx, amount=None))]
    fn insert_rows(&self, idx: u32, amount: Option<u32>, py: Python<'_>) -> PyResult<()> {
        self.with_sheet_mut(py, |ws| ws.insert_rows(idx, amount.unwrap_or(1)))
    }

    /// Insert `amount` blank columns before column `idx` (1-based).
    #[pyo3(signature = (idx, amount=None))]
    fn insert_cols(&self, idx: u32, amount: Option<u32>, py: Python<'_>) -> PyResult<()> {
        self.with_sheet_mut(py, |ws| ws.insert_columns(idx, amount.unwrap_or(1)))
    }

    /// Delete `amount` rows starting at row `idx` (1-based).
    #[pyo3(signature = (idx, amount=None))]
    fn delete_rows(&self, idx: u32, amount: Option<u32>, py: Python<'_>) -> PyResult<()> {
        self.with_sheet_mut(py, |ws| ws.delete_rows(idx, amount.unwrap_or(1)))
    }

    /// Delete `amount` columns starting at column `idx` (1-based).
    #[pyo3(signature = (idx, amount=None))]
    fn delete_cols(&self, idx: u32, amount: Option<u32>, py: Python<'_>) -> PyResult<()> {
        self.with_sheet_mut(py, |ws| ws.delete_columns(idx, amount.unwrap_or(1)))
    }

    /// Add a chart anchored at `anchor` (e.g. "E1"). It is written on save and
    /// opens in Excel with the given series, labels, title and legend.
    ///
    /// `chart_type` is one of bar, column, line, area, pie, doughnut, scatter.
    /// `series` is a value reference like "Sheet1!$B$1:$B$10", or a list whose
    /// items are such strings or dicts with keys values/name/categories/
    /// fill_color. `categories` supplies a default category (x-axis) reference
    /// for series that don't carry their own. `legend` is a position
    /// (r/l/t/b/tr) or None to hide it.
    #[pyo3(signature = (chart_type, series, anchor, title=None, categories=None, legend="r"))]
    #[allow(clippy::too_many_arguments)]
    fn add_chart(
        &self,
        chart_type: &str,
        series: &Bound<'_, PyAny>,
        anchor: &str,
        title: Option<&str>,
        categories: Option<&str>,
        legend: Option<&str>,
        py: Python<'_>,
    ) -> PyResult<()> {
        use rustypyxl_core::chart::{Chart, ChartAnchor, ChartLegend, ChartType};

        let ctype = match chart_type.to_ascii_lowercase().as_str() {
            "bar" => ChartType::Bar,
            "column" | "col" => ChartType::Column,
            "line" => ChartType::Line,
            "area" => ChartType::Area,
            "pie" => ChartType::Pie,
            "doughnut" => ChartType::Doughnut,
            "scatter" | "xy" => ChartType::Scatter,
            other => {
                return Err(PyValueError::new_err(format!(
                "unknown chart_type {other:?}; expected bar/column/line/area/pie/doughnut/scatter"
            )))
            }
        };

        let mut chart = match ctype {
            ChartType::Bar => Chart::bar(),
            ChartType::Column => Chart::column(),
            ChartType::Line => Chart::line(),
            ChartType::Area => Chart::area(),
            ChartType::Pie => Chart::pie(),
            ChartType::Scatter => Chart::scatter(),
            _ => Chart::new(ctype),
        };

        if let Some(t) = title {
            chart = chart.with_title(t);
        }
        chart = match legend {
            Some(pos) => chart.with_legend(ChartLegend::new().with_position(pos)),
            None => chart.with_legend(ChartLegend::new().with_visible(false)),
        };
        chart = chart.with_anchor(ChartAnchor::at(anchor));

        // series may be a single reference string, a single dict, or a list of
        // refs/dicts.
        if let Ok(single) = series.extract::<String>() {
            chart.add_series(build_series(&single, None, categories)?);
        } else if let Ok(list) = series.downcast::<PyList>() {
            for item in list.iter() {
                chart.add_series(parse_series_item(&item, categories)?);
            }
        } else if series.downcast::<PyDict>().is_ok() {
            chart.add_series(parse_series_item(series, categories)?);
        } else {
            return Err(PyValueError::new_err(
                "series must be a reference string, a dict, or a list of strings/dicts",
            ));
        }

        if chart.series.is_empty() {
            return Err(PyValueError::new_err("a chart needs at least one series"));
        }

        self.with_sheet_mut(py, |ws| ws.add_chart(chart))
    }

    /// Embed an image anchored at `anchor` (e.g. "B2"). It is written into the
    /// saved workbook and opens in Excel.
    ///
    /// `image` is a filesystem path or the raw image bytes; the format (PNG,
    /// JPEG, GIF, BMP, TIFF) is detected from the extension or magic bytes.
    /// Pass `to` for a two-cell anchor that resizes with the cells, or `width`
    /// and `height` (pixels) to set an explicit size.
    #[pyo3(signature = (image, anchor, to=None, width=None, height=None, name=None))]
    #[allow(clippy::too_many_arguments)]
    fn add_image(
        &self,
        image: &Bound<'_, PyAny>,
        anchor: &str,
        to: Option<&str>,
        width: Option<u32>,
        height: Option<u32>,
        name: Option<&str>,
        py: Python<'_>,
    ) -> PyResult<()> {
        use rustypyxl_core::image::{Image, ImageAnchor};

        let img_anchor = match to {
            Some(to_cell) => ImageAnchor::two_cell(anchor, to_cell),
            None => ImageAnchor::one_cell(anchor),
        };

        // A str/os.PathLike is a path; bytes is the raw image.
        let mut img = if let Ok(data) = image.extract::<Vec<u8>>() {
            Image::from_bytes(data, img_anchor)
                .ok_or_else(|| PyValueError::new_err("unrecognized image format"))?
        } else if let Ok(path) = image.extract::<std::path::PathBuf>() {
            Image::from_file(&path, img_anchor)
                .map_err(|e| PyValueError::new_err(format!("could not read image: {e}")))?
        } else {
            return Err(PyValueError::new_err("image must be a file path or bytes"));
        };

        if let (Some(w), Some(h)) = (width, height) {
            img = img.with_size_px(w, h);
        }
        if let Some(n) = name {
            img = img.with_name(n);
        }

        self.with_sheet_mut(py, |ws| ws.add_image(img))
    }

    /// Size a column to fit its content and return the width set (or None if the
    /// column is empty). `column` is 1-based. The width is an estimate from the
    /// displayed text length, not pixel-perfect.
    fn auto_fit_column(&self, column: u32, py: Python<'_>) -> PyResult<Option<f64>> {
        if let Some(ref wb) = self.workbook {
            let mut this = wb.borrow_mut(py);
            let idx = self.resolve_index(&this)?;
            return Ok(this.inner.worksheets[idx].auto_fit_column(column));
        }
        Err(PyValueError::new_err(
            "Worksheet is not attached to a workbook",
        ))
    }

    /// Auto-fit every column that has content.
    fn auto_fit_all(&self, py: Python<'_>) -> PyResult<()> {
        self.with_sheet_mut(py, |ws| ws.auto_fit_all())
    }

    /// Add an Excel table (ListObject) over a cell range. `name` is the table
    /// name, `ref` its range (e.g. "A1:C10"). `style` is a table style name
    /// like "TableStyleMedium9". `headers` names the columns (defaults to the
    /// values in the header row). The remaining flags toggle the table's
    /// display options.
    #[pyo3(signature = (name, r#ref, style=None, headers=None, totals_row=false, header_row=true, first_column=false, last_column=false, row_stripes=true, column_stripes=false, auto_filter=true))]
    #[allow(clippy::too_many_arguments)]
    fn add_table(
        &self,
        name: &str,
        r#ref: &str,
        style: Option<String>,
        headers: Option<Vec<String>>,
        totals_row: bool,
        header_row: bool,
        first_column: bool,
        last_column: bool,
        row_stripes: bool,
        column_stripes: bool,
        auto_filter: bool,
        py: Python<'_>,
    ) -> PyResult<()> {
        use rustypyxl_core::table::{Table, TableStyle};

        let id = self.with_sheet_ref(py, |ws| ws.tables.len() as u32 + 1)?;
        let mut table = match &headers {
            Some(h) => {
                let refs: Vec<&str> = h.iter().map(|s| s.as_str()).collect();
                Table::with_headers(id, name, r#ref, &refs)
            }
            None => Table::new(id, name, r#ref),
        };
        table.header_row = header_row;
        table.totals_row = totals_row;
        table.show_first_column = first_column;
        table.show_last_column = last_column;
        table.show_row_stripes = row_stripes;
        table.show_column_stripes = column_stripes;
        table.auto_filter = auto_filter;
        if let Some(s) = style {
            table.style = TableStyle::Custom(s);
        }
        self.with_sheet_mut(py, |ws| ws.add_table(table))
    }

    /// Add a conditional-formatting rule over a cell range. `rule` is a dict
    /// describing the rule; supported forms:
    ///   {"type":"cellIs","operator":"greaterThan","formula":"5","fill":"FF0000"}
    ///   {"type":"expression","formula":"$A1>0","font_color":"FF0000","bold":true}
    ///   {"type":"colorScale","preset":"red_yellow_green"}
    ///   {"type":"dataBar","color":"638EC6"}
    ///   {"type":"top","rank":10,"fill":"FFFF00"}
    ///   {"type":"containsText","text":"x","fill":"FF0000"}
    ///   {"type":"duplicateValues"} / "uniqueValues" / "aboveAverage" / "belowAverage"
    /// Format keys fill, font_color, and bold apply where a rule supports a
    /// differential format.
    fn add_conditional_formatting(
        &self,
        cells: &str,
        rule: &Bound<'_, PyDict>,
        py: Python<'_>,
    ) -> PyResult<()> {
        use rustypyxl_core::conditional::ConditionalFormatting;
        let built = build_conditional_rule(rule)?;
        let mut cf = ConditionalFormatting::new(cells);
        cf.add_rule(built);
        self.with_sheet_mut(py, |ws| ws.add_conditional_formatting(cf))
    }

    /// Add a data-validation rule over a cell range (e.g. "A1:A10"). `type` is
    /// one of whole, decimal, list, date, time, textLength, custom. `formula1`
    /// (and `formula2` for between/notBetween) supply the constraint -- for a
    /// list, `formula1` is like '"A,B,C"' or a range. `operator` is between,
    /// notBetween, equal, notEqual, greaterThan, lessThan, greaterThanOrEqual,
    /// or lessThanOrEqual.
    #[pyo3(signature = (cells, r#type, formula1=None, formula2=None, operator=None, allow_blank=true, show_error=true, error_title=None, error=None, show_input=true, prompt_title=None, prompt=None))]
    #[allow(clippy::too_many_arguments)]
    fn add_data_validation(
        &self,
        cells: &str,
        r#type: &str,
        formula1: Option<String>,
        formula2: Option<String>,
        operator: Option<String>,
        allow_blank: bool,
        show_error: bool,
        error_title: Option<String>,
        error: Option<String>,
        show_input: bool,
        prompt_title: Option<String>,
        prompt: Option<String>,
        py: Python<'_>,
    ) -> PyResult<()> {
        use rustypyxl_core::DataValidation;

        // Key the rule at the range's top-left cell; sqref carries the full range.
        let first = cells.split(':').next().unwrap_or(cells);
        let (row, col) =
            parse_coordinate(first).map_err(|e| PyValueError::new_err(e.to_string()))?;

        let dv = DataValidation {
            validation_type: r#type.to_string(),
            operator,
            formula1,
            formula2,
            error_style: None,
            allow_blank,
            show_error,
            error_title,
            error_message: error,
            show_input,
            prompt_title,
            prompt_message: prompt,
            sqref: Some(cells.to_string()),
        };
        self.with_sheet_mut(py, |ws| ws.add_data_validation(row, col, dv))
    }

    /// The data-validation rules on this sheet as a list of dicts with keys
    /// sqref, type, operator, formula1, and formula2.
    #[getter]
    fn data_validations(&self, py: Python<'_>) -> PyResult<PyObject> {
        use pyo3::types::{PyDict, PyList};
        let list = PyList::empty(py);
        self.with_sheet_ref(py, |ws| -> PyResult<()> {
            for ((row, col), dv) in &ws.data_validations {
                let d = PyDict::new(py);
                let sqref = dv
                    .sqref
                    .clone()
                    .unwrap_or_else(|| coordinate_from_row_col(*row, *col));
                d.set_item("sqref", sqref)?;
                d.set_item("type", &dv.validation_type)?;
                d.set_item("operator", dv.operator.clone())?;
                d.set_item("formula1", dv.formula1.clone())?;
                d.set_item("formula2", dv.formula2.clone())?;
                list.append(d)?;
            }
            Ok(())
        })??;
        Ok(list.into_any().unbind())
    }

    /// The tables on this sheet as a list of dicts with keys name, ref, and
    /// style.
    #[getter]
    fn tables(&self, py: Python<'_>) -> PyResult<PyObject> {
        use pyo3::types::{PyDict, PyList};
        let list = PyList::empty(py);
        self.with_sheet_ref(py, |ws| -> PyResult<()> {
            for t in &ws.tables {
                let d = PyDict::new(py);
                d.set_item("name", &t.name)?;
                d.set_item("ref", &t.range)?;
                d.set_item("style", t.style.style_name())?;
                list.append(d)?;
            }
            Ok(())
        })??;
        Ok(list.into_any().unbind())
    }

    /// Column dimensions, indexed by column letter:
    /// `ws.column_dimensions['A'].width = 20`.
    #[getter]
    fn column_dimensions(
        &self,
        py: Python<'_>,
    ) -> PyResult<crate::dimensions::PyColumnDimensions> {
        let wb = self
            .workbook
            .as_ref()
            .ok_or_else(|| PyValueError::new_err("Worksheet is not attached to a workbook"))?;
        Ok(crate::dimensions::PyColumnDimensions {
            workbook: wb.clone_ref(py),
            uid: self.uid,
        })
    }

    /// Row dimensions, indexed by row number: `ws.row_dimensions[1].height = 15`.
    #[getter]
    fn row_dimensions(&self, py: Python<'_>) -> PyResult<crate::dimensions::PyRowDimensions> {
        let wb = self
            .workbook
            .as_ref()
            .ok_or_else(|| PyValueError::new_err("Worksheet is not attached to a workbook"))?;
        Ok(crate::dimensions::PyRowDimensions {
            workbook: wb.clone_ref(py),
            uid: self.uid,
        })
    }

    /// Get the freeze-panes anchor cell, if any.
    #[getter]
    fn freeze_panes(&self, py: Python<'_>) -> PyResult<Option<String>> {
        if let Some(ref wb) = self.workbook {
            let this = wb.borrow(py);
            let idx = self.resolve_index(&this)?;
            return Ok(this.inner.worksheets[idx].freeze_panes.clone());
        }
        Ok(None)
    }

    /// Freeze panes at a cell (e.g. "B2"); pass None to unfreeze.
    #[setter]
    fn set_freeze_panes(&self, cell: Option<String>) -> PyResult<()> {
        Python::with_gil(|py| self.with_sheet_mut(py, move |ws| ws.set_freeze_panes(cell)))
    }

    fn __str__(&self, py: Python<'_>) -> String {
        format!("<Worksheet \"{}\">", self.title(py))
    }

    fn __repr__(&self, py: Python<'_>) -> String {
        self.__str__(py)
    }

    /// GC support: worksheet handles hold a workbook reference.
    fn __traverse__(&self, visit: pyo3::PyVisit<'_>) -> Result<(), pyo3::PyTraverseError> {
        if let Some(ref wb) = self.workbook {
            visit.call(wb)?;
        }
        Ok(())
    }

    fn __clear__(&mut self) {
        self.workbook = None;
    }
}

/// Lazy iterator over a worksheet range, yielding one tuple per row (or per
/// column for iter_cols) like openpyxl's generators. Resolves the sheet by
/// stable uid on every step, so concurrent sheet removal raises instead of
/// reading a neighbor.
#[pyclass(name = "CellRangeIterator")]
pub struct PyCellRangeIterator {
    workbook: Option<Py<PyWorkbook>>,
    sheet_uid: u64,
    min_row: u32,
    max_row: u32,
    min_col: u32,
    max_col: u32,
    values_only: bool,
    by_columns: bool,
    /// Next row (or column when by_columns) to yield.
    position: u32,
}

impl PyCellRangeIterator {
    /// Read a whole row (or column) of values in one pass.
    ///
    /// Resolves the sheet once rather than scanning the workbook's sheet list
    /// for every cell, and copies the values out before converting them, so no
    /// Python object is built while the workbook is borrowed.
    fn read_values(&self, coords: &[(u32, u32)], py: Python<'_>) -> PyResult<Vec<PyObject>> {
        let Some(ref wb) = self.workbook else {
            return Ok(coords.iter().map(|_| py.None()).collect());
        };

        let values: Vec<CellValue> = {
            let this = wb.borrow(py);
            let idx = this
                .inner
                .sheet_index_by_uid(self.sheet_uid)
                .ok_or_else(|| {
                    PyValueError::new_err("Worksheet no longer exists in this workbook")
                })?;
            let worksheet = &this.inner.worksheets[idx];
            coords
                .iter()
                .map(|&(row, col)| {
                    worksheet
                        .get_cell(row, col)
                        .map(|cell| cell.value.clone())
                        .unwrap_or(CellValue::Empty)
                })
                .collect()
        };

        Ok(values
            .iter()
            .map(|value| cell_value_to_python(value, py))
            .collect())
    }

    fn make_cell(&self, row: u32, col: u32, py: Python<'_>) -> PyResult<PyObject> {
        if let Some(ref wb) = self.workbook {
            Ok(Py::new(
                py,
                PyCell::connected(row, col, wb.clone_ref(py), self.sheet_uid),
            )?
            .into_any())
        } else {
            Ok(Py::new(py, PyCell::new(row, col))?.into_any())
        }
    }
}

#[pymethods]
impl PyCellRangeIterator {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<PyObject>> {
        use pyo3::types::PyTuple;

        let limit = if self.by_columns {
            self.max_col
        } else {
            self.max_row
        };
        if self.position > limit {
            return Ok(None);
        }
        let outer = self.position;
        self.position += 1;

        let coords: Vec<(u32, u32)> = if self.by_columns {
            (self.min_row..=self.max_row)
                .map(|row| (row, outer))
                .collect()
        } else {
            (self.min_col..=self.max_col)
                .map(|col| (outer, col))
                .collect()
        };

        let items: Vec<PyObject> = if self.values_only {
            self.read_values(&coords, py)?
        } else {
            coords
                .iter()
                .map(|&(row, col)| self.make_cell(row, col, py))
                .collect::<PyResult<_>>()?
        };
        Ok(Some(PyTuple::new(py, items)?.into_any().unbind()))
    }

    fn __traverse__(&self, visit: pyo3::PyVisit<'_>) -> Result<(), pyo3::PyTraverseError> {
        if let Some(ref wb) = self.workbook {
            visit.call(wb)?;
        }
        Ok(())
    }

    fn __clear__(&mut self) {
        self.workbook = None;
    }
}

/// Build a conditional-formatting rule from a Python dict describing it.
fn build_conditional_rule(
    dict: &Bound<'_, PyDict>,
) -> PyResult<rustypyxl_core::conditional::ConditionalRule> {
    use rustypyxl_core::conditional::{
        ColorScale, ConditionalColor, ConditionalFormat, ConditionalOperator, ConditionalRule,
        DataBar,
    };

    let get_str = |key: &str| -> PyResult<Option<String>> {
        match dict.get_item(key)? {
            Some(v) => Ok(Some(v.extract()?)),
            None => Ok(None),
        }
    };
    let rule_type = get_str("type")?
        .ok_or_else(|| PyValueError::new_err("conditional rule requires a 'type' key"))?;

    // The differential format shared by most rule kinds.
    let format = {
        let mut f = ConditionalFormat::new();
        let mut any = false;
        if let Some(c) = get_str("font_color")? {
            f = f.with_font_color(ConditionalColor::rgb(c));
            any = true;
        }
        if let Some(c) = get_str("fill")? {
            f = f.with_fill(ConditionalColor::rgb(c));
            any = true;
        }
        if let Some(b) = dict.get_item("bold")? {
            f = f.with_bold(b.extract()?);
            any = true;
        }
        if any {
            Some(f)
        } else {
            None
        }
    };

    let mut rule = match rule_type.as_str() {
        "cellIs" => {
            let op_str = get_str("operator")?
                .ok_or_else(|| PyValueError::new_err("cellIs requires 'operator'"))?;
            let op = ConditionalOperator::from_xml(&op_str)
                .ok_or_else(|| PyValueError::new_err(format!("unknown operator {op_str:?}")))?;
            let value = get_str("formula")?.unwrap_or_default();
            ConditionalRule::cell_is(op, &value)
        }
        "expression" | "formula" => {
            let f = get_str("formula")?
                .ok_or_else(|| PyValueError::new_err("expression requires 'formula'"))?;
            ConditionalRule::formula(&f)
        }
        "colorScale" => {
            let scale = match get_str("preset")?.as_deref() {
                Some("green_yellow_red") => ColorScale::green_yellow_red(),
                _ => ColorScale::red_yellow_green(),
            };
            ConditionalRule::with_color_scale(scale)
        }
        "dataBar" => {
            let mut bar = DataBar::new();
            if let Some(c) = get_str("color")? {
                bar = bar.with_color(ConditionalColor::rgb(c));
            }
            ConditionalRule::with_data_bar(bar)
        }
        "top" => ConditionalRule::top(rank(dict)?),
        "bottom" => ConditionalRule::bottom(rank(dict)?),
        "containsText" => ConditionalRule::contains_text(&text_of(dict)?),
        "beginsWith" => ConditionalRule::begins_with(&text_of(dict)?),
        "endsWith" => ConditionalRule::ends_with(&text_of(dict)?),
        "duplicateValues" => ConditionalRule::duplicate_values(),
        "uniqueValues" => ConditionalRule::unique_values(),
        "aboveAverage" => ConditionalRule::above_average(),
        "belowAverage" => ConditionalRule::below_average(),
        other => {
            return Err(PyValueError::new_err(format!(
                "unsupported conditional rule type {other:?}"
            )))
        }
    };
    if let Some(f) = format {
        rule = rule.with_format(f);
    }
    Ok(rule)
}

/// Read the integer "rank" key of a top/bottom rule dict (default 10).
fn rank(dict: &Bound<'_, PyDict>) -> PyResult<u32> {
    match dict.get_item("rank")? {
        Some(v) => v.extract(),
        None => Ok(10),
    }
}

/// Read the required "text" key of a text rule dict.
fn text_of(dict: &Bound<'_, PyDict>) -> PyResult<String> {
    dict.get_item("text")?
        .ok_or_else(|| PyValueError::new_err("text rule requires a 'text' key"))?
        .extract()
}

/// Build a chart series from a values reference, an optional name, and an
/// optional categories reference (falling back to the chart-wide default).
fn build_series(
    values: &str,
    name: Option<&str>,
    default_categories: Option<&str>,
) -> PyResult<rustypyxl_core::chart::ChartSeries> {
    use rustypyxl_core::chart::ChartSeries;
    let mut s = ChartSeries::new(values);
    if let Some(n) = name {
        s = s.with_name(n);
    }
    if let Some(c) = default_categories {
        s = s.with_categories(c);
    }
    Ok(s)
}

/// Parse one item of the `series` argument: either a reference string or a dict
/// with keys values (required), name, categories, fill_color.
fn parse_series_item(
    item: &Bound<'_, PyAny>,
    default_categories: Option<&str>,
) -> PyResult<rustypyxl_core::chart::ChartSeries> {
    if let Ok(values) = item.extract::<String>() {
        return build_series(&values, None, default_categories);
    }

    let dict = item
        .downcast::<PyDict>()
        .map_err(|_| PyValueError::new_err("each series must be a reference string or a dict"))?;

    let values: String = dict
        .get_item("values")?
        .ok_or_else(|| PyValueError::new_err("series dict requires a 'values' key"))?
        .extract()?;
    let name: Option<String> = match dict.get_item("name")? {
        Some(v) => Some(v.extract()?),
        None => None,
    };
    let categories: Option<String> = match dict.get_item("categories")? {
        Some(v) => Some(v.extract()?),
        None => default_categories.map(|s| s.to_string()),
    };
    let fill_color: Option<String> = match dict.get_item("fill_color")? {
        Some(v) => Some(v.extract()?),
        None => None,
    };

    let mut s = rustypyxl_core::chart::ChartSeries::new(values);
    if let Some(n) = name {
        s = s.with_name(n);
    }
    if let Some(c) = categories {
        s = s.with_categories(c);
    }
    if let Some(fc) = fill_color {
        s = s.with_fill_color(fc);
    }
    Ok(s)
}

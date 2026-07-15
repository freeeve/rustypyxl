//! Worksheet representation and cell operations.

use crate::autofilter::AutoFilter;
use crate::cell::{CellValue, InternedString};
use crate::conditional::ConditionalFormatting;
use crate::pagesetup::PageSetup;
use crate::style::CellStyle;
use crate::table::Table;
#[cfg(feature = "fast-hash")]
use hashbrown::HashMap;
#[cfg(not(feature = "fast-hash"))]
use std::collections::HashMap;
use std::sync::Arc;

#[cfg(feature = "fast-hash")]
pub type CellMap = hashbrown::HashMap<u64, CellData, ahash::RandomState>;
#[cfg(not(feature = "fast-hash"))]
pub type CellMap = std::collections::HashMap<u64, CellData>;

#[inline]
pub(crate) fn cell_key(row: u32, column: u32) -> u64 {
    ((row as u64) << 32) | (column as u64)
}

#[inline]
pub(crate) fn decode_cell_key(key: u64) -> (u32, u32) {
    ((key >> 32) as u32, key as u32)
}

/// Data associated with a single cell.
#[derive(Clone, Debug, Default)]
pub struct CellData {
    /// The cell's value.
    pub value: CellValue,
    /// Cell style (font, alignment, etc.) - Arc for cheap cloning.
    pub style: Option<Arc<CellStyle>>,
    /// Style index for writing (preserves original style during roundtrip).
    pub style_index: Option<u32>,
    /// Number format string. Interned: the same format is shared by every cell
    /// in a column, and it is cloned per cell when styles are resolved on save.
    pub number_format: Option<InternedString>,
    /// Data type (s=string, n=number, b=boolean, d=date). Always one of a fixed
    /// set of codes, so it borrows rather than allocating per cell.
    pub data_type: Option<&'static str>,
    /// Hyperlink URL.
    pub hyperlink: Option<String>,
    /// Cell comment text.
    pub comment: Option<String>,
    /// Last calculated result of a formula cell, as the raw `<v>` text.
    /// Written back on save so viewers that don't recalculate show a value;
    /// `data_type` carries the matching `t` attribute (str/b/e or numeric).
    pub cached_formula_value: Option<String>,
    /// Per-run formatting when the cell's string is rich text. When present,
    /// `value` holds the concatenated plain text and the cell is written as a
    /// rich string; when None, the string is a plain `<t>`.
    pub rich_text: Option<crate::rich_text::RichText>,
}

impl CellData {
    /// Create empty cell data.
    pub fn new() -> Self {
        Self::default()
    }

    /// Create cell data with a value.
    pub fn with_value(value: CellValue) -> Self {
        CellData {
            value,
            ..Default::default()
        }
    }
}

/// Sheet visibility as stored on the workbook.xml `<sheet state>` attribute.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum SheetVisibility {
    /// Shown in the tab bar (default).
    #[default]
    Visible,
    /// Hidden, but can be unhidden from the Excel UI.
    Hidden,
    /// Hidden and only unhideable through VBA / the object model.
    VeryHidden,
}

impl SheetVisibility {
    /// Attribute value for workbook.xml.
    pub fn as_str(&self) -> &'static str {
        match self {
            SheetVisibility::Visible => "visible",
            SheetVisibility::Hidden => "hidden",
            SheetVisibility::VeryHidden => "veryHidden",
        }
    }

    /// Parse the workbook.xml attribute value; unknown values load as Visible.
    pub fn from_attr(value: &str) -> Self {
        match value {
            "hidden" => SheetVisibility::Hidden,
            "veryHidden" => SheetVisibility::VeryHidden,
            _ => SheetVisibility::Visible,
        }
    }
}

/// Data validation rule for a cell.
#[derive(Clone, Debug)]
pub struct DataValidation {
    /// Type: whole, decimal, list, date, time, textLength, custom.
    pub validation_type: String,
    /// Comparison operator: between, notBetween, equal, notEqual, greaterThan,
    /// lessThan, greaterThanOrEqual, lessThanOrEqual. None means Excel's
    /// default, "between".
    pub operator: Option<String>,
    /// First formula/value constraint.
    pub formula1: Option<String>,
    /// Second formula/value constraint (for between/notBetween).
    pub formula2: Option<String>,
    /// Severity of the error dialog: stop, warning, or information.
    pub error_style: Option<String>,
    /// Allow blank values.
    pub allow_blank: bool,
    /// Show error message on invalid input.
    pub show_error: bool,
    /// Error dialog title.
    pub error_title: Option<String>,
    /// Error message text.
    pub error_message: Option<String>,
    /// Show input message when cell is selected.
    pub show_input: bool,
    /// Input prompt title.
    pub prompt_title: Option<String>,
    /// Input prompt message.
    pub prompt_message: Option<String>,
    /// Full sqref the rule applies to (may span multiple cells/ranges).
    /// When None, the rule applies to the single cell it is keyed under.
    pub sqref: Option<String>,
}

impl Default for DataValidation {
    fn default() -> Self {
        DataValidation {
            validation_type: "whole".to_string(),
            operator: None,
            formula1: None,
            formula2: None,
            error_style: None,
            allow_blank: true,
            show_error: true,
            error_title: None,
            error_message: None,
            show_input: true,
            prompt_title: None,
            prompt_message: None,
            sqref: None,
        }
    }
}

/// Worksheet protection settings.
#[derive(Clone, Debug, Default)]
pub struct WorksheetProtection {
    /// Sheet protection enabled.
    pub sheet: bool,
    /// Plaintext password; hashed with the legacy Excel verifier on save.
    pub password: Option<String>,
    /// Pre-hashed password verifier loaded from an existing file.
    /// Takes precedence over `password` on save so a loaded hash is never re-hashed.
    pub password_hash: Option<String>,
    /// Allow selecting locked cells.
    pub select_locked_cells: bool,
    /// Allow selecting unlocked cells.
    pub select_unlocked_cells: bool,
    /// Allow formatting cells.
    pub format_cells: bool,
    /// Allow formatting columns.
    pub format_columns: bool,
    /// Allow formatting rows.
    pub format_rows: bool,
    /// Allow inserting columns.
    pub insert_columns: bool,
    /// Allow inserting rows.
    pub insert_rows: bool,
    /// Allow inserting hyperlinks.
    pub insert_hyperlinks: bool,
    /// Allow deleting columns.
    pub delete_columns: bool,
    /// Allow deleting rows.
    pub delete_rows: bool,
    /// Allow sorting.
    pub sort: bool,
    /// Allow using autofilter.
    pub auto_filter: bool,
    /// Allow editing pivot tables.
    pub pivot_tables: bool,
    /// Allow editing objects.
    pub objects: bool,
    /// Allow editing scenarios.
    pub scenarios: bool,
}

/// Represents a worksheet in an Excel workbook.
#[derive(Clone, Debug)]
pub struct Worksheet {
    /// Worksheet title/name.
    pub title: String,
    /// Cell data indexed by packed (row, column) key - both 1-indexed.
    pub cells: CellMap,
    /// Merged cell ranges as (start_coord, end_coord) strings.
    pub merged_cells: Vec<(String, String)>,
    /// Column widths indexed by column number.
    pub column_dimensions: HashMap<u32, f64>,
    /// Row heights indexed by row number.
    pub row_dimensions: HashMap<u32, f64>,
    /// Data validations indexed by (row, column).
    pub data_validations: HashMap<(u32, u32), DataValidation>,
    /// Sheet protection settings.
    pub protection: Option<WorksheetProtection>,
    /// Maximum row with data (for optimization).
    pub max_row: u32,
    /// Maximum column with data (for optimization).
    pub max_column: u32,
    /// AutoFilter configuration.
    pub auto_filter: Option<AutoFilter>,
    /// Conditional formatting rules.
    pub conditional_formatting: Vec<ConditionalFormatting>,
    /// Excel Tables (ListObjects).
    pub tables: Vec<Table>,
    /// Page setup and print settings.
    pub page_setup: Option<PageSetup>,
    /// Freeze panes anchor cell (e.g. "B2"); rows above and columns left of it stay frozen.
    pub freeze_panes: Option<String>,
    /// Sheet visibility (visible / hidden / veryHidden).
    pub visibility: SheetVisibility,
    /// Stable identity within the owning workbook. Assigned by the workbook
    /// (never reused), so handles survive sheet removal, reordering, and
    /// renames. 0 means the worksheet is not attached to a workbook.
    pub uid: u64,
}

impl Worksheet {
    /// Create a new worksheet with the given title.
    pub fn new<S: Into<String>>(title: S) -> Self {
        Worksheet {
            title: title.into(),
            cells: CellMap::default(),
            merged_cells: Vec::new(),
            column_dimensions: HashMap::new(),
            row_dimensions: HashMap::new(),
            data_validations: HashMap::new(),
            protection: None,
            max_row: 0,
            max_column: 0,
            auto_filter: None,
            conditional_formatting: Vec::new(),
            tables: Vec::new(),
            page_setup: None,
            freeze_panes: None,
            visibility: SheetVisibility::default(),
            uid: 0,
        }
    }

    /// Freeze panes at the given anchor cell (e.g. "B2"). Pass `None` to unfreeze.
    pub fn set_freeze_panes(&mut self, cell: Option<String>) {
        self.freeze_panes = cell;
    }

    /// Set an AutoFilter for this worksheet.
    pub fn set_auto_filter(&mut self, auto_filter: AutoFilter) {
        self.auto_filter = Some(auto_filter);
    }

    /// Add a conditional formatting rule.
    pub fn add_conditional_formatting(&mut self, cf: ConditionalFormatting) {
        self.conditional_formatting.push(cf);
    }

    /// Add an Excel Table.
    pub fn add_table(&mut self, table: Table) {
        self.tables.push(table);
    }

    /// Set page setup.
    pub fn set_page_setup(&mut self, page_setup: PageSetup) {
        self.page_setup = Some(page_setup);
    }

    /// Get the worksheet title.
    pub fn title(&self) -> &str {
        &self.title
    }

    /// Set the worksheet title.
    pub fn set_title<S: Into<String>>(&mut self, title: S) {
        self.title = title.into();
    }

    /// Get cell data at the specified row and column (1-indexed).
    pub fn get_cell(&self, row: u32, column: u32) -> Option<&CellData> {
        self.cells.get(&cell_key(row, column))
    }

    /// Get mutable cell data at the specified row and column (1-indexed).
    pub fn get_cell_mut(&mut self, row: u32, column: u32) -> Option<&mut CellData> {
        self.cells.get_mut(&cell_key(row, column))
    }

    /// Get the cell value at the specified position.
    pub fn get_cell_value(&self, row: u32, column: u32) -> Option<&CellValue> {
        self.cells.get(&cell_key(row, column)).map(|cd| &cd.value)
    }

    /// Set a cell value at the specified row and column (1-indexed).
    pub fn set_cell_value<V: Into<CellValue>>(&mut self, row: u32, column: u32, value: V) {
        let cell_data = self.cells.entry(cell_key(row, column)).or_default();
        cell_data.value = value.into();
        self.update_dimensions(row, column);
    }

    /// Set a rich-text value on a cell. The cell's plain value becomes the
    /// concatenated run text and the runs are preserved (and written as a rich
    /// string on save).
    pub fn set_cell_rich_text(&mut self, row: u32, column: u32, rich: crate::rich_text::RichText) {
        let plain = rich.plain();
        let cell = self.cells.entry(cell_key(row, column)).or_default();
        cell.value = CellValue::String(Arc::from(plain.as_str()));
        cell.data_type = Some("s");
        cell.rich_text = Some(rich);
        self.update_dimensions(row, column);
    }

    /// Get a mutable reference to a cell, creating it if it doesn't exist.
    pub fn get_or_create_cell_mut(&mut self, row: u32, column: u32) -> &mut CellData {
        self.update_dimensions(row, column);
        self.cells.entry(cell_key(row, column)).or_default()
    }

    /// Set complete cell data at the specified position.
    pub fn set_cell_data(&mut self, row: u32, column: u32, data: CellData) {
        self.cells.insert(cell_key(row, column), data);
        self.update_dimensions(row, column);
    }

    /// Set a formula in a cell.
    pub fn set_cell_formula<S: Into<String>>(&mut self, row: u32, column: u32, formula: S) {
        let cell_data = self.cells.entry(cell_key(row, column)).or_default();
        cell_data.value = CellValue::Formula(formula.into());
        self.update_dimensions(row, column);
    }

    /// Set a cell's hyperlink.
    pub fn set_cell_hyperlink(&mut self, row: u32, column: u32, url: String) {
        let cell_data = self.cells.entry(cell_key(row, column)).or_default();
        cell_data.hyperlink = Some(url);
        self.update_dimensions(row, column);
    }

    /// Set a cell's comment.
    pub fn set_cell_comment(&mut self, row: u32, column: u32, comment: String) {
        let cell_data = self.cells.entry(cell_key(row, column)).or_default();
        cell_data.comment = Some(comment);
        self.update_dimensions(row, column);
    }

    /// Set a cell's style.
    pub fn set_cell_style(&mut self, row: u32, column: u32, style: CellStyle) {
        let cell_data = self.cells.entry(cell_key(row, column)).or_default();
        cell_data.style = Some(Arc::new(style));
        // Invalidate any loaded xf index so the new style is re-resolved on save
        cell_data.style_index = None;
        self.update_dimensions(row, column);
    }

    /// Set a cell's font, merging with any existing style on the cell.
    pub fn set_cell_font(&mut self, row: u32, column: u32, font: crate::style::Font) {
        let cell_data = self.cells.entry(cell_key(row, column)).or_default();
        let mut style = cell_data
            .style
            .as_deref()
            .cloned()
            .unwrap_or_else(CellStyle::new);
        style.font = Some(font);
        cell_data.style = Some(Arc::new(style));
        cell_data.style_index = None;
        self.update_dimensions(row, column);
    }

    /// Set a cell's alignment, merging with any existing style on the cell.
    pub fn set_cell_alignment(
        &mut self,
        row: u32,
        column: u32,
        alignment: crate::style::Alignment,
    ) {
        let cell_data = self.cells.entry(cell_key(row, column)).or_default();
        let mut style = cell_data
            .style
            .as_deref()
            .cloned()
            .unwrap_or_else(CellStyle::new);
        style.alignment = Some(alignment);
        cell_data.style = Some(Arc::new(style));
        cell_data.style_index = None;
        self.update_dimensions(row, column);
    }

    /// Set a cell's number format.
    pub fn set_cell_number_format<S: AsRef<str>>(&mut self, row: u32, column: u32, format: S) {
        let cell_data = self.cells.entry(cell_key(row, column)).or_default();
        cell_data.number_format = Some(Arc::from(format.as_ref()));
        // Invalidate any loaded xf index so the format is re-resolved on save
        cell_data.style_index = None;
        self.update_dimensions(row, column);
    }

    /// Add a merged cell range.
    pub fn add_merged_cell<S: Into<String>>(&mut self, start: S, end: S) {
        self.merged_cells.push((start.into(), end.into()));
    }

    /// Merge cells in a range (e.g., "A1:B2").
    pub fn merge_cells(&mut self, range: &str) {
        if let Some(colon_pos) = range.find(':') {
            let start = range[..colon_pos].to_string();
            let end = range[colon_pos + 1..].to_string();
            self.merged_cells.push((start, end));
        }
    }

    /// Unmerge cells in a range.
    pub fn unmerge_cells(&mut self, range: &str) {
        if let Some(colon_pos) = range.find(':') {
            let start = range[..colon_pos].to_string();
            let end = range[colon_pos + 1..].to_string();
            self.merged_cells
                .retain(|(s, e)| !(s == &start && e == &end));
        }
    }

    /// Set column width.
    pub fn set_column_width(&mut self, column: u32, width: f64) {
        self.column_dimensions.insert(column, width);
    }

    /// Get column width.
    pub fn get_column_width(&self, column: u32) -> Option<f64> {
        self.column_dimensions.get(&column).copied()
    }

    /// Set row height.
    pub fn set_row_height(&mut self, row: u32, height: f64) {
        self.row_dimensions.insert(row, height);
    }

    /// Get row height.
    pub fn get_row_height(&self, row: u32) -> Option<f64> {
        self.row_dimensions.get(&row).copied()
    }

    /// Add data validation to a cell.
    pub fn add_data_validation(&mut self, row: u32, column: u32, validation: DataValidation) {
        self.data_validations.insert((row, column), validation);
    }

    /// Get data validation for a cell.
    pub fn get_data_validation(&self, row: u32, column: u32) -> Option<&DataValidation> {
        self.data_validations.get(&(row, column))
    }

    /// Enable sheet protection.
    pub fn enable_protection(&mut self, password: Option<String>) {
        self.protection = Some(WorksheetProtection {
            sheet: true,
            password,
            ..Default::default()
        });
    }

    /// Disable sheet protection.
    pub fn disable_protection(&mut self) {
        self.protection = None;
    }

    /// Check if sheet is protected.
    pub fn is_protected(&self) -> bool {
        self.protection.as_ref().is_some_and(|p| p.sheet)
    }

    /// Get the maximum row number with data.
    pub fn max_row(&self) -> u32 {
        self.max_row
    }

    /// Get the maximum column number with data.
    pub fn max_column(&self) -> u32 {
        self.max_column
    }

    /// Get dimensions as (min_row, min_col, max_row, max_col).
    pub fn dimensions(&self) -> (u32, u32, u32, u32) {
        if self.cells.is_empty() {
            return (1, 1, 1, 1);
        }

        let mut min_row = u32::MAX;
        let mut min_col = u32::MAX;
        let mut max_row = 0;
        let mut max_col = 0;

        for &key in self.cells.keys() {
            let (row, col) = decode_cell_key(key);
            min_row = min_row.min(row);
            min_col = min_col.min(col);
            max_row = max_row.max(row);
            max_col = max_col.max(col);
        }

        (min_row, min_col, max_row, max_col)
    }

    /// Iterate over all cells in row-major order.
    pub fn iter_cells(&self) -> impl Iterator<Item = ((u32, u32), &CellData)> {
        let mut cells: Vec<_> = self.cells.iter().map(|(k, v)| (*k, v)).collect();
        cells.sort_by_key(|(k, _)| decode_cell_key(*k));
        cells.into_iter().map(|(k, v)| (decode_cell_key(k), v))
    }

    /// Iterate over the populated cells of one row, in column order.
    ///
    /// Probes the row's columns by key rather than scanning the whole cell map:
    /// filtering and sorting every cell to serve one row made row-by-row
    /// iteration O(rows * total_cells).
    pub fn iter_row(&self, row: u32) -> impl Iterator<Item = (u32, &CellData)> + '_ {
        (1..=self.max_column).filter_map(move |col| {
            self.cells
                .get(&cell_key(row, col))
                .map(|cell_data| (col, cell_data))
        })
    }

    /// Update max_row and max_column.
    fn update_dimensions(&mut self, row: u32, column: u32) {
        self.max_row = self.max_row.max(row);
        self.max_column = self.max_column.max(column);
    }

    /// Insert `amount` blank rows before row `idx` (1-based). Cells at or below
    /// `idx` shift down; merged ranges, data validations, conditional
    /// formatting, tables, and the autofilter/freeze anchors move with them.
    /// Formula text is left unchanged, matching openpyxl.
    pub fn insert_rows(&mut self, idx: u32, amount: u32) {
        if amount > 0 && idx >= 1 {
            self.apply_shift(Shift::Insert { at: idx, amount }, true);
        }
    }

    /// Delete `amount` rows starting at row `idx` (1-based).
    pub fn delete_rows(&mut self, idx: u32, amount: u32) {
        if amount > 0 && idx >= 1 {
            self.apply_shift(Shift::Delete { at: idx, amount }, true);
        }
    }

    /// Insert `amount` blank columns before column `idx` (1-based).
    pub fn insert_columns(&mut self, idx: u32, amount: u32) {
        if amount > 0 && idx >= 1 {
            self.apply_shift(Shift::Insert { at: idx, amount }, false);
        }
    }

    /// Delete `amount` columns starting at column `idx` (1-based).
    pub fn delete_columns(&mut self, idx: u32, amount: u32) {
        if amount > 0 && idx >= 1 {
            self.apply_shift(Shift::Delete { at: idx, amount }, false);
        }
    }

    /// Apply a row or column insert/delete to every position-bearing part of
    /// the sheet. `is_row` selects the axis.
    fn apply_shift(&mut self, shift: Shift, is_row: bool) {
        let map_pos = |row: u32, col: u32| -> Option<(u32, u32)> {
            if is_row {
                shift.map(row).map(|r| (r, col))
            } else {
                shift.map(col).map(|c| (row, c))
            }
        };

        // Cells (with their per-cell styles/hyperlinks/comments): rebuild the
        // map with shifted keys, dropping any cell in a deleted band.
        let mut new_cells = CellMap::default();
        new_cells.reserve(self.cells.len());
        for (key, data) in self.cells.drain() {
            let (row, col) = decode_cell_key(key);
            if let Some((r, c)) = map_pos(row, col) {
                new_cells.insert(cell_key(r, c), data);
            }
        }
        self.cells = new_cells;

        // Row heights / column widths: shift keys on the affected axis only.
        if is_row {
            self.row_dimensions = shift_dim_keys(&self.row_dimensions, shift);
        } else {
            self.column_dimensions = shift_dim_keys(&self.column_dimensions, shift);
        }

        // Merged ranges: move/grow/shrink; drop if collapsed to nothing or to a
        // single cell (no longer a merge).
        self.merged_cells
            .retain_mut(|(s, e)| match shift_merge(s, e, shift, is_row) {
                Some((ns, ne)) => {
                    *s = ns;
                    *e = ne;
                    true
                }
                None => false,
            });

        // Data validations: shift the keying cell and each rule's sqref.
        let mut new_dv = HashMap::new();
        for ((row, col), mut dv) in std::mem::take(&mut self.data_validations) {
            if let Some(pos) = map_pos(row, col) {
                if let Some(ref sq) = dv.sqref {
                    dv.sqref = shift_sqref(sq, shift, is_row);
                }
                new_dv.insert(pos, dv);
            }
        }
        self.data_validations = new_dv;

        // Range-bearing features.
        self.conditional_formatting.retain_mut(|cf| {
            match shift_range_str(&cf.range, shift, is_row) {
                Some(r) => {
                    cf.range = r;
                    true
                }
                None => false,
            }
        });
        self.tables
            .retain_mut(|t| match shift_range_str(&t.range, shift, is_row) {
                Some(r) => {
                    t.range = r;
                    true
                }
                None => false,
            });
        if let Some(af) = self.auto_filter.as_mut() {
            if let Some(r) = shift_range_str(&af.range, shift, is_row) {
                af.range = r;
            }
        }
        if let Some(anchor) = self.freeze_panes.take() {
            self.freeze_panes = shift_coord_str(&anchor, shift, is_row);
        }

        self.recompute_dimensions();
    }

    /// Recompute max_row/max_column by scanning the (already shifted) cell map.
    fn recompute_dimensions(&mut self) {
        let (mut max_row, mut max_col) = (0, 0);
        for &key in self.cells.keys() {
            let (r, c) = decode_cell_key(key);
            max_row = max_row.max(r);
            max_col = max_col.max(c);
        }
        self.max_row = max_row;
        self.max_column = max_col;
    }
}

/// A row or column insert/delete on one axis; positions are 1-based.
#[derive(Clone, Copy)]
enum Shift {
    Insert { at: u32, amount: u32 },
    Delete { at: u32, amount: u32 },
}

impl Shift {
    /// New position of a single cell/dimension, or None if it was deleted.
    fn map(self, p: u32) -> Option<u32> {
        match self {
            Shift::Insert { at, amount } => Some(if p >= at { p + amount } else { p }),
            Shift::Delete { at, amount } => {
                if p < at {
                    Some(p)
                } else if p < at + amount {
                    None
                } else {
                    Some(p - amount)
                }
            }
        }
    }

    /// New start endpoint of a range; a delete clamps up to the first surviving
    /// position so a range straddling the gap shrinks rather than vanishes.
    fn map_start(self, a: u32) -> u32 {
        match self {
            Shift::Insert { at, amount } => {
                if a >= at {
                    a + amount
                } else {
                    a
                }
            }
            Shift::Delete { at, amount } => {
                if a < at {
                    a
                } else if a < at + amount {
                    at
                } else {
                    a - amount
                }
            }
        }
    }

    /// New end endpoint of a range; a delete clamps down to the last surviving
    /// position before the gap.
    fn map_end(self, b: u32) -> u32 {
        match self {
            Shift::Insert { at, amount } => {
                if b >= at {
                    b + amount
                } else {
                    b
                }
            }
            Shift::Delete { at, amount } => {
                if b < at {
                    b
                } else if b < at + amount {
                    at.saturating_sub(1)
                } else {
                    b - amount
                }
            }
        }
    }
}

/// Shift an `"A1"` coordinate on one axis; None if the cell was deleted.
fn shift_coord_str(coord: &str, shift: Shift, is_row: bool) -> Option<String> {
    let (row, col) = crate::utils::parse_coordinate(coord).ok()?;
    let (nr, nc) = if is_row {
        (shift.map(row)?, col)
    } else {
        (row, shift.map(col)?)
    };
    Some(crate::utils::coordinate_from_row_col(nr, nc))
}

/// Shift a merged range's `(start, end)` coordinates; None if it collapses to
/// nothing or to a single cell (no longer a merge).
fn shift_merge(s: &str, e: &str, shift: Shift, is_row: bool) -> Option<(String, String)> {
    let (r1, c1) = crate::utils::parse_coordinate(s).ok()?;
    let (r2, c2) = crate::utils::parse_coordinate(e).ok()?;
    let (nr1, nc1, nr2, nc2) = if is_row {
        (shift.map_start(r1), c1, shift.map_end(r2), c2)
    } else {
        (r1, shift.map_start(c1), r2, shift.map_end(c2))
    };
    if nr1 > nr2 || nc1 > nc2 || (nr1 == nr2 && nc1 == nc2) {
        return None;
    }
    Some((
        crate::utils::coordinate_from_row_col(nr1, nc1),
        crate::utils::coordinate_from_row_col(nr2, nc2),
    ))
}

/// Shift a range string `"A1:B2"` (or a bare `"A1"`) on one axis; None if the
/// range was entirely deleted. A range that collapses to one cell is returned
/// as a bare coordinate.
fn shift_range_str(range: &str, shift: Shift, is_row: bool) -> Option<String> {
    let (s, e) = match range.split_once(':') {
        Some((a, b)) => (a.trim(), b.trim()),
        None => (range.trim(), range.trim()),
    };
    let (r1, c1) = crate::utils::parse_coordinate(s).ok()?;
    let (r2, c2) = crate::utils::parse_coordinate(e).ok()?;
    let (nr1, nc1, nr2, nc2) = if is_row {
        (shift.map_start(r1), c1, shift.map_end(r2), c2)
    } else {
        (r1, shift.map_start(c1), r2, shift.map_end(c2))
    };
    if nr1 > nr2 || nc1 > nc2 {
        return None;
    }
    let start = crate::utils::coordinate_from_row_col(nr1, nc1);
    if (nr1, nc1) == (nr2, nc2) {
        Some(start)
    } else {
        Some(format!(
            "{}:{}",
            start,
            crate::utils::coordinate_from_row_col(nr2, nc2)
        ))
    }
}

/// Shift a whitespace-separated multi-range sqref; None if every range was
/// deleted.
fn shift_sqref(sqref: &str, shift: Shift, is_row: bool) -> Option<String> {
    let parts: Vec<String> = sqref
        .split_whitespace()
        .filter_map(|r| shift_range_str(r, shift, is_row))
        .collect();
    if parts.is_empty() {
        None
    } else {
        Some(parts.join(" "))
    }
}

/// Shift the keys of a row/column dimension map, dropping deleted lines.
fn shift_dim_keys(dims: &HashMap<u32, f64>, shift: Shift) -> HashMap<u32, f64> {
    let mut out = HashMap::with_capacity(dims.len());
    for (&k, &v) in dims {
        if let Some(nk) = shift.map(k) {
            out.insert(nk, v);
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_worksheet_new() {
        let ws = Worksheet::new("Sheet1");
        assert_eq!(ws.title(), "Sheet1");
        assert!(ws.cells.is_empty());
    }

    /// iter_row yields only the populated cells of the requested row, in
    /// column order, and skips gaps.
    #[test]
    fn test_iter_row_returns_populated_cells_in_column_order() {
        let mut ws = Worksheet::new("S");
        ws.set_cell_value(2, 3, CellValue::Number(3.0));
        ws.set_cell_value(2, 1, CellValue::Number(1.0));
        ws.set_cell_value(1, 2, CellValue::Number(99.0)); // different row
        ws.set_cell_value(2, 6, CellValue::Number(6.0));

        let row: Vec<(u32, f64)> = ws
            .iter_row(2)
            .map(|(col, cell)| match cell.value {
                CellValue::Number(n) => (col, n),
                _ => panic!("expected a number"),
            })
            .collect();

        assert_eq!(row, vec![(1, 1.0), (3, 3.0), (6, 6.0)]);
        assert_eq!(ws.iter_row(3).count(), 0, "empty row yields nothing");
    }

    #[test]
    fn test_set_cell_value() {
        let mut ws = Worksheet::new("Sheet1");
        ws.set_cell_value(1, 1, "Hello");

        let val = ws.get_cell_value(1, 1);
        assert!(matches!(val, Some(CellValue::String(s)) if s.as_ref() == "Hello"));
        assert_eq!(ws.max_row(), 1);
        assert_eq!(ws.max_column(), 1);
    }

    #[test]
    fn test_set_cell_formula() {
        let mut ws = Worksheet::new("Sheet1");
        ws.set_cell_formula(1, 1, "SUM(A1:A10)");

        let val = ws.get_cell_value(1, 1);
        assert!(matches!(val, Some(CellValue::Formula(f)) if f == "SUM(A1:A10)"));
    }

    #[test]
    fn test_merged_cells() {
        let mut ws = Worksheet::new("Sheet1");
        ws.merge_cells("A1:B2");
        assert_eq!(ws.merged_cells.len(), 1);

        ws.unmerge_cells("A1:B2");
        assert!(ws.merged_cells.is_empty());
    }

    #[test]
    fn test_column_dimensions() {
        let mut ws = Worksheet::new("Sheet1");
        ws.set_column_width(1, 15.0);
        assert_eq!(ws.get_column_width(1), Some(15.0));
        assert_eq!(ws.get_column_width(2), None);
    }

    #[test]
    fn test_row_dimensions() {
        let mut ws = Worksheet::new("Sheet1");
        ws.set_row_height(1, 20.0);
        assert_eq!(ws.get_row_height(1), Some(20.0));
        assert_eq!(ws.get_row_height(2), None);
    }

    #[test]
    fn test_protection() {
        let mut ws = Worksheet::new("Sheet1");
        assert!(!ws.is_protected());

        ws.enable_protection(Some("password".to_string()));
        assert!(ws.is_protected());

        ws.disable_protection();
        assert!(!ws.is_protected());
    }

    #[test]
    fn test_dimensions() {
        let mut ws = Worksheet::new("Sheet1");
        ws.set_cell_value(2, 3, "A");
        ws.set_cell_value(5, 1, "B");

        let (min_r, min_c, max_r, max_c) = ws.dimensions();
        assert_eq!((min_r, min_c), (2, 1));
        assert_eq!((max_r, max_c), (5, 3));
    }
}

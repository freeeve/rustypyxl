//! Worksheet representation and cell operations.

#[cfg(feature = "fast-hash")]
use hashbrown::HashMap;
#[cfg(not(feature = "fast-hash"))]
use std::collections::HashMap;
use std::sync::Arc;
use crate::cell::CellValue;
use crate::style::CellStyle;
use crate::autofilter::AutoFilter;
use crate::conditional::ConditionalFormatting;
use crate::table::Table;
use crate::pagesetup::PageSetup;

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
    /// Number format string.
    pub number_format: Option<String>,
    /// Data type (s=string, n=number, b=boolean, d=date).
    pub data_type: Option<String>,
    /// Hyperlink URL.
    pub hyperlink: Option<String>,
    /// Cell comment text.
    pub comment: Option<String>,
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

/// Data validation rule for a cell.
#[derive(Clone, Debug)]
pub struct DataValidation {
    /// Type: whole, decimal, list, date, time, textLength, custom.
    pub validation_type: String,
    /// First formula/value constraint.
    pub formula1: Option<String>,
    /// Second formula/value constraint (for between/notBetween).
    pub formula2: Option<String>,
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
}

impl Default for DataValidation {
    fn default() -> Self {
        DataValidation {
            validation_type: "whole".to_string(),
            formula1: None,
            formula2: None,
            allow_blank: true,
            show_error: true,
            error_title: None,
            error_message: None,
            show_input: true,
            prompt_title: None,
            prompt_message: None,
        }
    }
}

/// Worksheet protection settings.
#[derive(Clone, Debug, Default)]
pub struct WorksheetProtection {
    /// Sheet protection enabled.
    pub sheet: bool,
    /// Password hash (for Excel compatibility).
    pub password: Option<String>,
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
        }
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
        let cell_data = self
            .cells
            .entry(cell_key(row, column))
            .or_insert_with(CellData::new);
        cell_data.value = value.into();
        self.update_dimensions(row, column);
    }

    /// Set complete cell data at the specified position.
    pub fn set_cell_data(&mut self, row: u32, column: u32, data: CellData) {
        self.cells.insert(cell_key(row, column), data);
        self.update_dimensions(row, column);
    }

    /// Set a formula in a cell.
    pub fn set_cell_formula<S: Into<String>>(&mut self, row: u32, column: u32, formula: S) {
        let cell_data = self
            .cells
            .entry(cell_key(row, column))
            .or_insert_with(CellData::new);
        cell_data.value = CellValue::Formula(formula.into());
        self.update_dimensions(row, column);
    }

    /// Set a cell's hyperlink.
    pub fn set_cell_hyperlink(&mut self, row: u32, column: u32, url: String) {
        let cell_data = self
            .cells
            .entry(cell_key(row, column))
            .or_insert_with(CellData::new);
        cell_data.hyperlink = Some(url);
        self.update_dimensions(row, column);
    }

    /// Set a cell's comment.
    pub fn set_cell_comment(&mut self, row: u32, column: u32, comment: String) {
        let cell_data = self
            .cells
            .entry(cell_key(row, column))
            .or_insert_with(CellData::new);
        cell_data.comment = Some(comment);
        self.update_dimensions(row, column);
    }

    /// Set a cell's style.
    pub fn set_cell_style(&mut self, row: u32, column: u32, style: CellStyle) {
        let cell_data = self
            .cells
            .entry(cell_key(row, column))
            .or_insert_with(CellData::new);
        cell_data.style = Some(Arc::new(style));
        self.update_dimensions(row, column);
    }

    /// Set a cell's number format.
    pub fn set_cell_number_format<S: Into<String>>(&mut self, row: u32, column: u32, format: S) {
        let cell_data = self
            .cells
            .entry(cell_key(row, column))
            .or_insert_with(CellData::new);
        cell_data.number_format = Some(format.into());
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
            self.merged_cells.retain(|(s, e)| !(s == &start && e == &end));
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
        let mut protection = WorksheetProtection::default();
        protection.sheet = true;
        protection.password = password;
        self.protection = Some(protection);
    }

    /// Disable sheet protection.
    pub fn disable_protection(&mut self) {
        self.protection = None;
    }

    /// Check if sheet is protected.
    pub fn is_protected(&self) -> bool {
        self.protection.as_ref().map_or(false, |p| p.sheet)
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

    /// Iterate over cells in a specific row.
    pub fn iter_row(&self, row: u32) -> impl Iterator<Item = (u32, &CellData)> + '_ {
        let mut cells: Vec<_> = self
            .cells
            .iter()
            .filter_map(|(k, v)| {
                let (r, c) = decode_cell_key(*k);
                if r == row { Some((c, v)) } else { None }
            })
            .collect();
        cells.sort_by_key(|(c, _)| *c);
        cells.into_iter()
    }

    /// Update max_row and max_column.
    fn update_dimensions(&mut self, row: u32, column: u32) {
        self.max_row = self.max_row.max(row);
        self.max_column = self.max_column.max(column);
    }
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

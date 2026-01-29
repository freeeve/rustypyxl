//! Excel Table (ListObject) support.
//!
//! This module provides structures for creating and managing Excel Tables,
//! which provide structured references, auto-filtering, and formatting.

/// Table style preset names.
#[derive(Clone, Debug, PartialEq)]
pub enum TableStyle {
    /// No style
    None,
    /// Light style 1-21
    Light(u32),
    /// Medium style 1-28
    Medium(u32),
    /// Dark style 1-11
    Dark(u32),
    /// Custom style name
    Custom(String),
}

impl TableStyle {
    /// Get the style name for XML.
    pub fn style_name(&self) -> String {
        match self {
            TableStyle::None => "TableStyleMedium2".to_string(),
            TableStyle::Light(n) => format!("TableStyleLight{}", n.clamp(&1, &21)),
            TableStyle::Medium(n) => format!("TableStyleMedium{}", n.clamp(&1, &28)),
            TableStyle::Dark(n) => format!("TableStyleDark{}", n.clamp(&1, &11)),
            TableStyle::Custom(name) => name.clone(),
        }
    }

    /// Common presets.
    pub fn blue() -> Self {
        TableStyle::Medium(2)
    }

    pub fn green() -> Self {
        TableStyle::Medium(7)
    }

    pub fn red() -> Self {
        TableStyle::Medium(3)
    }

    pub fn gray() -> Self {
        TableStyle::Medium(1)
    }
}

impl Default for TableStyle {
    fn default() -> Self {
        TableStyle::Medium(2)
    }
}

/// Totals row function for a column.
#[derive(Clone, Debug, PartialEq)]
pub enum TotalsRowFunction {
    /// No function
    None,
    /// Average
    Average,
    /// Count
    Count,
    /// Count numbers
    CountNums,
    /// Maximum
    Max,
    /// Minimum
    Min,
    /// Standard deviation
    StdDev,
    /// Sum
    Sum,
    /// Variance
    Var,
    /// Custom formula
    Custom(String),
}

impl TotalsRowFunction {
    /// Get the XML function name.
    pub fn xml_name(&self) -> Option<&str> {
        match self {
            TotalsRowFunction::None => None,
            TotalsRowFunction::Average => Some("average"),
            TotalsRowFunction::Count => Some("count"),
            TotalsRowFunction::CountNums => Some("countNums"),
            TotalsRowFunction::Max => Some("max"),
            TotalsRowFunction::Min => Some("min"),
            TotalsRowFunction::StdDev => Some("stdDev"),
            TotalsRowFunction::Sum => Some("sum"),
            TotalsRowFunction::Var => Some("var"),
            TotalsRowFunction::Custom(_) => Some("custom"),
        }
    }
}

/// A column in an Excel Table.
#[derive(Clone, Debug)]
pub struct TableColumn {
    /// Column ID (1-based).
    pub id: u32,
    /// Column name/header.
    pub name: String,
    /// Totals row function.
    pub totals_row_function: TotalsRowFunction,
    /// Totals row label (if no function).
    pub totals_row_label: Option<String>,
    /// Column formula (for calculated columns).
    pub calculated_column_formula: Option<String>,
}

impl TableColumn {
    /// Create a new table column.
    pub fn new(id: u32, name: &str) -> Self {
        TableColumn {
            id,
            name: name.to_string(),
            totals_row_function: TotalsRowFunction::None,
            totals_row_label: None,
            calculated_column_formula: None,
        }
    }

    /// Set the totals row function.
    pub fn with_totals_function(mut self, function: TotalsRowFunction) -> Self {
        self.totals_row_function = function;
        self
    }

    /// Set the totals row label.
    pub fn with_totals_label<S: Into<String>>(mut self, label: S) -> Self {
        self.totals_row_label = Some(label.into());
        self
    }

    /// Set a calculated column formula.
    pub fn with_formula<S: Into<String>>(mut self, formula: S) -> Self {
        self.calculated_column_formula = Some(formula.into());
        self
    }
}

/// An Excel Table (ListObject).
#[derive(Clone, Debug)]
pub struct Table {
    /// Table ID (unique within workbook).
    pub id: u32,
    /// Table name (unique within workbook).
    pub name: String,
    /// Display name.
    pub display_name: String,
    /// Cell range reference (e.g., "A1:D10").
    pub range: String,
    /// Table columns.
    pub columns: Vec<TableColumn>,
    /// Table style.
    pub style: TableStyle,
    /// Show header row.
    pub header_row: bool,
    /// Show totals row.
    pub totals_row: bool,
    /// Show first column formatting.
    pub show_first_column: bool,
    /// Show last column formatting.
    pub show_last_column: bool,
    /// Show row stripes.
    pub show_row_stripes: bool,
    /// Show column stripes.
    pub show_column_stripes: bool,
    /// AutoFilter enabled.
    pub auto_filter: bool,
    /// Comment/description.
    pub comment: Option<String>,
}

impl Table {
    /// Create a new table.
    pub fn new(id: u32, name: &str, range: &str) -> Self {
        Table {
            id,
            name: name.to_string(),
            display_name: name.to_string(),
            range: range.to_string(),
            columns: Vec::new(),
            style: TableStyle::default(),
            header_row: true,
            totals_row: false,
            show_first_column: false,
            show_last_column: false,
            show_row_stripes: true,
            show_column_stripes: false,
            auto_filter: true,
            comment: None,
        }
    }

    /// Create a table with auto-generated columns from header names.
    pub fn with_headers(id: u32, name: &str, range: &str, headers: &[&str]) -> Self {
        let mut table = Self::new(id, name, range);
        for (i, header) in headers.iter().enumerate() {
            table.columns.push(TableColumn::new((i + 1) as u32, header));
        }
        table
    }

    /// Set the table style.
    pub fn with_style(mut self, style: TableStyle) -> Self {
        self.style = style;
        self
    }

    /// Enable totals row.
    pub fn with_totals_row(mut self) -> Self {
        self.totals_row = true;
        self
    }

    /// Disable header row.
    pub fn without_header_row(mut self) -> Self {
        self.header_row = false;
        self
    }

    /// Set first column highlighting.
    pub fn with_first_column(mut self) -> Self {
        self.show_first_column = true;
        self
    }

    /// Set last column highlighting.
    pub fn with_last_column(mut self) -> Self {
        self.show_last_column = true;
        self
    }

    /// Enable column stripes instead of row stripes.
    pub fn with_column_stripes(mut self) -> Self {
        self.show_row_stripes = false;
        self.show_column_stripes = true;
        self
    }

    /// Disable auto-filter.
    pub fn without_auto_filter(mut self) -> Self {
        self.auto_filter = false;
        self
    }

    /// Add a column.
    pub fn add_column(&mut self, column: TableColumn) {
        self.columns.push(column);
    }

    /// Set totals function for a column by name.
    pub fn set_column_totals(&mut self, column_name: &str, function: TotalsRowFunction) {
        for col in &mut self.columns {
            if col.name == column_name {
                col.totals_row_function = function;
                break;
            }
        }
    }

    /// Get a structured reference for the table.
    pub fn structured_ref(&self, column: Option<&str>, specifier: Option<&str>) -> String {
        let mut ref_str = format!("{}[", self.name);

        if let Some(spec) = specifier {
            ref_str.push_str(&format!("[{}]", spec));
        }

        if let Some(col) = column {
            if specifier.is_some() {
                ref_str.push(',');
            }
            ref_str.push_str(&format!("[{}]", col));
        }

        ref_str.push(']');
        ref_str
    }
}

/// Structured reference specifiers.
pub mod specifiers {
    /// All data (no headers or totals).
    pub const DATA: &str = "#Data";
    /// Headers row only.
    pub const HEADERS: &str = "#Headers";
    /// Totals row only.
    pub const TOTALS: &str = "#Totals";
    /// All including headers and totals.
    pub const ALL: &str = "#All";
    /// Current row (for formulas within table).
    pub const THIS_ROW: &str = "#This Row";
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_table_creation() {
        let table = Table::new(1, "SalesData", "A1:D100");
        assert_eq!(table.name, "SalesData");
        assert_eq!(table.range, "A1:D100");
        assert!(table.header_row);
        assert!(!table.totals_row);
    }

    #[test]
    fn test_table_with_headers() {
        let table = Table::with_headers(
            1,
            "Products",
            "A1:C10",
            &["Name", "Price", "Quantity"],
        );

        assert_eq!(table.columns.len(), 3);
        assert_eq!(table.columns[0].name, "Name");
        assert_eq!(table.columns[1].name, "Price");
    }

    #[test]
    fn test_table_style() {
        let table = Table::new(1, "Test", "A1:B10").with_style(TableStyle::blue());
        assert_eq!(table.style.style_name(), "TableStyleMedium2");
    }

    #[test]
    fn test_table_totals() {
        let mut table = Table::with_headers(1, "Data", "A1:B10", &["Item", "Value"]);
        table = table.with_totals_row();
        table.set_column_totals("Value", TotalsRowFunction::Sum);

        assert!(table.totals_row);
        assert_eq!(
            table.columns[1].totals_row_function,
            TotalsRowFunction::Sum
        );
    }

    #[test]
    fn test_structured_reference() {
        let table = Table::new(1, "Sales", "A1:C10");

        assert_eq!(table.structured_ref(Some("Amount"), None), "Sales[[Amount]]");
        assert_eq!(
            table.structured_ref(Some("Amount"), Some("#Data")),
            "Sales[[#Data],[Amount]]"
        );
        assert_eq!(table.structured_ref(None, Some("#Totals")), "Sales[[#Totals]]");
    }

    #[test]
    fn test_table_column_formula() {
        let col = TableColumn::new(1, "Total")
            .with_formula("[@Price]*[@Quantity]");

        assert_eq!(
            col.calculated_column_formula,
            Some("[@Price]*[@Quantity]".to_string())
        );
    }
}

//! AutoFilter support for Excel workbooks.
//!
//! This module provides structures for creating and managing AutoFilter in worksheets.

/// Filter type for a column.
#[derive(Clone, Debug, PartialEq)]
pub enum FilterType {
    /// Filter by specific values.
    Values(Vec<String>),
    /// Custom filter with operators.
    Custom(CustomFilter),
    /// Filter by color.
    ColorFilter(ColorFilter),
    /// Dynamic filter (e.g., today, this month).
    DynamicFilter(DynamicFilterType),
    /// Top/bottom N filter.
    Top10Filter(Top10Filter),
}

/// Custom filter with operator and value.
#[derive(Clone, Debug, PartialEq)]
pub struct CustomFilter {
    /// First operator.
    pub operator1: FilterOperator,
    /// First value.
    pub value1: String,
    /// AND or OR.
    pub and: bool,
    /// Second operator (optional).
    pub operator2: Option<FilterOperator>,
    /// Second value (optional).
    pub value2: Option<String>,
}

impl CustomFilter {
    /// Create a single custom filter.
    pub fn new(operator: FilterOperator, value: &str) -> Self {
        CustomFilter {
            operator1: operator,
            value1: value.to_string(),
            and: true,
            operator2: None,
            value2: None,
        }
    }

    /// Add a second condition with AND.
    pub fn and(mut self, operator: FilterOperator, value: &str) -> Self {
        self.and = true;
        self.operator2 = Some(operator);
        self.value2 = Some(value.to_string());
        self
    }

    /// Add a second condition with OR.
    pub fn or(mut self, operator: FilterOperator, value: &str) -> Self {
        self.and = false;
        self.operator2 = Some(operator);
        self.value2 = Some(value.to_string());
        self
    }
}

/// Filter operator for custom filters.
#[derive(Clone, Debug, PartialEq)]
pub enum FilterOperator {
    Equal,
    NotEqual,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
}

impl FilterOperator {
    /// Get the XML attribute value.
    pub fn xml_value(&self) -> &'static str {
        match self {
            FilterOperator::Equal => "equal",
            FilterOperator::NotEqual => "notEqual",
            FilterOperator::GreaterThan => "greaterThan",
            FilterOperator::GreaterThanOrEqual => "greaterThanOrEqual",
            FilterOperator::LessThan => "lessThan",
            FilterOperator::LessThanOrEqual => "lessThanOrEqual",
        }
    }
}

/// Color filter configuration.
#[derive(Clone, Debug, PartialEq)]
pub struct ColorFilter {
    /// Filter by cell color (true) or font color (false).
    pub cell_color: bool,
    /// Color value (theme index or RGB).
    pub color: String,
}

/// Dynamic filter type.
#[derive(Clone, Debug, PartialEq)]
pub enum DynamicFilterType {
    Today,
    Yesterday,
    Tomorrow,
    ThisWeek,
    NextWeek,
    LastWeek,
    ThisMonth,
    NextMonth,
    LastMonth,
    ThisQuarter,
    NextQuarter,
    LastQuarter,
    ThisYear,
    NextYear,
    LastYear,
    YearToDate,
    AboveAverage,
    BelowAverage,
}

impl DynamicFilterType {
    /// Get the XML type name.
    pub fn xml_type(&self) -> &'static str {
        match self {
            DynamicFilterType::Today => "today",
            DynamicFilterType::Yesterday => "yesterday",
            DynamicFilterType::Tomorrow => "tomorrow",
            DynamicFilterType::ThisWeek => "thisWeek",
            DynamicFilterType::NextWeek => "nextWeek",
            DynamicFilterType::LastWeek => "lastWeek",
            DynamicFilterType::ThisMonth => "thisMonth",
            DynamicFilterType::NextMonth => "nextMonth",
            DynamicFilterType::LastMonth => "lastMonth",
            DynamicFilterType::ThisQuarter => "thisQuarter",
            DynamicFilterType::NextQuarter => "nextQuarter",
            DynamicFilterType::LastQuarter => "lastQuarter",
            DynamicFilterType::ThisYear => "thisYear",
            DynamicFilterType::NextYear => "nextYear",
            DynamicFilterType::LastYear => "lastYear",
            DynamicFilterType::YearToDate => "yearToDate",
            DynamicFilterType::AboveAverage => "aboveAverage",
            DynamicFilterType::BelowAverage => "belowAverage",
        }
    }
}

/// Top/bottom N filter.
#[derive(Clone, Debug, PartialEq)]
pub struct Top10Filter {
    /// Filter top N (true) or bottom N (false).
    pub top: bool,
    /// Number of items or percent.
    pub value: f64,
    /// Filter by percent (true) or count (false).
    pub percent: bool,
}

impl Top10Filter {
    /// Top N items.
    pub fn top(n: u32) -> Self {
        Top10Filter {
            top: true,
            value: n as f64,
            percent: false,
        }
    }

    /// Bottom N items.
    pub fn bottom(n: u32) -> Self {
        Top10Filter {
            top: false,
            value: n as f64,
            percent: false,
        }
    }

    /// Top N percent.
    pub fn top_percent(pct: f64) -> Self {
        Top10Filter {
            top: true,
            value: pct,
            percent: true,
        }
    }
}

/// Filter column configuration.
#[derive(Clone, Debug)]
pub struct FilterColumn {
    /// Column index (0-based).
    pub column_id: u32,
    /// Filter type and configuration.
    pub filter: FilterType,
    /// Show button (default true).
    pub show_button: bool,
}

impl FilterColumn {
    /// Create a value filter for a column.
    pub fn values(column_id: u32, values: Vec<String>) -> Self {
        FilterColumn {
            column_id,
            filter: FilterType::Values(values),
            show_button: true,
        }
    }

    /// Create a custom filter for a column.
    pub fn custom(column_id: u32, filter: CustomFilter) -> Self {
        FilterColumn {
            column_id,
            filter: FilterType::Custom(filter),
            show_button: true,
        }
    }
}

/// AutoFilter configuration for a worksheet.
#[derive(Clone, Debug)]
pub struct AutoFilter {
    /// Range reference (e.g., "A1:D100").
    pub range: String,
    /// Column filters.
    pub columns: Vec<FilterColumn>,
    /// Sort state.
    pub sort_column: Option<u32>,
    /// Sort descending.
    pub sort_descending: bool,
}

impl AutoFilter {
    /// Create an AutoFilter for the specified range.
    pub fn new<S: Into<String>>(range: S) -> Self {
        AutoFilter {
            range: range.into(),
            columns: Vec::new(),
            sort_column: None,
            sort_descending: false,
        }
    }

    /// Add a column filter.
    pub fn add_filter(&mut self, column: FilterColumn) {
        self.columns.push(column);
    }

    /// Set sort column.
    pub fn sort_by(&mut self, column: u32, descending: bool) {
        self.sort_column = Some(column);
        self.sort_descending = descending;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_autofilter_creation() {
        let af = AutoFilter::new("A1:D100");
        assert_eq!(af.range, "A1:D100");
        assert!(af.columns.is_empty());
    }

    #[test]
    fn test_value_filter() {
        let mut af = AutoFilter::new("A1:B10");
        af.add_filter(FilterColumn::values(0, vec!["Apple".to_string(), "Orange".to_string()]));

        assert_eq!(af.columns.len(), 1);
    }

    #[test]
    fn test_custom_filter() {
        let filter = CustomFilter::new(FilterOperator::GreaterThan, "100")
            .and(FilterOperator::LessThan, "200");

        let mut af = AutoFilter::new("A1:A100");
        af.add_filter(FilterColumn::custom(0, filter));

        assert_eq!(af.columns.len(), 1);
    }

    #[test]
    fn test_top10_filter() {
        let filter = Top10Filter::top(10);
        assert!(filter.top);
        assert!(!filter.percent);
        assert_eq!(filter.value, 10.0);
    }
}

//! Cell value types and utilities.

use std::fmt;
use std::sync::Arc;

pub type InternedString = Arc<str>;

/// Represents the value of a cell in an Excel worksheet.
#[derive(Clone, Debug, PartialEq)]
pub enum CellValue {
    /// String value.
    String(InternedString),
    /// Numeric value (stored as f64 for Excel compatibility).
    Number(f64),
    /// Boolean value.
    Boolean(bool),
    /// Date value stored as ISO 8601 string or Excel serial number.
    Date(String),
    /// Formula (without the leading '=' sign).
    Formula(String),
    /// Empty cell.
    Empty,
}

impl CellValue {
    /// Check if the cell value is empty.
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }

    /// Check if the cell value is a formula.
    pub fn is_formula(&self) -> bool {
        matches!(self, CellValue::Formula(_))
    }

    /// Get the value as a string representation.
    pub fn as_string(&self) -> Option<&str> {
        match self {
            CellValue::String(s) => Some(s.as_ref()),
            _ => None,
        }
    }

    /// Get the value as a number.
    pub fn as_number(&self) -> Option<f64> {
        match self {
            CellValue::Number(n) => Some(*n),
            _ => None,
        }
    }

    /// Get the value as a boolean.
    pub fn as_bool(&self) -> Option<bool> {
        match self {
            CellValue::Boolean(b) => Some(*b),
            _ => None,
        }
    }

    /// Get the formula string (without '=' prefix).
    pub fn as_formula(&self) -> Option<&str> {
        match self {
            CellValue::Formula(f) => Some(f),
            _ => None,
        }
    }

    /// Get the Excel data type code.
    pub fn data_type_code(&self) -> &'static str {
        match self {
            CellValue::String(_) => "s",
            CellValue::Number(_) => "n",
            CellValue::Boolean(_) => "b",
            CellValue::Date(_) => "d",
            CellValue::Formula(_) => "str",
            CellValue::Empty => "",
        }
    }
}

impl Default for CellValue {
    fn default() -> Self {
        CellValue::Empty
    }
}

impl fmt::Display for CellValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            CellValue::String(s) => write!(f, "{}", s.as_ref()),
            CellValue::Number(n) => write!(f, "{}", n),
            CellValue::Boolean(b) => write!(f, "{}", if *b { "TRUE" } else { "FALSE" }),
            CellValue::Date(d) => write!(f, "{}", d),
            CellValue::Formula(formula) => write!(f, "={}", formula),
            CellValue::Empty => write!(f, ""),
        }
    }
}

impl From<String> for CellValue {
    fn from(s: String) -> Self {
        CellValue::String(Arc::from(s))
    }
}

impl From<&str> for CellValue {
    fn from(s: &str) -> Self {
        CellValue::String(Arc::from(s))
    }
}

impl From<f64> for CellValue {
    fn from(n: f64) -> Self {
        CellValue::Number(n)
    }
}

impl From<i32> for CellValue {
    fn from(n: i32) -> Self {
        CellValue::Number(n as f64)
    }
}

impl From<i64> for CellValue {
    fn from(n: i64) -> Self {
        CellValue::Number(n as f64)
    }
}

impl From<bool> for CellValue {
    fn from(b: bool) -> Self {
        CellValue::Boolean(b)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_value_string() {
        let val = CellValue::String(Arc::from("Hello"));
        assert_eq!(val.as_string(), Some("Hello"));
        assert_eq!(val.to_string(), "Hello");
        assert_eq!(val.data_type_code(), "s");
    }

    #[test]
    fn test_cell_value_number() {
        let val = CellValue::Number(42.5);
        assert_eq!(val.as_number(), Some(42.5));
        assert_eq!(val.to_string(), "42.5");
        assert_eq!(val.data_type_code(), "n");
    }

    #[test]
    fn test_cell_value_boolean() {
        let val = CellValue::Boolean(true);
        assert_eq!(val.as_bool(), Some(true));
        assert_eq!(val.to_string(), "TRUE");
        assert_eq!(val.data_type_code(), "b");
    }

    #[test]
    fn test_cell_value_formula() {
        let val = CellValue::Formula("SUM(A1:A10)".to_string());
        assert!(val.is_formula());
        assert_eq!(val.as_formula(), Some("SUM(A1:A10)"));
        assert_eq!(val.to_string(), "=SUM(A1:A10)");
    }

    #[test]
    fn test_cell_value_empty() {
        let val = CellValue::Empty;
        assert!(val.is_empty());
        assert_eq!(val.to_string(), "");
    }

    #[test]
    fn test_cell_value_from() {
        let val: CellValue = "Hello".into();
        assert_eq!(val, CellValue::String(Arc::from("Hello")));

        let val: CellValue = 42.5.into();
        assert_eq!(val, CellValue::Number(42.5));

        let val: CellValue = true.into();
        assert_eq!(val, CellValue::Boolean(true));
    }
}

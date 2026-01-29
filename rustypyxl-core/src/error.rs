//! Error types for rustypyxl-core.

use thiserror::Error;

/// Result type alias using RustypyxlError.
pub type Result<T> = std::result::Result<T, RustypyxlError>;

/// Errors that can occur when working with Excel files.
#[derive(Error, Debug)]
pub enum RustypyxlError {
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),

    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("Invalid cell coordinate: {0}")]
    InvalidCoordinate(String),

    #[error("Worksheet not found: {0}")]
    WorksheetNotFound(String),

    #[error("Worksheet already exists: {0}")]
    WorksheetAlreadyExists(String),

    #[error("Named range not found: {0}")]
    NamedRangeNotFound(String),

    #[error("Named range already exists: {0}")]
    NamedRangeAlreadyExists(String),

    #[error("No worksheets in workbook")]
    NoWorksheets,

    #[error("Invalid file format: {0}")]
    InvalidFormat(String),

    #[error("Parse error: {0}")]
    ParseError(String),

    #[error("{0}")]
    Custom(String),
}

impl RustypyxlError {
    /// Create a custom error with a message.
    pub fn custom<S: Into<String>>(msg: S) -> Self {
        RustypyxlError::Custom(msg.into())
    }
}

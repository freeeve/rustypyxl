//! rustypyxl: Fast Rust library for reading and writing Excel xlsx files.
//!
//! This crate provides the functionality for working with Excel files without
//! any Python dependencies. The `rustypyxl` package on PyPI is a PyO3 binding
//! over this same library, exposing an openpyxl-compatible API.
//!
//! # Example
//!
//! ```no_run
//! use rustypyxl::{Workbook, CellValue};
//!
//! // Create a new workbook
//! let mut wb = Workbook::new();
//! let ws = wb.create_sheet(Some("Data".to_string())).unwrap();
//!
//! // Set cell values
//! wb.set_cell_value_in_sheet("Data", 1, 1, CellValue::from("Hello")).unwrap();
//! wb.set_cell_value_in_sheet("Data", 1, 2, CellValue::Number(42.0)).unwrap();
//!
//! // Save the workbook
//! wb.save("output.xlsx").unwrap();
//!
//! // Load an existing workbook
//! let wb = Workbook::load("input.xlsx").unwrap();
//! let ws = wb.active().unwrap();
//! println!("Sheet title: {}", ws.title());
//! ```

pub mod cell;
pub mod chart;
pub mod chart_writer;
pub mod conditional;
pub mod drawing_writer;
pub mod error;
pub mod formula;
pub mod image;
pub mod numfmt;
pub mod rich_text;
pub mod style;
pub mod utils;
pub mod workbook;
pub mod worksheet;
pub mod writer;

// Phase 3 additional modules
pub mod autofilter;
pub mod pagesetup;
pub mod streaming;
pub mod table;

// Optional parquet support
#[cfg(feature = "parquet")]
pub mod parquet_import;

// Optional S3 support
#[cfg(feature = "s3")]
pub mod s3;

// Re-export main types at crate level
pub use cell::CellValue;
pub use error::{Result, RustypyxlError};
pub use formula::{evaluate as evaluate_formula, CellResolver, FormulaValue};
pub use numfmt::{builtin_format_code, format_number, format_value};
pub use rich_text::{RichText, RunFont, TextRun};
pub use style::{
    Alignment, Border, BorderStyle, CellStyle, Color, Fill, Font, GradientFill, GradientStop,
    Protection,
};
pub use utils::{
    column_to_letter, coordinate_from_row_col, letter_to_column, parse_coordinate,
    parse_coordinate_bytes, parse_f64_bytes, parse_range, parse_u32_bytes,
};
pub use workbook::{CompressionLevel, NamedRange, Workbook};
pub use worksheet::{CellData, DataValidation, SheetVisibility, Worksheet, WorksheetProtection};

#[cfg(feature = "parquet")]
pub use parquet_import::{
    ColumnType, ParquetCompression, ParquetExportOptions, ParquetExportResult,
    ParquetImportOptions, ParquetImportResult,
};

#[cfg(feature = "s3")]
pub use s3::S3Config;

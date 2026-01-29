//! rustypyxl-core: Core Rust library for reading and writing Excel xlsx files.
//!
//! This crate provides the core functionality for working with Excel files
//! without any Python dependencies. For Python bindings, see the `rustypyxl-pyo3` crate.
//!
//! # Example
//!
//! ```no_run
//! use rustypyxl_core::{Workbook, CellValue};
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
pub mod conditional;
pub mod error;
pub mod image;
pub mod style;
pub mod utils;
pub mod workbook;
pub mod worksheet;
pub mod writer;

// Phase 3 additional modules
pub mod autofilter;
pub mod pagesetup;
pub mod table;

// Optional parquet support
#[cfg(feature = "parquet")]
pub mod parquet_import;

// Re-export main types at crate level
pub use cell::CellValue;
pub use error::{Result, RustypyxlError};
pub use style::{Alignment, Border, BorderStyle, CellStyle, Fill, Font};
pub use utils::{column_to_letter, coordinate_from_row_col, letter_to_column, parse_coordinate, parse_coordinate_bytes, parse_f64_bytes, parse_u32_bytes, parse_range};
pub use workbook::{CompressionLevel, NamedRange, Workbook};
pub use worksheet::{CellData, DataValidation, Worksheet, WorksheetProtection};

#[cfg(feature = "parquet")]
pub use parquet_import::{ParquetImportOptions, ParquetImportResult};

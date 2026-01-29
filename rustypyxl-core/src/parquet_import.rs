//! Parquet file import functionality.
//!
//! This module provides fast import of Parquet files directly into Excel worksheets,
//! bypassing FFI overhead for maximum performance.

use crate::cell::CellValue;
use crate::error::{Result, RustypyxlError};
use crate::worksheet::Worksheet;
use crate::Workbook;

use arrow::array::*;
use arrow::datatypes::{DataType, TimeUnit};
use parquet::arrow::arrow_reader::ParquetRecordBatchReaderBuilder;
use std::collections::HashMap;
use std::fs::File;
use std::sync::Arc;

/// Result of a parquet import operation.
#[derive(Debug, Clone)]
pub struct ParquetImportResult {
    /// Number of rows imported (excluding header).
    pub rows_imported: u32,
    /// Number of columns imported.
    pub columns_imported: u32,
    /// Starting row of data (1-indexed).
    pub start_row: u32,
    /// Starting column of data (1-indexed).
    pub start_col: u32,
    /// Ending row of data (1-indexed).
    pub end_row: u32,
    /// Ending column of data (1-indexed).
    pub end_col: u32,
    /// Column names as imported (after any renaming).
    pub column_names: Vec<String>,
}

impl ParquetImportResult {
    /// Get the range string (e.g., "A1:Z1000") for the imported data including headers.
    pub fn range_with_headers(&self) -> String {
        format!(
            "{}{}:{}{}",
            crate::utils::column_to_letter(self.start_col),
            self.start_row,
            crate::utils::column_to_letter(self.end_col),
            self.end_row
        )
    }

    /// Get the range string for just the data (excluding headers).
    pub fn data_range(&self) -> String {
        format!(
            "{}{}:{}{}",
            crate::utils::column_to_letter(self.start_col),
            self.start_row + 1,
            crate::utils::column_to_letter(self.end_col),
            self.end_row
        )
    }

    /// Get the range string for just the headers.
    pub fn header_range(&self) -> String {
        format!(
            "{}{}:{}{}",
            crate::utils::column_to_letter(self.start_col),
            self.start_row,
            crate::utils::column_to_letter(self.end_col),
            self.start_row
        )
    }
}

/// Options for parquet import.
#[derive(Debug, Clone, Default)]
pub struct ParquetImportOptions {
    /// Column name mappings (original_name -> new_name).
    pub column_renames: HashMap<String, String>,
    /// If true, include headers in the first row. Default: true.
    pub include_headers: bool,
    /// Specific columns to import (by name). If empty, import all.
    pub columns: Vec<String>,
    /// Batch size for reading. Default: 65536.
    pub batch_size: usize,
}

impl ParquetImportOptions {
    pub fn new() -> Self {
        Self {
            column_renames: HashMap::new(),
            include_headers: true,
            columns: Vec::new(),
            batch_size: 65536,
        }
    }

    /// Add a column rename mapping.
    pub fn rename_column(mut self, from: &str, to: &str) -> Self {
        self.column_renames.insert(from.to_string(), to.to_string());
        self
    }

    /// Set whether to include headers.
    pub fn with_headers(mut self, include: bool) -> Self {
        self.include_headers = include;
        self
    }

    /// Select specific columns to import.
    pub fn select_columns(mut self, columns: Vec<String>) -> Self {
        self.columns = columns;
        self
    }

    /// Set batch size for reading.
    pub fn with_batch_size(mut self, size: usize) -> Self {
        self.batch_size = size;
        self
    }
}

impl Workbook {
    /// Import data from a Parquet file into a worksheet.
    ///
    /// This is the fastest way to load large datasets into Excel, as it
    /// bypasses the Python FFI entirely and reads directly from Parquet
    /// into the internal cell storage.
    ///
    /// # Arguments
    /// * `sheet_name` - Name of the worksheet to insert into
    /// * `path` - Path to the Parquet file
    /// * `start_row` - Starting row (1-indexed)
    /// * `start_col` - Starting column (1-indexed)
    /// * `options` - Import options (headers, column renames, etc.)
    ///
    /// # Returns
    /// Information about what was imported, including the range.
    pub fn insert_from_parquet(
        &mut self,
        sheet_name: &str,
        path: &str,
        start_row: u32,
        start_col: u32,
        options: Option<ParquetImportOptions>,
    ) -> Result<ParquetImportResult> {
        let options = options.unwrap_or_default();
        let opts = if options.batch_size == 0 {
            ParquetImportOptions {
                batch_size: 65536,
                ..options
            }
        } else {
            options
        };

        // Open the parquet file
        let file = File::open(path).map_err(|e| {
            RustypyxlError::ParseError(format!("Failed to open parquet file: {}", e))
        })?;

        // Build the reader
        let builder = ParquetRecordBatchReaderBuilder::try_new(file).map_err(|e| {
            RustypyxlError::ParseError(format!("Failed to read parquet metadata: {}", e))
        })?;

        // Get schema and determine columns to read
        let schema = builder.schema().clone();
        let all_column_names: Vec<String> = schema.fields().iter().map(|f| f.name().clone()).collect();

        // Determine which columns to import
        let columns_to_import: Vec<usize> = if opts.columns.is_empty() {
            (0..all_column_names.len()).collect()
        } else {
            opts.columns
                .iter()
                .filter_map(|name| all_column_names.iter().position(|n| n == name))
                .collect()
        };

        if columns_to_import.is_empty() {
            return Err(RustypyxlError::ParseError(
                "No matching columns found in parquet file".to_string(),
            ));
        }

        // Build reader with batch size
        let reader = builder
            .with_batch_size(opts.batch_size)
            .build()
            .map_err(|e| RustypyxlError::ParseError(format!("Failed to build parquet reader: {}", e)))?;

        // Get the worksheet
        let worksheet = self.get_sheet_by_name_mut(sheet_name)?;

        // Prepare column names (with renames applied)
        let final_column_names: Vec<String> = columns_to_import
            .iter()
            .map(|&idx| {
                let original = &all_column_names[idx];
                opts.column_renames
                    .get(original)
                    .cloned()
                    .unwrap_or_else(|| original.clone())
            })
            .collect();

        let mut current_row = start_row;

        // Write headers if requested
        if opts.include_headers {
            for (col_offset, name) in final_column_names.iter().enumerate() {
                let col = start_col + col_offset as u32;
                worksheet.set_cell_value(current_row, col, CellValue::String(Arc::from(name.as_str())));
            }
            current_row += 1;
        }

        let data_start_row = current_row;
        let mut total_rows: u32 = 0;

        // Read batches and write to worksheet
        for batch_result in reader {
            let batch = batch_result.map_err(|e| {
                RustypyxlError::ParseError(format!("Failed to read parquet batch: {}", e))
            })?;

            let num_rows = batch.num_rows();

            // Process each column
            for (col_offset, &schema_idx) in columns_to_import.iter().enumerate() {
                let col = start_col + col_offset as u32;
                let array = batch.column(schema_idx);

                write_arrow_array_to_worksheet(
                    worksheet,
                    array,
                    current_row,
                    col,
                    num_rows,
                );
            }

            current_row += num_rows as u32;
            total_rows += num_rows as u32;
        }

        let end_row_with_header = if opts.include_headers && total_rows > 0 {
            start_row + total_rows
        } else if total_rows > 0 {
            start_row + total_rows - 1
        } else {
            start_row
        };

        Ok(ParquetImportResult {
            rows_imported: total_rows,
            columns_imported: columns_to_import.len() as u32,
            start_row,
            start_col,
            end_row: end_row_with_header,
            end_col: start_col + columns_to_import.len() as u32 - 1,
            column_names: final_column_names,
        })
    }
}

/// Write an Arrow array to a worksheet column.
fn write_arrow_array_to_worksheet(
    worksheet: &mut Worksheet,
    array: &ArrayRef,
    start_row: u32,
    col: u32,
    num_rows: usize,
) {
    match array.data_type() {
        DataType::Null => {
            // All nulls - nothing to write
        }
        DataType::Boolean => {
            let arr = array.as_any().downcast_ref::<BooleanArray>().unwrap();
            for i in 0..num_rows {
                let row = start_row + i as u32;
                if arr.is_valid(i) {
                    worksheet.set_cell_value(row, col, CellValue::Boolean(arr.value(i)));
                }
            }
        }
        DataType::Int8 => write_int_array::<Int8Array>(worksheet, array, start_row, col, num_rows),
        DataType::Int16 => write_int_array::<Int16Array>(worksheet, array, start_row, col, num_rows),
        DataType::Int32 => write_int_array::<Int32Array>(worksheet, array, start_row, col, num_rows),
        DataType::Int64 => write_int_array::<Int64Array>(worksheet, array, start_row, col, num_rows),
        DataType::UInt8 => write_uint_array::<UInt8Array>(worksheet, array, start_row, col, num_rows),
        DataType::UInt16 => write_uint_array::<UInt16Array>(worksheet, array, start_row, col, num_rows),
        DataType::UInt32 => write_uint_array::<UInt32Array>(worksheet, array, start_row, col, num_rows),
        DataType::UInt64 => write_uint_array::<UInt64Array>(worksheet, array, start_row, col, num_rows),
        DataType::Float16 => {
            let arr = array.as_any().downcast_ref::<Float16Array>().unwrap();
            for i in 0..num_rows {
                let row = start_row + i as u32;
                if arr.is_valid(i) {
                    worksheet.set_cell_value(row, col, CellValue::Number(arr.value(i).to_f64()));
                }
            }
        }
        DataType::Float32 => {
            let arr = array.as_any().downcast_ref::<Float32Array>().unwrap();
            for i in 0..num_rows {
                let row = start_row + i as u32;
                if arr.is_valid(i) {
                    worksheet.set_cell_value(row, col, CellValue::Number(arr.value(i) as f64));
                }
            }
        }
        DataType::Float64 => {
            let arr = array.as_any().downcast_ref::<Float64Array>().unwrap();
            for i in 0..num_rows {
                let row = start_row + i as u32;
                if arr.is_valid(i) {
                    worksheet.set_cell_value(row, col, CellValue::Number(arr.value(i)));
                }
            }
        }
        DataType::Utf8 => {
            let arr = array.as_any().downcast_ref::<StringArray>().unwrap();
            for i in 0..num_rows {
                let row = start_row + i as u32;
                if arr.is_valid(i) {
                    worksheet.set_cell_value(row, col, CellValue::String(Arc::from(arr.value(i))));
                }
            }
        }
        DataType::LargeUtf8 => {
            let arr = array.as_any().downcast_ref::<LargeStringArray>().unwrap();
            for i in 0..num_rows {
                let row = start_row + i as u32;
                if arr.is_valid(i) {
                    worksheet.set_cell_value(row, col, CellValue::String(Arc::from(arr.value(i))));
                }
            }
        }
        DataType::Date32 => {
            let arr = array.as_any().downcast_ref::<Date32Array>().unwrap();
            for i in 0..num_rows {
                let row = start_row + i as u32;
                if arr.is_valid(i) {
                    // Date32 is days since Unix epoch
                    let days = arr.value(i);
                    // Convert to Excel serial number (Excel epoch is 1900-01-01, but with the 1900 leap year bug)
                    // Unix epoch (1970-01-01) is Excel serial 25569
                    let excel_serial = days + 25569;
                    worksheet.set_cell_value(row, col, CellValue::Number(excel_serial as f64));
                }
            }
        }
        DataType::Date64 => {
            let arr = array.as_any().downcast_ref::<Date64Array>().unwrap();
            for i in 0..num_rows {
                let row = start_row + i as u32;
                if arr.is_valid(i) {
                    // Date64 is milliseconds since Unix epoch
                    let ms = arr.value(i);
                    let days = ms as f64 / (24.0 * 60.0 * 60.0 * 1000.0);
                    let excel_serial = days + 25569.0;
                    worksheet.set_cell_value(row, col, CellValue::Number(excel_serial));
                }
            }
        }
        DataType::Timestamp(unit, _tz) => {
            match unit {
                TimeUnit::Second => {
                    let arr = array.as_any().downcast_ref::<TimestampSecondArray>().unwrap();
                    for i in 0..num_rows {
                        let row = start_row + i as u32;
                        if arr.is_valid(i) {
                            let secs = arr.value(i) as f64;
                            let days = secs / (24.0 * 60.0 * 60.0);
                            let excel_serial = days + 25569.0;
                            worksheet.set_cell_value(row, col, CellValue::Number(excel_serial));
                        }
                    }
                }
                TimeUnit::Millisecond => {
                    let arr = array.as_any().downcast_ref::<TimestampMillisecondArray>().unwrap();
                    for i in 0..num_rows {
                        let row = start_row + i as u32;
                        if arr.is_valid(i) {
                            let ms = arr.value(i) as f64;
                            let days = ms / (24.0 * 60.0 * 60.0 * 1000.0);
                            let excel_serial = days + 25569.0;
                            worksheet.set_cell_value(row, col, CellValue::Number(excel_serial));
                        }
                    }
                }
                TimeUnit::Microsecond => {
                    let arr = array.as_any().downcast_ref::<TimestampMicrosecondArray>().unwrap();
                    for i in 0..num_rows {
                        let row = start_row + i as u32;
                        if arr.is_valid(i) {
                            let us = arr.value(i) as f64;
                            let days = us / (24.0 * 60.0 * 60.0 * 1_000_000.0);
                            let excel_serial = days + 25569.0;
                            worksheet.set_cell_value(row, col, CellValue::Number(excel_serial));
                        }
                    }
                }
                TimeUnit::Nanosecond => {
                    let arr = array.as_any().downcast_ref::<TimestampNanosecondArray>().unwrap();
                    for i in 0..num_rows {
                        let row = start_row + i as u32;
                        if arr.is_valid(i) {
                            let ns = arr.value(i) as f64;
                            let days = ns / (24.0 * 60.0 * 60.0 * 1_000_000_000.0);
                            let excel_serial = days + 25569.0;
                            worksheet.set_cell_value(row, col, CellValue::Number(excel_serial));
                        }
                    }
                }
            }
        }
        DataType::Decimal128(_, scale) => {
            let arr = array.as_any().downcast_ref::<Decimal128Array>().unwrap();
            let scale_factor = 10f64.powi(*scale as i32);
            for i in 0..num_rows {
                let row = start_row + i as u32;
                if arr.is_valid(i) {
                    // arr.value(i) returns i128 directly
                    let val = arr.value(i) as f64 / scale_factor;
                    worksheet.set_cell_value(row, col, CellValue::Number(val));
                }
            }
        }
        DataType::Decimal256(_, scale) => {
            let arr = array.as_any().downcast_ref::<Decimal256Array>().unwrap();
            let scale_factor = 10f64.powi(*scale as i32);
            for i in 0..num_rows {
                let row = start_row + i as u32;
                if arr.is_valid(i) {
                    // Convert i256 to f64 - may lose precision for very large numbers
                    let bytes = arr.value(i).to_le_bytes();
                    let val = i128::from_le_bytes(bytes[0..16].try_into().unwrap()) as f64 / scale_factor;
                    worksheet.set_cell_value(row, col, CellValue::Number(val));
                }
            }
        }
        // For other types, convert to string representation
        _ => {
            for i in 0..num_rows {
                let row = start_row + i as u32;
                if array.is_valid(i) {
                    let formatter = arrow::util::display::ArrayFormatter::try_new(
                        array.as_ref(),
                        &arrow::util::display::FormatOptions::default(),
                    );
                    if let Ok(fmt) = formatter {
                        let s = fmt.value(i).to_string();
                        worksheet.set_cell_value(row, col, CellValue::String(Arc::from(s)));
                    }
                }
            }
        }
    }
}

fn write_int_array<T: arrow::array::Array + 'static>(
    worksheet: &mut Worksheet,
    array: &ArrayRef,
    start_row: u32,
    col: u32,
    num_rows: usize,
) where
    T: std::fmt::Debug,
{
    // Use the primitive array trait for numeric types
    if let Some(arr) = array.as_any().downcast_ref::<Int8Array>() {
        for i in 0..num_rows {
            if arr.is_valid(i) {
                worksheet.set_cell_value(start_row + i as u32, col, CellValue::Number(arr.value(i) as f64));
            }
        }
    } else if let Some(arr) = array.as_any().downcast_ref::<Int16Array>() {
        for i in 0..num_rows {
            if arr.is_valid(i) {
                worksheet.set_cell_value(start_row + i as u32, col, CellValue::Number(arr.value(i) as f64));
            }
        }
    } else if let Some(arr) = array.as_any().downcast_ref::<Int32Array>() {
        for i in 0..num_rows {
            if arr.is_valid(i) {
                worksheet.set_cell_value(start_row + i as u32, col, CellValue::Number(arr.value(i) as f64));
            }
        }
    } else if let Some(arr) = array.as_any().downcast_ref::<Int64Array>() {
        for i in 0..num_rows {
            if arr.is_valid(i) {
                worksheet.set_cell_value(start_row + i as u32, col, CellValue::Number(arr.value(i) as f64));
            }
        }
    }
}

fn write_uint_array<T: arrow::array::Array + 'static>(
    worksheet: &mut Worksheet,
    array: &ArrayRef,
    start_row: u32,
    col: u32,
    num_rows: usize,
) where
    T: std::fmt::Debug,
{
    if let Some(arr) = array.as_any().downcast_ref::<UInt8Array>() {
        for i in 0..num_rows {
            if arr.is_valid(i) {
                worksheet.set_cell_value(start_row + i as u32, col, CellValue::Number(arr.value(i) as f64));
            }
        }
    } else if let Some(arr) = array.as_any().downcast_ref::<UInt16Array>() {
        for i in 0..num_rows {
            if arr.is_valid(i) {
                worksheet.set_cell_value(start_row + i as u32, col, CellValue::Number(arr.value(i) as f64));
            }
        }
    } else if let Some(arr) = array.as_any().downcast_ref::<UInt32Array>() {
        for i in 0..num_rows {
            if arr.is_valid(i) {
                worksheet.set_cell_value(start_row + i as u32, col, CellValue::Number(arr.value(i) as f64));
            }
        }
    } else if let Some(arr) = array.as_any().downcast_ref::<UInt64Array>() {
        for i in 0..num_rows {
            if arr.is_valid(i) {
                worksheet.set_cell_value(start_row + i as u32, col, CellValue::Number(arr.value(i) as f64));
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use tempfile::NamedTempFile;

    // Note: These tests require creating parquet files, which needs the parquet writer
    // For now, we'll have basic compilation tests

    #[test]
    fn test_import_options_builder() {
        let opts = ParquetImportOptions::new()
            .rename_column("old_name", "new_name")
            .with_headers(true)
            .select_columns(vec!["col1".to_string(), "col2".to_string()])
            .with_batch_size(1000);

        assert_eq!(opts.column_renames.get("old_name"), Some(&"new_name".to_string()));
        assert!(opts.include_headers);
        assert_eq!(opts.columns, vec!["col1", "col2"]);
        assert_eq!(opts.batch_size, 1000);
    }

    #[test]
    fn test_import_result_ranges() {
        let result = ParquetImportResult {
            rows_imported: 100,
            columns_imported: 5,
            start_row: 1,
            start_col: 1,
            end_row: 101,
            end_col: 5,
            column_names: vec!["A".into(), "B".into(), "C".into(), "D".into(), "E".into()],
        };

        assert_eq!(result.range_with_headers(), "A1:E101");
        assert_eq!(result.data_range(), "A2:E101");
        assert_eq!(result.header_range(), "A1:E1");
    }
}

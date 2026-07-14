//! Parquet file import and export functionality.
//!
//! This module provides fast import/export of Parquet files directly to/from Excel worksheets,
//! bypassing FFI overhead for maximum performance.

use crate::cell::CellValue;
use crate::error::{Result, RustypyxlError};
use crate::worksheet::Worksheet;
use crate::Workbook;

use arrow::array::{
    Array, ArrayRef, BooleanArray, Date32Array, Date64Array, Decimal128Array, Decimal256Array,
    Float16Array, Float32Array, Float64Array, Int16Array, Int32Array, Int64Array, Int8Array,
    LargeStringArray, StringArray, TimestampMicrosecondArray, TimestampMillisecondArray,
    TimestampNanosecondArray, TimestampSecondArray, UInt16Array, UInt32Array, UInt64Array,
    UInt8Array,
};
use arrow::datatypes::{DataType, Field, Schema, TimeUnit};
use arrow::record_batch::RecordBatch;
use parquet::arrow::arrow_reader::ParquetRecordBatchReaderBuilder;
use parquet::arrow::ArrowWriter;
use parquet::arrow::ProjectionMask;
use parquet::basic::Compression;
use parquet::file::properties::WriterProperties;
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
#[derive(Debug, Clone)]
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

/// Default matches `new()`: a derived Default would zero `batch_size` and
/// disable headers, silently diverging from the documented defaults.
impl Default for ParquetImportOptions {
    fn default() -> Self {
        Self {
            column_renames: HashMap::new(),
            include_headers: true,
            columns: Vec::new(),
            batch_size: 65536,
        }
    }
}

impl ParquetImportOptions {
    pub fn new() -> Self {
        Self::default()
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
        let all_column_names: Vec<String> =
            schema.fields().iter().map(|f| f.name().clone()).collect();

        // Determine which columns to import (in requested order); an unknown
        // name is an error rather than a silently dropped column
        let columns_to_import: Vec<usize> = if opts.columns.is_empty() {
            (0..all_column_names.len()).collect()
        } else {
            opts.columns
                .iter()
                .map(|name| {
                    all_column_names
                        .iter()
                        .position(|n| n == name)
                        .ok_or_else(|| {
                            RustypyxlError::ParseError(format!(
                                "Column '{}' not found in parquet file (available: {})",
                                name,
                                all_column_names.join(", ")
                            ))
                        })
                })
                .collect::<Result<_>>()?
        };

        if columns_to_import.is_empty() {
            return Err(RustypyxlError::ParseError(
                "No matching columns found in parquet file".to_string(),
            ));
        }

        // Push the selection down to the reader so unselected columns are
        // never decoded. The projected batch keeps file-schema order, so map
        // each requested column to its position within the projection.
        let mut projected: Vec<usize> = columns_to_import.clone();
        projected.sort_unstable();
        projected.dedup();
        let batch_indices: Vec<usize> = columns_to_import
            .iter()
            .map(|idx| projected.iter().position(|p| p == idx).unwrap())
            .collect();
        let projection = ProjectionMask::roots(builder.parquet_schema(), projected.iter().copied());

        // Build reader with batch size
        let reader = builder
            .with_batch_size(opts.batch_size)
            .with_projection(projection)
            .build()
            .map_err(|e| {
                RustypyxlError::ParseError(format!("Failed to build parquet reader: {}", e))
            })?;

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
                worksheet.set_cell_value(
                    current_row,
                    col,
                    CellValue::String(Arc::from(name.as_str())),
                );
            }
            current_row += 1;
        }

        let _data_start_row = current_row;
        let mut total_rows: u32 = 0;

        // Read batches and write to worksheet
        for batch_result in reader {
            let batch = batch_result.map_err(|e| {
                RustypyxlError::ParseError(format!("Failed to read parquet batch: {}", e))
            })?;

            let num_rows = batch.num_rows();

            // Process each column
            for (col_offset, &batch_idx) in batch_indices.iter().enumerate() {
                let col = start_col + col_offset as u32;
                let array = batch.column(batch_idx);

                write_arrow_array_to_worksheet(worksheet, array, current_row, col, num_rows);
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

/// Number formats applied to imported date/timestamp cells so Excel renders
/// them as dates instead of bare serial numbers. Interned once so that every
/// cell in a timestamp column shares one allocation.
const DATE_FORMAT: &str = "yyyy-mm-dd";
const DATETIME_FORMAT: &str = "yyyy-mm-dd hh:mm:ss";

fn date_format() -> &'static crate::cell::InternedString {
    static FORMAT: std::sync::OnceLock<crate::cell::InternedString> = std::sync::OnceLock::new();
    FORMAT.get_or_init(|| Arc::from(DATE_FORMAT))
}

fn datetime_format() -> &'static crate::cell::InternedString {
    static FORMAT: std::sync::OnceLock<crate::cell::InternedString> = std::sync::OnceLock::new();
    FORMAT.get_or_init(|| Arc::from(DATETIME_FORMAT))
}

/// Largest magnitude an f64 represents exactly. Excel stores every number as
/// an f64, so an integer beyond this cannot be held as a number without
/// silently rounding -- 9007199254740993 would come back as ...992. Keep the
/// exact digits as text instead: rounding an ID column is the worse failure.
const MAX_EXACT_INT: u64 = 1 << 53;

/// An i64 as a number when f64 holds it exactly, as text otherwise.
fn int64_cell_value(value: i64) -> CellValue {
    if value.unsigned_abs() <= MAX_EXACT_INT {
        CellValue::Number(value as f64)
    } else {
        CellValue::String(Arc::from(value.to_string()))
    }
}

/// A u64 as a number when f64 holds it exactly, as text otherwise.
fn uint64_cell_value(value: u64) -> CellValue {
    if value <= MAX_EXACT_INT {
        CellValue::Number(value as f64)
    } else {
        CellValue::String(Arc::from(value.to_string()))
    }
}

/// Write an Excel date serial with the matching number format.
///
/// The format is interned and cloned per cell: a million-row timestamp column
/// otherwise allocates the same short string a million times.
fn set_date_cell(
    worksheet: &mut Worksheet,
    row: u32,
    col: u32,
    serial: f64,
    format: &crate::cell::InternedString,
) {
    let cell = worksheet.get_or_create_cell_mut(row, col);
    cell.value = CellValue::Number(serial);
    cell.number_format = Some(format.clone());
    // Force re-resolution of the style at save time
    cell.style_index = None;
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
        DataType::Int8 => write_int_array(worksheet, array, start_row, col, num_rows),
        DataType::Int16 => write_int_array(worksheet, array, start_row, col, num_rows),
        DataType::Int32 => write_int_array(worksheet, array, start_row, col, num_rows),
        DataType::Int64 => write_int_array(worksheet, array, start_row, col, num_rows),
        DataType::UInt8 => write_uint_array(worksheet, array, start_row, col, num_rows),
        DataType::UInt16 => write_uint_array(worksheet, array, start_row, col, num_rows),
        DataType::UInt32 => write_uint_array(worksheet, array, start_row, col, num_rows),
        DataType::UInt64 => write_uint_array(worksheet, array, start_row, col, num_rows),
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
                    set_date_cell(worksheet, row, col, excel_serial as f64, date_format());
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
                    set_date_cell(worksheet, row, col, excel_serial, date_format());
                }
            }
        }
        // Timestamps are converted in UTC: Arrow stores tz-aware timestamps
        // as UTC instants and Excel serials have no timezone concept.
        DataType::Timestamp(unit, _tz) => match unit {
            TimeUnit::Second => {
                let arr = array
                    .as_any()
                    .downcast_ref::<TimestampSecondArray>()
                    .unwrap();
                for i in 0..num_rows {
                    let row = start_row + i as u32;
                    if arr.is_valid(i) {
                        let secs = arr.value(i) as f64;
                        let days = secs / (24.0 * 60.0 * 60.0);
                        let excel_serial = days + 25569.0;
                        set_date_cell(worksheet, row, col, excel_serial, datetime_format());
                    }
                }
            }
            TimeUnit::Millisecond => {
                let arr = array
                    .as_any()
                    .downcast_ref::<TimestampMillisecondArray>()
                    .unwrap();
                for i in 0..num_rows {
                    let row = start_row + i as u32;
                    if arr.is_valid(i) {
                        let ms = arr.value(i) as f64;
                        let days = ms / (24.0 * 60.0 * 60.0 * 1000.0);
                        let excel_serial = days + 25569.0;
                        set_date_cell(worksheet, row, col, excel_serial, datetime_format());
                    }
                }
            }
            TimeUnit::Microsecond => {
                let arr = array
                    .as_any()
                    .downcast_ref::<TimestampMicrosecondArray>()
                    .unwrap();
                for i in 0..num_rows {
                    let row = start_row + i as u32;
                    if arr.is_valid(i) {
                        let us = arr.value(i) as f64;
                        let days = us / (24.0 * 60.0 * 60.0 * 1_000_000.0);
                        let excel_serial = days + 25569.0;
                        set_date_cell(worksheet, row, col, excel_serial, datetime_format());
                    }
                }
            }
            TimeUnit::Nanosecond => {
                let arr = array
                    .as_any()
                    .downcast_ref::<TimestampNanosecondArray>()
                    .unwrap();
                for i in 0..num_rows {
                    let row = start_row + i as u32;
                    if arr.is_valid(i) {
                        let ns = arr.value(i) as f64;
                        let days = ns / (24.0 * 60.0 * 60.0 * 1_000_000_000.0);
                        let excel_serial = days + 25569.0;
                        set_date_cell(worksheet, row, col, excel_serial, datetime_format());
                    }
                }
            }
        },
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
                    // Go through the decimal string so values beyond i128
                    // keep the right magnitude and sign (truncating to the
                    // low 128 bits produced arbitrary wrong numbers); f64
                    // rounding is the only remaining loss.
                    let val = arr
                        .value(i)
                        .to_string()
                        .parse::<f64>()
                        .map(|v| v / scale_factor)
                        .unwrap_or(f64::NAN);
                    worksheet.set_cell_value(row, col, CellValue::Number(val));
                }
            }
        }
        // For other types, convert to string representation
        _ => {
            // One formatter per column, not per row
            let formatter = arrow::util::display::ArrayFormatter::try_new(
                array.as_ref(),
                &arrow::util::display::FormatOptions::default(),
            );
            if let Ok(fmt) = formatter {
                for i in 0..num_rows {
                    let row = start_row + i as u32;
                    if array.is_valid(i) {
                        let s = fmt.value(i).to_string();
                        worksheet.set_cell_value(row, col, CellValue::String(Arc::from(s)));
                    }
                }
            }
        }
    }
}

fn write_int_array(
    worksheet: &mut Worksheet,
    array: &ArrayRef,
    start_row: u32,
    col: u32,
    num_rows: usize,
) {
    // Use the primitive array trait for numeric types
    if let Some(arr) = array.as_any().downcast_ref::<Int8Array>() {
        for i in 0..num_rows {
            if arr.is_valid(i) {
                worksheet.set_cell_value(
                    start_row + i as u32,
                    col,
                    CellValue::Number(arr.value(i) as f64),
                );
            }
        }
    } else if let Some(arr) = array.as_any().downcast_ref::<Int16Array>() {
        for i in 0..num_rows {
            if arr.is_valid(i) {
                worksheet.set_cell_value(
                    start_row + i as u32,
                    col,
                    CellValue::Number(arr.value(i) as f64),
                );
            }
        }
    } else if let Some(arr) = array.as_any().downcast_ref::<Int32Array>() {
        for i in 0..num_rows {
            if arr.is_valid(i) {
                worksheet.set_cell_value(
                    start_row + i as u32,
                    col,
                    CellValue::Number(arr.value(i) as f64),
                );
            }
        }
    } else if let Some(arr) = array.as_any().downcast_ref::<Int64Array>() {
        for i in 0..num_rows {
            if arr.is_valid(i) {
                worksheet.set_cell_value(start_row + i as u32, col, int64_cell_value(arr.value(i)));
            }
        }
    }
}

fn write_uint_array(
    worksheet: &mut Worksheet,
    array: &ArrayRef,
    start_row: u32,
    col: u32,
    num_rows: usize,
) {
    if let Some(arr) = array.as_any().downcast_ref::<UInt8Array>() {
        for i in 0..num_rows {
            if arr.is_valid(i) {
                worksheet.set_cell_value(
                    start_row + i as u32,
                    col,
                    CellValue::Number(arr.value(i) as f64),
                );
            }
        }
    } else if let Some(arr) = array.as_any().downcast_ref::<UInt16Array>() {
        for i in 0..num_rows {
            if arr.is_valid(i) {
                worksheet.set_cell_value(
                    start_row + i as u32,
                    col,
                    CellValue::Number(arr.value(i) as f64),
                );
            }
        }
    } else if let Some(arr) = array.as_any().downcast_ref::<UInt32Array>() {
        for i in 0..num_rows {
            if arr.is_valid(i) {
                worksheet.set_cell_value(
                    start_row + i as u32,
                    col,
                    CellValue::Number(arr.value(i) as f64),
                );
            }
        }
    } else if let Some(arr) = array.as_any().downcast_ref::<UInt64Array>() {
        for i in 0..num_rows {
            if arr.is_valid(i) {
                worksheet.set_cell_value(
                    start_row + i as u32,
                    col,
                    uint64_cell_value(arr.value(i)),
                );
            }
        }
    }
}

// ============================================================================
// EXPORT FUNCTIONALITY
// ============================================================================

/// Result of a parquet export operation.
#[derive(Debug, Clone)]
pub struct ParquetExportResult {
    /// Number of rows exported (excluding header row if present).
    pub rows_exported: u32,
    /// Number of columns exported.
    pub columns_exported: u32,
    /// Column names as exported.
    pub column_names: Vec<String>,
    /// File size in bytes.
    pub file_size: u64,
}

/// Column type hint for parquet export.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ColumnType {
    /// Infer type from data (default).
    #[default]
    Auto,
    /// Force string type.
    String,
    /// Force float64 type.
    Float64,
    /// Force int64 type.
    Int64,
    /// Force boolean type.
    Boolean,
    /// Force date type (Excel serial → Date32).
    Date,
    /// Force datetime type (Excel serial → Timestamp).
    DateTime,
}

/// Options for parquet export.
#[derive(Debug, Clone)]
pub struct ParquetExportOptions {
    /// Column name mappings (original_name -> new_name).
    pub column_renames: HashMap<String, String>,
    /// Whether the first row contains headers. Default: true.
    pub has_headers: bool,
    /// Compression to use. Default: Snappy.
    pub compression: ParquetCompression,
    /// Column type hints (column_name -> type).
    pub column_types: HashMap<String, ColumnType>,
    /// Row group size. Default: 65536.
    pub row_group_size: usize,
}

/// Compression options for parquet export.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ParquetCompression {
    /// No compression.
    None,
    /// Snappy compression (default, good balance).
    #[default]
    Snappy,
    /// Gzip compression (better compression, slower).
    Gzip,
    /// Zstd compression (excellent compression and speed).
    Zstd,
    /// LZ4 compression (very fast, less compression).
    Lz4,
}

impl From<ParquetCompression> for Compression {
    fn from(c: ParquetCompression) -> Self {
        match c {
            ParquetCompression::None => Compression::UNCOMPRESSED,
            ParquetCompression::Snappy => Compression::SNAPPY,
            ParquetCompression::Gzip => Compression::GZIP(Default::default()),
            ParquetCompression::Zstd => Compression::ZSTD(Default::default()),
            ParquetCompression::Lz4 => Compression::LZ4,
        }
    }
}

/// Default matches `new()`: a derived Default would set `row_group_size`
/// to 0, which the parquet writer rejects, and disable headers.
impl Default for ParquetExportOptions {
    fn default() -> Self {
        Self {
            column_renames: HashMap::new(),
            has_headers: true,
            compression: ParquetCompression::default(),
            column_types: HashMap::new(),
            row_group_size: 65536,
        }
    }
}

impl ParquetExportOptions {
    pub fn new() -> Self {
        Self::default()
    }

    /// Set whether the first row contains headers.
    pub fn with_headers(mut self, has_headers: bool) -> Self {
        self.has_headers = has_headers;
        self
    }

    /// Add a column rename mapping.
    pub fn rename_column(mut self, from: &str, to: &str) -> Self {
        self.column_renames.insert(from.to_string(), to.to_string());
        self
    }

    /// Set compression type.
    pub fn with_compression(mut self, compression: ParquetCompression) -> Self {
        self.compression = compression;
        self
    }

    /// Set type hint for a column.
    pub fn with_column_type(mut self, column: &str, col_type: ColumnType) -> Self {
        self.column_types.insert(column.to_string(), col_type);
        self
    }

    /// Set row group size.
    pub fn with_row_group_size(mut self, size: usize) -> Self {
        self.row_group_size = size;
        self
    }
}

impl Workbook {
    /// Export a worksheet to a Parquet file.
    ///
    /// This exports cell data from the worksheet directly to Parquet format,
    /// with automatic type inference based on cell values.
    ///
    /// # Arguments
    /// * `sheet_name` - Name of the worksheet to export
    /// * `path` - Output path for the Parquet file
    /// * `options` - Export options (headers, compression, etc.)
    ///
    /// # Returns
    /// Information about what was exported.
    ///
    /// # Example
    /// ```no_run
    /// use rustypyxl::{Workbook, parquet_import::{ParquetExportOptions, ParquetCompression}};
    ///
    /// let wb = Workbook::load("data.xlsx").unwrap();
    /// let result = wb.export_to_parquet(
    ///     "Sheet1",
    ///     "output.parquet",
    ///     None,
    /// ).unwrap();
    /// println!("Exported {} rows", result.rows_exported);
    /// ```
    pub fn export_to_parquet(
        &self,
        sheet_name: &str,
        path: &str,
        options: Option<ParquetExportOptions>,
    ) -> Result<ParquetExportResult> {
        let worksheet = self.get_sheet_by_name(sheet_name)?;
        let (min_row, min_col, max_row, max_col) = worksheet.dimensions();
        if max_row < min_row || max_col < min_col {
            return Err(RustypyxlError::custom("Worksheet is empty"));
        }
        self.export_cells(
            sheet_name,
            path,
            min_row,
            min_col,
            max_row,
            max_col,
            options.unwrap_or_default(),
        )
    }

    /// Export a specific range from a worksheet to a Parquet file.
    ///
    /// # Arguments
    /// * `sheet_name` - Name of the worksheet to export
    /// * `path` - Output path for the Parquet file
    /// * `min_row` - Starting row (1-indexed)
    /// * `min_col` - Starting column (1-indexed)
    /// * `max_row` - Ending row (1-indexed)
    /// * `max_col` - Ending column (1-indexed)
    /// * `options` - Export options
    #[allow(clippy::too_many_arguments)]
    pub fn export_range_to_parquet(
        &self,
        sheet_name: &str,
        path: &str,
        min_row: u32,
        min_col: u32,
        max_row: u32,
        max_col: u32,
        options: Option<ParquetExportOptions>,
    ) -> Result<ParquetExportResult> {
        if max_row < min_row || max_col < min_col {
            return Err(RustypyxlError::custom("Invalid range"));
        }
        self.export_cells(
            sheet_name,
            path,
            min_row,
            min_col,
            max_row,
            max_col,
            options.unwrap_or_default(),
        )
    }

    /// Shared export implementation. Rows are written in row_group_size
    /// chunks so peak memory is bounded by one chunk instead of the whole
    /// sheet; column types are inferred in a first streaming pass so every
    /// chunk shares the same schema.
    #[allow(clippy::too_many_arguments)]
    fn export_cells(
        &self,
        sheet_name: &str,
        path: &str,
        min_row: u32,
        min_col: u32,
        max_row: u32,
        max_col: u32,
        options: ParquetExportOptions,
    ) -> Result<ParquetExportResult> {
        let worksheet = self.get_sheet_by_name(sheet_name)?;

        let num_cols = (max_col - min_col + 1) as usize;
        let data_start_row = if options.has_headers {
            min_row + 1
        } else {
            min_row
        };
        let num_data_rows = if max_row >= data_start_row {
            (max_row - data_start_row + 1) as usize
        } else {
            0
        };

        // Extract column names
        let column_names: Vec<String> = if options.has_headers {
            (min_col..=max_col)
                .map(|col| {
                    let original = worksheet
                        .get_cell_value(min_row, col)
                        .map(|v| v.to_string())
                        .unwrap_or_else(|| format!("Column{}", col - min_col + 1));
                    options
                        .column_renames
                        .get(&original)
                        .cloned()
                        .unwrap_or(original)
                })
                .collect()
        } else {
            (min_col..=max_col)
                .map(|col| format!("Column{}", col - min_col + 1))
                .collect()
        };

        // Pass 1: resolve each column's type, materializing one column at a
        // time at most (only for columns without an explicit hint)
        let resolved_types: Vec<ColumnType> = column_names
            .iter()
            .enumerate()
            .map(|(col_idx, col_name)| {
                let hint = options
                    .column_types
                    .get(col_name)
                    .copied()
                    .unwrap_or(ColumnType::Auto);
                if hint != ColumnType::Auto || num_data_rows == 0 {
                    return hint;
                }
                let col = min_col + col_idx as u32;
                let col_data: Vec<Option<&CellValue>> = (data_start_row..=max_row)
                    .map(|row| worksheet.get_cell_value(row, col))
                    .collect();
                infer_column_type(&col_data)
            })
            .collect();

        let fields: Vec<Field> = column_names
            .iter()
            .zip(&resolved_types)
            .map(|(name, col_type)| {
                // Field metadata must match what build_arrow_column produces
                let (field, _) = build_arrow_column(name, &[], *col_type);
                field
            })
            .collect();
        let schema = Arc::new(Schema::new(fields));

        // Write to parquet
        let file = File::create(path)
            .map_err(|e| RustypyxlError::custom(format!("Failed to create file: {}", e)))?;

        let props = WriterProperties::builder()
            .set_compression(options.compression.into())
            .set_max_row_group_size(options.row_group_size)
            .build();

        let mut writer = ArrowWriter::try_new(file, schema.clone(), Some(props)).map_err(|e| {
            RustypyxlError::custom(format!("Failed to create parquet writer: {}", e))
        })?;

        // Pass 2: build and write one RecordBatch per row-group-sized chunk
        let chunk_rows = options.row_group_size.max(1) as u32;
        let mut chunk_start = data_start_row;
        while chunk_start <= max_row && num_data_rows > 0 {
            let chunk_end = chunk_start.saturating_add(chunk_rows - 1).min(max_row);

            let mut arrays: Vec<ArrayRef> = Vec::with_capacity(num_cols);
            for (col_idx, col_name) in column_names.iter().enumerate() {
                let col = min_col + col_idx as u32;
                let col_data: Vec<Option<&CellValue>> = (chunk_start..=chunk_end)
                    .map(|row| worksheet.get_cell_value(row, col))
                    .collect();
                let (_, array) = build_arrow_column(col_name, &col_data, resolved_types[col_idx]);
                arrays.push(array);
            }

            let batch = RecordBatch::try_new(schema.clone(), arrays).map_err(|e| {
                RustypyxlError::custom(format!("Failed to create record batch: {}", e))
            })?;
            writer
                .write(&batch)
                .map_err(|e| RustypyxlError::custom(format!("Failed to write batch: {}", e)))?;

            chunk_start = chunk_end + 1;
        }

        writer
            .close()
            .map_err(|e| RustypyxlError::custom(format!("Failed to close writer: {}", e)))?;

        // Get file size
        let file_size = std::fs::metadata(path).map(|m| m.len()).unwrap_or(0);

        Ok(ParquetExportResult {
            rows_exported: num_data_rows as u32,
            columns_exported: num_cols as u32,
            column_names,
            file_size,
        })
    }
}

/// Infer column type from cell values.
fn infer_column_type(values: &[Option<&CellValue>]) -> ColumnType {
    let mut has_string = false;
    let mut has_number = false;
    let mut has_boolean = false;
    let mut all_integers = true;

    for value in values.iter().flatten() {
        match value {
            CellValue::String(_) | CellValue::Formula(_) | CellValue::Date(_) => {
                has_string = true;
            }
            CellValue::Number(n) => {
                has_number = true;
                if n.fract() != 0.0 {
                    all_integers = false;
                }
            }
            CellValue::Boolean(_) => {
                has_boolean = true;
            }
            CellValue::Empty => {}
        }
    }

    // Priority: if any strings, use string; otherwise prefer numbers
    if has_string {
        ColumnType::String
    } else if has_number {
        if all_integers {
            ColumnType::Int64
        } else {
            ColumnType::Float64
        }
    } else if has_boolean {
        ColumnType::Boolean
    } else {
        ColumnType::String // default for empty columns
    }
}

/// Build an Arrow column from cell values.
fn build_arrow_column(
    name: &str,
    values: &[Option<&CellValue>],
    type_hint: ColumnType,
) -> (Field, ArrayRef) {
    let col_type = if type_hint == ColumnType::Auto {
        infer_column_type(values)
    } else {
        type_hint
    };

    match col_type {
        ColumnType::String | ColumnType::Auto => {
            let arr: StringArray = values.iter().map(|v| v.map(|cv| cv.to_string())).collect();
            (
                Field::new(name, DataType::Utf8, true),
                Arc::new(arr) as ArrayRef,
            )
        }
        ColumnType::Float64 => {
            let arr: Float64Array = values
                .iter()
                .map(|v| v.and_then(cell_value_to_f64))
                .collect();
            (
                Field::new(name, DataType::Float64, true),
                Arc::new(arr) as ArrayRef,
            )
        }
        ColumnType::Int64 => {
            let arr: Int64Array = values
                .iter()
                .map(|v| v.and_then(cell_value_to_i64))
                .collect();
            (
                Field::new(name, DataType::Int64, true),
                Arc::new(arr) as ArrayRef,
            )
        }
        ColumnType::Boolean => {
            let arr: BooleanArray = values
                .iter()
                .map(|v| v.and_then(cell_value_to_bool))
                .collect();
            (
                Field::new(name, DataType::Boolean, true),
                Arc::new(arr) as ArrayRef,
            )
        }
        ColumnType::Date => {
            // Excel serial number to days since Unix epoch
            let arr: Date32Array = values
                .iter()
                .map(|v| v.and_then(cell_value_to_date32))
                .collect();
            (
                Field::new(name, DataType::Date32, true),
                Arc::new(arr) as ArrayRef,
            )
        }
        ColumnType::DateTime => {
            // Excel serial number to milliseconds since Unix epoch
            let arr: TimestampMillisecondArray = values
                .iter()
                .map(|v| v.and_then(cell_value_to_timestamp_ms))
                .collect();
            (
                Field::new(name, DataType::Timestamp(TimeUnit::Millisecond, None), true),
                Arc::new(arr) as ArrayRef,
            )
        }
    }
}

fn cell_value_to_f64(value: &CellValue) -> Option<f64> {
    match value {
        CellValue::Number(n) => Some(*n),
        CellValue::Boolean(b) => Some(if *b { 1.0 } else { 0.0 }),
        CellValue::String(s) => s.parse().ok(),
        CellValue::Formula(s) => s.parse().ok(),
        _ => None,
    }
}

fn cell_value_to_i64(value: &CellValue) -> Option<i64> {
    match value {
        CellValue::Number(n) => Some(*n as i64),
        CellValue::Boolean(b) => Some(if *b { 1 } else { 0 }),
        CellValue::String(s) => s.parse().ok(),
        CellValue::Formula(s) => s.parse().ok(),
        _ => None,
    }
}

fn cell_value_to_bool(value: &CellValue) -> Option<bool> {
    match value {
        CellValue::Boolean(b) => Some(*b),
        CellValue::Number(n) => Some(*n != 0.0),
        CellValue::String(s) => {
            let lower = s.to_lowercase();
            if lower == "true" || lower == "yes" || lower == "1" {
                Some(true)
            } else if lower == "false" || lower == "no" || lower == "0" {
                Some(false)
            } else {
                None
            }
        }
        _ => None,
    }
}

fn cell_value_to_date32(value: &CellValue) -> Option<i32> {
    match value {
        CellValue::Number(n) => {
            // Excel serial to days since Unix epoch
            // Excel epoch is 1900-01-01 (serial 1), but with 1900 leap year bug
            // Unix epoch (1970-01-01) is Excel serial 25569
            Some((*n as i32) - 25569)
        }
        _ => None,
    }
}

fn cell_value_to_timestamp_ms(value: &CellValue) -> Option<i64> {
    match value {
        CellValue::Number(n) => {
            // Excel serial to milliseconds since Unix epoch
            // Days since Unix epoch, then convert to ms
            let days_since_unix = *n - 25569.0;
            let ms = days_since_unix * 24.0 * 60.0 * 60.0 * 1000.0;
            Some(ms as i64)
        }
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::NamedTempFile;

    /// Write a single-column parquet file and import it into a fresh workbook.
    fn import_column(field: Field, array: ArrayRef) -> Workbook {
        let schema = Arc::new(Schema::new(vec![field]));
        let batch = RecordBatch::try_new(schema.clone(), vec![array]).unwrap();

        let file = NamedTempFile::new().unwrap();
        let mut writer = ArrowWriter::try_new(file.reopen().unwrap(), schema, None).unwrap();
        writer.write(&batch).unwrap();
        writer.close().unwrap();

        let mut wb = Workbook::new();
        wb.create_sheet(Some("Data".to_string())).unwrap();
        wb.insert_from_parquet("Data", file.path().to_str().unwrap(), 1, 1, None)
            .unwrap();
        wb
    }

    /// An i64 past 2^53 cannot be held in an f64: 9007199254740993 would come
    /// back as ...992. Keep the exact digits as text rather than corrupting an
    /// ID column.
    #[test]
    fn test_int64_beyond_f64_precision_is_kept_exact() {
        let values = Int64Array::from(vec![
            42i64,
            9_007_199_254_740_992,  // 2^53, still exact
            9_007_199_254_740_993,  // 2^53 + 1, not representable
            -9_007_199_254_740_993, // and the negative side
        ]);
        let wb = import_column(
            Field::new("id", DataType::Int64, false),
            Arc::new(values) as ArrayRef,
        );
        let ws = wb.get_sheet_by_name("Data").unwrap();

        // Row 1 is the header
        assert_eq!(ws.get_cell_value(2, 1), Some(&CellValue::Number(42.0)));
        assert_eq!(
            ws.get_cell_value(3, 1),
            Some(&CellValue::Number(9_007_199_254_740_992.0))
        );
        assert_eq!(
            ws.get_cell_value(4, 1),
            Some(&CellValue::String(Arc::from("9007199254740993"))),
            "a value past 2^53 must keep its exact digits"
        );
        assert_eq!(
            ws.get_cell_value(5, 1),
            Some(&CellValue::String(Arc::from("-9007199254740993")))
        );
    }

    /// Same for the unsigned side, where the gap opens at the same magnitude.
    #[test]
    fn test_uint64_beyond_f64_precision_is_kept_exact() {
        let values = UInt64Array::from(vec![
            7u64,
            9_007_199_254_740_993,
            u64::MAX, // 18446744073709551615
        ]);
        let wb = import_column(
            Field::new("id", DataType::UInt64, false),
            Arc::new(values) as ArrayRef,
        );
        let ws = wb.get_sheet_by_name("Data").unwrap();

        assert_eq!(ws.get_cell_value(2, 1), Some(&CellValue::Number(7.0)));
        assert_eq!(
            ws.get_cell_value(3, 1),
            Some(&CellValue::String(Arc::from("9007199254740993")))
        );
        assert_eq!(
            ws.get_cell_value(4, 1),
            Some(&CellValue::String(Arc::from("18446744073709551615")))
        );
    }

    /// Int32 columns are always exact in an f64 and stay numbers.
    #[test]
    fn test_int32_stays_numeric() {
        let values = Int32Array::from(vec![i32::MIN, 0, i32::MAX]);
        let wb = import_column(
            Field::new("n", DataType::Int32, false),
            Arc::new(values) as ArrayRef,
        );
        let ws = wb.get_sheet_by_name("Data").unwrap();

        assert_eq!(
            ws.get_cell_value(2, 1),
            Some(&CellValue::Number(i32::MIN as f64))
        );
        assert_eq!(
            ws.get_cell_value(4, 1),
            Some(&CellValue::Number(i32::MAX as f64))
        );
    }

    #[test]
    fn test_import_options_builder() {
        let opts = ParquetImportOptions::new()
            .rename_column("old_name", "new_name")
            .with_headers(true)
            .select_columns(vec!["col1".to_string(), "col2".to_string()])
            .with_batch_size(1000);

        assert_eq!(
            opts.column_renames.get("old_name"),
            Some(&"new_name".to_string())
        );
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

    #[test]
    fn test_export_options_builder() {
        let opts = ParquetExportOptions::new()
            .rename_column("old_name", "new_name")
            .with_headers(true)
            .with_compression(ParquetCompression::Zstd)
            .with_column_type("numbers", ColumnType::Float64)
            .with_row_group_size(10000);

        assert_eq!(
            opts.column_renames.get("old_name"),
            Some(&"new_name".to_string())
        );
        assert!(opts.has_headers);
        assert_eq!(opts.compression, ParquetCompression::Zstd);
        assert_eq!(opts.column_types.get("numbers"), Some(&ColumnType::Float64));
        assert_eq!(opts.row_group_size, 10000);
    }

    #[test]
    fn test_infer_column_type_numbers() {
        let v1 = CellValue::Number(1.0);
        let v2 = CellValue::Number(2.0);
        let v3 = CellValue::Number(3.0);
        let values: Vec<Option<&CellValue>> = vec![Some(&v1), Some(&v2), Some(&v3)];
        assert_eq!(infer_column_type(&values), ColumnType::Int64);

        let v4 = CellValue::Number(1.5);
        let values2: Vec<Option<&CellValue>> = vec![Some(&v1), Some(&v4)];
        assert_eq!(infer_column_type(&values2), ColumnType::Float64);
    }

    #[test]
    fn test_infer_column_type_strings() {
        let v1 = CellValue::String(Arc::from("hello"));
        let v2 = CellValue::Number(42.0);
        let values: Vec<Option<&CellValue>> = vec![Some(&v1), Some(&v2)];
        assert_eq!(infer_column_type(&values), ColumnType::String);
    }

    #[test]
    fn test_infer_column_type_booleans() {
        let v1 = CellValue::Boolean(true);
        let v2 = CellValue::Boolean(false);
        let values: Vec<Option<&CellValue>> = vec![Some(&v1), Some(&v2)];
        assert_eq!(infer_column_type(&values), ColumnType::Boolean);
    }

    #[test]
    fn test_export_chunks_rows_into_row_groups() {
        let mut wb = Workbook::new();
        wb.create_sheet(Some("Big".to_string())).unwrap();
        {
            let ws = wb.get_sheet_by_name_mut("Big").unwrap();
            ws.set_cell_value(1, 1, CellValue::from("n"));
            for r in 0..100u32 {
                ws.set_cell_value(r + 2, 1, CellValue::Number(r as f64));
            }
        }

        let dir = std::env::temp_dir().join("rustypyxl_parquet_tests");
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("chunked.parquet");
        let path_str = path.to_str().unwrap();

        let opts = ParquetExportOptions {
            row_group_size: 10,
            ..Default::default()
        };
        let result = wb.export_to_parquet("Big", path_str, Some(opts)).unwrap();
        assert_eq!(result.rows_exported, 100);

        // One batch per chunk means one row group per 10 rows, and the data
        // must still round-trip in order
        let file = File::open(&path).unwrap();
        let builder = ParquetRecordBatchReaderBuilder::try_new(file).unwrap();
        assert_eq!(builder.metadata().num_row_groups(), 10);
        let reader = builder.build().unwrap();
        let mut values = Vec::new();
        for batch in reader {
            let batch = batch.unwrap();
            // Whole numbers are inferred as Int64 on export
            let arr = batch
                .column(0)
                .as_any()
                .downcast_ref::<Int64Array>()
                .unwrap();
            for i in 0..arr.len() {
                values.push(arr.value(i) as f64);
            }
        }
        assert_eq!(values.len(), 100);
        assert_eq!(values[0], 0.0);
        assert_eq!(values[99], 99.0);

        std::fs::remove_file(&path).ok();
    }

    #[test]
    fn test_export_roundtrip() {
        // Create a workbook with test data
        let mut wb = Workbook::new();
        wb.create_sheet(Some("TestSheet".to_string())).unwrap();

        // Set header row
        wb.set_cell_value_in_sheet("TestSheet", 1, 1, CellValue::String(Arc::from("Name")))
            .unwrap();
        wb.set_cell_value_in_sheet("TestSheet", 1, 2, CellValue::String(Arc::from("Age")))
            .unwrap();
        wb.set_cell_value_in_sheet("TestSheet", 1, 3, CellValue::String(Arc::from("Score")))
            .unwrap();

        // Set data rows
        wb.set_cell_value_in_sheet("TestSheet", 2, 1, CellValue::String(Arc::from("Alice")))
            .unwrap();
        wb.set_cell_value_in_sheet("TestSheet", 2, 2, CellValue::Number(30.0))
            .unwrap();
        wb.set_cell_value_in_sheet("TestSheet", 2, 3, CellValue::Number(95.5))
            .unwrap();

        wb.set_cell_value_in_sheet("TestSheet", 3, 1, CellValue::String(Arc::from("Bob")))
            .unwrap();
        wb.set_cell_value_in_sheet("TestSheet", 3, 2, CellValue::Number(25.0))
            .unwrap();
        wb.set_cell_value_in_sheet("TestSheet", 3, 3, CellValue::Number(87.3))
            .unwrap();

        // Export to parquet
        let temp = NamedTempFile::new().unwrap();
        let path = temp.path().to_str().unwrap();

        let result = wb.export_to_parquet("TestSheet", path, None).unwrap();

        assert_eq!(result.rows_exported, 2);
        assert_eq!(result.columns_exported, 3);
        assert_eq!(result.column_names, vec!["Name", "Age", "Score"]);
        assert!(result.file_size > 0);

        // Import back
        let mut wb2 = Workbook::new();
        wb2.create_sheet(Some("Imported".to_string())).unwrap();

        let import_result = wb2
            .insert_from_parquet("Imported", path, 1, 1, None)
            .unwrap();

        assert_eq!(import_result.rows_imported, 2);
        assert_eq!(import_result.columns_imported, 3);

        // Verify data
        let ws = wb2.get_sheet_by_name("Imported").unwrap();
        assert_eq!(
            ws.get_cell_value(1, 1),
            Some(&CellValue::String(Arc::from("Name")))
        );
        assert_eq!(
            ws.get_cell_value(2, 1),
            Some(&CellValue::String(Arc::from("Alice")))
        );
        assert_eq!(
            ws.get_cell_value(3, 1),
            Some(&CellValue::String(Arc::from("Bob")))
        );
    }

    #[test]
    fn test_parquet_roundtrip_parquet_to_sheet_to_parquet() {
        // This tests: parquet -> sheet -> parquet -> sheet -> verify
        //
        // 1. Create a source parquet file
        // 2. Import to worksheet
        // 3. Export back to parquet
        // 4. Import that parquet to another sheet
        // 5. Verify data matches

        // Step 1: Create source parquet file
        use arrow::datatypes::Schema;
        use arrow::record_batch::RecordBatch;
        use parquet::arrow::ArrowWriter;

        let temp_parquet1 = NamedTempFile::new().unwrap();
        let temp_parquet2 = NamedTempFile::new().unwrap();
        let path1 = temp_parquet1.path().to_str().unwrap();
        let path2 = temp_parquet2.path().to_str().unwrap();

        // Create test data in parquet format
        let schema = Arc::new(Schema::new(vec![
            Field::new("id", DataType::Int64, false),
            Field::new("name", DataType::Utf8, true),
            Field::new("value", DataType::Float64, true),
            Field::new("active", DataType::Boolean, true),
        ]));

        let id_array = Int64Array::from(vec![1, 2, 3, 4, 5]);
        let name_array = StringArray::from(vec![
            Some("Alice"),
            Some("Bob"),
            Some("Charlie"),
            None,
            Some("Eve"),
        ]);
        let value_array = Float64Array::from(vec![
            Some(100.5),
            Some(200.0),
            None,
            Some(400.25),
            Some(500.75),
        ]);
        let active_array =
            BooleanArray::from(vec![Some(true), Some(false), Some(true), None, Some(false)]);

        let batch = RecordBatch::try_new(
            schema.clone(),
            vec![
                Arc::new(id_array),
                Arc::new(name_array),
                Arc::new(value_array),
                Arc::new(active_array),
            ],
        )
        .unwrap();

        let file = File::create(path1).unwrap();
        let mut writer = ArrowWriter::try_new(file, schema, None).unwrap();
        writer.write(&batch).unwrap();
        writer.close().unwrap();

        // Step 2: Import parquet to worksheet
        let mut wb = Workbook::new();
        wb.create_sheet(Some("Data".to_string())).unwrap();

        let import_result = wb.insert_from_parquet("Data", path1, 1, 1, None).unwrap();
        assert_eq!(import_result.rows_imported, 5);
        assert_eq!(import_result.columns_imported, 4);

        // Step 3: Export worksheet to new parquet
        let export_result = wb.export_to_parquet("Data", path2, None).unwrap();
        assert_eq!(export_result.rows_exported, 5);
        assert_eq!(export_result.columns_exported, 4);

        // Step 4: Import new parquet to another worksheet
        let mut wb2 = Workbook::new();
        wb2.create_sheet(Some("Imported".to_string())).unwrap();

        let import_result2 = wb2
            .insert_from_parquet("Imported", path2, 1, 1, None)
            .unwrap();
        assert_eq!(import_result2.rows_imported, 5);
        assert_eq!(import_result2.columns_imported, 4);

        // Step 5: Verify data matches
        let ws1 = wb.get_sheet_by_name("Data").unwrap();
        let ws2 = wb2.get_sheet_by_name("Imported").unwrap();

        // Check headers
        assert_eq!(
            ws1.get_cell_value(1, 1).map(|v| v.to_string()),
            ws2.get_cell_value(1, 1).map(|v| v.to_string())
        );
        assert_eq!(
            ws1.get_cell_value(1, 2).map(|v| v.to_string()),
            ws2.get_cell_value(1, 2).map(|v| v.to_string())
        );
        assert_eq!(
            ws1.get_cell_value(1, 3).map(|v| v.to_string()),
            ws2.get_cell_value(1, 3).map(|v| v.to_string())
        );
        assert_eq!(
            ws1.get_cell_value(1, 4).map(|v| v.to_string()),
            ws2.get_cell_value(1, 4).map(|v| v.to_string())
        );

        // Check data rows
        for row in 2..=6 {
            for col in 1..=4 {
                let v1 = ws1.get_cell_value(row, col).map(|v| v.to_string());
                let v2 = ws2.get_cell_value(row, col).map(|v| v.to_string());
                assert_eq!(v1, v2, "Mismatch at row {} col {}", row, col);
            }
        }
    }

    #[test]
    fn test_parquet_compression_options() {
        let temp = NamedTempFile::new().unwrap();
        let path = temp.path().to_str().unwrap();

        let mut wb = Workbook::new();
        wb.create_sheet(Some("Data".to_string())).unwrap();
        wb.set_cell_value_in_sheet("Data", 1, 1, CellValue::String(Arc::from("Col1")))
            .unwrap();
        wb.set_cell_value_in_sheet("Data", 2, 1, CellValue::Number(42.0))
            .unwrap();

        // Test different compression options
        let opts_zstd = ParquetExportOptions::new().with_compression(ParquetCompression::Zstd);
        let result = wb.export_to_parquet("Data", path, Some(opts_zstd)).unwrap();
        assert!(result.file_size > 0);

        let opts_none = ParquetExportOptions::new().with_compression(ParquetCompression::None);
        let result = wb.export_to_parquet("Data", path, Some(opts_none)).unwrap();
        assert!(result.file_size > 0);
    }

    #[test]
    fn test_parquet_column_type_hints() {
        let temp = NamedTempFile::new().unwrap();
        let path = temp.path().to_str().unwrap();

        let mut wb = Workbook::new();
        wb.create_sheet(Some("Data".to_string())).unwrap();

        // Create data with mixed types that could be interpreted differently
        wb.set_cell_value_in_sheet("Data", 1, 1, CellValue::String(Arc::from("Value")))
            .unwrap();
        wb.set_cell_value_in_sheet("Data", 2, 1, CellValue::Number(1.0))
            .unwrap();
        wb.set_cell_value_in_sheet("Data", 3, 1, CellValue::Number(2.0))
            .unwrap();

        // Force it to be exported as float64 even though values are integers
        let opts = ParquetExportOptions::new().with_column_type("Value", ColumnType::Float64);

        let result = wb.export_to_parquet("Data", path, Some(opts)).unwrap();
        assert_eq!(result.rows_exported, 2);
        assert!(result.file_size > 0);
    }
}

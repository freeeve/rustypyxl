# Parquet import: Int64 precision loss and per-row overhead

rustypyxl-core/src/parquet_import.rs.

- Correctness: Int64/UInt64 columns are cast straight to f64
  (`CellValue::Number(arr.value(i) as f64)`, parquet_import.rs:578-588 and
  :628-638). Values above 2^53 silently lose precision: importing an ID
  column containing 9007199254740993 stores 9007199254740992, and
  export_to_parquet writes the corrupted value back. Decimal256 already
  goes through string parsing to avoid truncation--do the same for i64/u64
  values outside the f64-exact range (store as string, or document and
  error). Silent corruption of high-magnitude integer keys is the worst
  failure mode.

- Performance: the fallback conversion branch constructs an ArrayFormatter
  per ROW (`ArrayFormatter::try_new(...)` inside `for i in 0..num_rows`,
  parquet_import.rs:522-536). Hoist it above the loop--one formatter per
  column.

- Performance (minor): set_date_cell allocates a fresh String per
  date/timestamp cell for a number_format that is always one of two
  `&'static str` constants (parquet_import.rs:316-322). 1M-row timestamp
  column -> 1M identical allocations. Intern or Arc the two formats.

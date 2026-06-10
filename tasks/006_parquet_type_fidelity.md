# Parquet import/export type fidelity

- Date32/Date64/Timestamp import as bare serial numbers with no date number
  format applied, so users see `45123` instead of a date; timezones are
  silently discarded (parquet_import.rs:349-427).
- Decimal256 truncates to the low 128 bits — values outside i128 range
  become wrong numbers (including sign flips), not just rounding
  (parquet_import.rs:447-448).
- `select_columns` silently drops unknown column names (errors only when all
  are missing) and pushes no projection mask to the reader, so all columns
  are decoded even when a subset is requested.
- Export materializes the whole sheet as a dense rows×cols Vec plus a single
  RecordBatch; `row_group_size` doesn't bound peak memory.
- `export_to_parquet` and `export_range_to_parquet` are ~110 duplicated
  lines.

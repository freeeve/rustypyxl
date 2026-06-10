# Parquet import/export type fidelity — DONE

- Date32/Date64/Timestamp import with date number formats applied
  (yyyy-mm-dd / yyyy-mm-dd hh:mm:ss); openpyxl reads them back as datetime
  objects. Timestamps convert in UTC (documented; Arrow tz-aware timestamps
  are UTC instants and Excel serials carry no timezone).
- Decimal256 converts via the decimal string: values beyond i128 (e.g.
  10^40) keep correct magnitude and sign; previously the low-128-bit
  truncation produced arbitrary numbers.
- select_columns errors on unknown names (listing available columns),
  pushes a ProjectionMask to the parquet reader so unselected columns are
  never decoded, and honors the requested column order.
- Export unified into one implementation writing one RecordBatch per
  row_group_size chunk - peak memory bounded by a chunk, not the sheet.
  Verified: 100 rows with row_group_size=10 produce 10 row groups with
  data intact.

Verified by tests/test_parquet.py::TestParquetTypeFidelity (5 pyarrow
differential tests) and a Rust chunked-export row-group test.

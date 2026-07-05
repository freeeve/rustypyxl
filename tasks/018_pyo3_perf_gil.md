# PyO3 perf: streaming append_row holds GIL; values_only per-cell overhead

- Streaming `append_row` holds the GIL across row-XML generation, deflate
  compression, and disk I/O (rustypyxl-pyo3/src/streaming.rs:64-84).
  After the Python->CellValue conversion (which needs the GIL), the
  `wb.append_row(sheet, cell_values)` call is pure Rust; wrap it in
  `py.allow_threads(|| ...)` like save/save_to_bytes/insert_from_parquet
  already do. For the intended million-row streaming use case this blocks
  all other Python threads for the duration of every row write.

- `iter_rows(values_only=True)` re-borrows the workbook and does a linear
  sheet_index_by_uid scan per CELL (rustypyxl-pyo3/src/worksheet.rs:520-534
  `read_value`, driven by `__next__` at :557-581; core scan at
  rustypyxl-core/src/workbook.rs:255-260). Resolve the sheet index once per
  row (or per iterator) instead.

Verify with benchmarks/ scripts (streaming write throughput, iter_rows
values_only read).

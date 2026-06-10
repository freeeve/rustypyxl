# Streaming workbook cannot write more than one sheet — DONE

- create_sheet finalizes the currently open sheet, making the existing
  multi-sheet workbook.xml plumbing reachable; verified by a round-trip
  test loading both sheets back.
- StreamingSheet handles carry their index; append_row on a stale or
  foreign handle errors instead of writing into whichever sheet is open.
- Excel row/column limits enforced in append_row; sheet names validated
  (1-31 chars, no []:*?/\\, no duplicates).
- finish() closes without a sheet handle; zero-sheet workbooks get a
  default "Sheet1". Python close() works at any point.
- WriteOnlyWorkbook supports `with` blocks (__enter__/__exit__), closing
  on exit without masking propagating exceptions.

Verified by 4 new Rust streaming tests and 5 new Python tests
(tests/test_streaming.py::TestStreamingMultiSheet).

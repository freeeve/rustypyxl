# PyO3 bindings: active-sheet tracking and conversion gaps

Correctness issues in rustypyxl-pyo3 (openpyxl-compat divergences).

- `wb.active` points at the wrong sheet after `remove()`. Core
  `remove_sheet` (rustypyxl-core/src/workbook.rs:263-272) deletes from
  worksheets/sheet_names but never adjusts `active_sheet`; the Python
  `active` getter only clamps to len-1 (rustypyxl-pyo3/src/workbook.rs:80-94).
  With sheets [A,B,C] and active==1 (B), `wb.remove(wb["A"])` leaves
  active==1 which is now C; a subsequent write to `wb.active` targets the
  wrong sheet. Decrement `active_sheet` when a sheet at or before it is
  removed (match openpyxl behavior for removing the active sheet itself).

- `coerce_color` silently drops theme-only/indexed-only colors: it returns
  only `color.rgb` from a PyColor, ignoring theme/indexed/tint
  (rustypyxl-pyo3/src/style.rs:9-23). `Font(color=Color(theme=1))` produces
  a font with no color. Support the other variants or raise instead of
  silently discarding.

- `cell.data_type` never returns 'd' for datetime cells
  (rustypyxl-pyo3/src/cell.rs:354-371): datetime values fall through the
  String/bool/f64 extraction and report "s". openpyxl returns 'd'.

- Minor re-entrancy hazard: write_rows (workbook.rs:428-450) and
  `__setitem__` (worksheet.rs:210-227) run `python_to_cell_value`'s `.str()`
  fallback (arbitrary `__str__`) while holding borrow_mut on the PyWorkbook;
  touching the same workbook from `__str__` raises "Already borrowed".
  `append` already collects values before borrowing--do the same in the
  other two.

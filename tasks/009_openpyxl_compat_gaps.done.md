# Remaining openpyxl API compatibility gaps (pyo3) — DONE

- iter_rows/iter_cols are lazy iterators of tuples (CellRangeIterator),
  resolving the sheet by uid per step; next() works; no up-front
  materialization. Dense iteration over the used range matches openpyxl's
  own semantics.
- PatternFill(start_color=..., end_color=...) supported with alias getters;
  Font/Side/PatternFill color params accept Color objects or rgb strings.
- Integral cells read back as int (exact within 2^53).
- ws.append takes any iterable or a dict keyed by column letter/index.
- cell.number_format = None clears (done earlier in task 002).
- wb.active honors the loaded active tab (done earlier in task 001).
- GC traverse/clear on Cell/Worksheet/CellRangeIterator - cycles through
  workbook references are collectable.
- Type stubs ship in the wheel (rustypyxl/__init__.pyi + py.typed).

Verified by tests/test_openpyxl_api_compat.py (17 tests).

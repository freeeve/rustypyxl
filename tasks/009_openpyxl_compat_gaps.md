# Remaining openpyxl API compatibility gaps (pyo3)

- `iter_rows`/`iter_cols` materialize full lists of PyCell objects instead
  of yielding lazily; `next(ws.iter_rows())` raises TypeError. Dense loops
  over min..max mean one stray cell at XFD1048576 iterates ~17B times.
- `PatternFill(start_color=..., end_color=...)` unsupported (the canonical
  openpyxl invocation); `Font(color=...)`/`Side(color=...)` accept only str,
  not Color objects; `PyColor`/`PyGradientFill` are registered but accepted
  nowhere.
- `cell.number_format = None` on a connected cell doesn't clear the
  workbook-side format.
- Number cells always return float (openpyxl returns int for integral
  values); `wb.active` always returns sheet 0 (no active-index tracking).
- No `.pyi` type stubs shipped — no IDE/type-checker support.
- `ws.append` accepts only sequences, not generators or dicts.
- PyCell holds `Py<PyWorkbook>` + arbitrary PyObject without
  `__traverse__`/`__clear__` — uncollectable reference cycles possible.

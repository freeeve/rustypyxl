# PyO3 worksheet handles silently target the wrong sheet

`PyWorksheet` resolves by raw `index` for half its methods (`append`,
`sheet_dims`, `read_value`, `set_title`, `freeze_panes`, …) and by
`cached_title` for the other half (`__setitem__`, every `PyCell` operation).
`wb.remove()`, `create_sheet(index=...)`, and `move_sheet()` shift the
worksheet Vec without invalidating outstanding handles.

Repro: hold `ws2 = wb["Sheet2"]` (index 1); `wb.remove(wb["Sheet1"])`; now
`ws2.append([...])` writes into Sheet3 (the new index 1) with no error, while
`ws2["A1"] = v` still writes to Sheet2 by name. Renaming a sheet orphans all
existing cell handles (cell setters then silently no-op, getters raise).

Fix direction: stable sheet IDs in rustypyxl-core (e.g. a monotonically
increasing id per worksheet, never reused), with PyWorksheet/PyCell resolving
through the id via a single lookup path. Cell setters must propagate errors
instead of `let _ =` (rustypyxl-pyo3/src/cell.rs).

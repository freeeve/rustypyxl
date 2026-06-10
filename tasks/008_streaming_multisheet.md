# Streaming workbook cannot write more than one sheet

`StreamingWorkbook::create_sheet` errors if a sheet is open, but the only
thing that closes a sheet is the private `finalize_sheet`, called solely
from `close(mut self)` — which consumes the workbook. The multi-sheet loops
in `write_content_types`/`write_workbook_xml`/`write_workbook_rels` are dead
code (streaming.rs). The Python docstring claims creating a new sheet
"finalizes the previous one", which is false.

Fix: make `create_sheet` finalize the currently open sheet (or expose
`finalize_sheet`), and validate row/column limits in `append_row`
(currently rows past 1,048,576 and >16,384 columns are written silently).
Add `__enter__`/`__exit__` to the Python binding so `with` blocks close the
file; dropping without close() currently leaves a truncated file.

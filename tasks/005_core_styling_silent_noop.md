# Core Rust styling APIs are silent no-ops on save

`Worksheet::set_cell_style` / `Workbook::set_cell_font` /
`set_cell_alignment` / `set_cell_number_format` store styles on
`CellData.style` / `.number_format`, but `write_cell_direct` emits the `s=`
attribute solely from `CellData.style_index`, which nothing in core ever
populates (only the pyo3 binding wires styles through
`StyleRegistry::get_or_add_cell_xf`). Every styling call on the core API
produces unstyled output.

Fix: during `write_workbook_contents`, resolve `cell.style` /
`cell.number_format` through the style registry to produce `style_index`.

Related style-registry fidelity bugs:
- Custom numFmt id allocation is `164 + len()`, which collides with loaded
  ids (use max(existing)+1, floor 164) — style.rs:560.
- `get_num_fmt_string` resolves only ids 0-4/9/10/14; common date/time
  builtins (20, 22, ...) round-trip as `number_format: None`.
- Styled-only cells inserted via `set_cell_font`/`set_cell_alignment` bypass
  `update_dimensions`, desyncing `max_row`/`max_column`.

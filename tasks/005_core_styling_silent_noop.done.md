# Core Rust styling APIs are silent no-ops on save — DONE

- Save resolves CellData.style / .number_format into registry xfs (cloned
  StyleRegistry + per-sheet style_index overrides threaded into the cell
  writer), so core-API styling reaches the file. The pyo3 path, which
  resolves style_index eagerly, is unchanged.
- Core setters clear style_index on modification, so restyling a loaded
  cell re-resolves instead of writing the stale xf.
- set_cell_font/set_cell_alignment moved into Worksheet, merge with the
  existing cell style, and update dimensions (styled-only cells previously
  desynced max_row/max_column).
- Custom numFmt id allocation uses max(existing)+1 with a floor of 164
  (was 164 + len(), colliding with loaded non-contiguous ids).
- Built-in number formats unified into StyleRegistry::builtin_num_fmt_code
  (ids 0-49) used by both load passes and get_num_fmt_string; common
  date/time builtins (h:mm, m/d/yy h:mm, ...) now survive round-trips.

Verified by integration test (core-styled file has s= attributes, styles
in styles.xml, and styles/formats/alignment survive reload) and registry
unit tests (id collision, builtin table inverse property).

Not done (low priority, tracked in review notes): max_row never shrinks
when cells are removed; get_cell_style returns alignment ungated by
apply_alignment.

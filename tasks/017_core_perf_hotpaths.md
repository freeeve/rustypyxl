# Core perf: per-cell allocations and quadratic row iteration

Performance issues in rustypyxl-core hot paths, roughly ordered by impact.

## Write path (writer.rs)

- Row-grouping HashMap pre-sized to max_row, not populated-row count
  (writer.rs:1103-1105): `HashMap::with_capacity(worksheet.max_row as usize)`.
  A sparse sheet with one cell at row 1,000,000 reserves ~1M buckets for a
  handful of entries. Bound by `min(max_row, cells.len())`.
- Two heap allocations per cell for the coordinate string
  (writer.rs:1152 parallel path, :1201 sequential):
  `format!("{}{}", column_to_letter(col), row)`--column_to_letter returns an
  owned String and format! allocates a second. ~2M allocs per 1M-cell sheet.
  Write letters+row into a stack buffer or straight into buf in
  write_cell_direct.
- Per-styled-cell String alloc for the ` s="N"` attribute
  (writer.rs:150-157): build it directly into buf with the itoa buffer
  already in scope.

## Load path (workbook.rs, worksheet.rs)

- One String allocation per typed cell for CellData.data_type
  (workbook.rs:2965-2973, duplicated at :2513-2521 for self-closing cells):
  `b's' => Some("s".to_string())` runs for every shared-string cell.
  Make data_type a small enum or &'static str--millions of allocs on large
  files.
- `iter_row` is O(total_cells) per call (worksheet.rs:504-519): it filters
  and sorts a copy of the entire cell map for one row, so row-by-row
  iteration is O(rows * total_cells). Build a row index or sort once and
  slice.
- styles.xml is parsed twice from scratch (workbook.rs:1509, second
  Reader over the same bytes to build cellXfs). Low impact, redundant.
- Untrusted `<dimension>` drives a large upfront reserve: a few-byte
  `<dimension ref="A1:E1000000"/>` yields the 5M-cell cap and
  `cells.reserve(5_000_000)` (workbook.rs:1862-1868, :2361, :2548)--hundreds
  of MB from attacker-controlled input before any cell is read. Lower the
  cap or reserve incrementally.

Benchmark before/after with the existing criterion benches
(cargo bench -p rustypyxl-core) on large-file load/save.

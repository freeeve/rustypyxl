# rustypyxl-core

Fast, dependency-light Rust library for reading and writing Excel (XLSX) files.

This is the pure-Rust core of [rustypyxl](https://github.com/freeeve/rustypyxl).
It has no Python dependency and can be used directly from Rust. The
[`rustypyxl`](https://pypi.org/project/rustypyxl/) Python package is a thin PyO3
binding over this crate, exposing an openpyxl-compatible API.

## Usage

```toml
[dependencies]
rustypyxl-core = "0.5"
```

```rust
use rustypyxl_core::{Workbook, CellValue};

// Write
let mut wb = Workbook::new();
wb.create_sheet(Some("Data".to_string())).unwrap();
wb.set_cell_value_in_sheet("Data", 1, 1, CellValue::from("Hello")).unwrap();
wb.set_cell_value_in_sheet("Data", 1, 2, CellValue::Number(42.0)).unwrap();
wb.save("output.xlsx").unwrap();

// Read
let wb = Workbook::load("output.xlsx").unwrap();
let ws = wb.get_sheet_by_name("Data").unwrap();
assert_eq!(ws.get_cell_value(1, 2), Some(&CellValue::Number(42.0)));
```

## What it does

- Read and write XLSX with full round-trip fidelity for the parts it models
- Cell values: strings, numbers, booleans, dates, formulas (with cached results)
- Styling: fonts, fills, borders, alignment, number formats, theme/indexed/tint
  colors
- Sheet features: merged cells, data validation, conditional formatting, tables,
  autofilters, named ranges, comments, hyperlinks, page setup, protection
- Streaming writes (`StreamingWorkbook`) for constant-memory output of large
  files
- Parallel worksheet parsing and row generation via Rayon

## Feature flags

- `fast-hash` (default): `ahash`/`hashbrown`-backed cell storage
- `parquet`: Parquet import/export via `arrow`/`parquet`
- `s3`: load/save against S3 via the AWS SDK

## Performance

Reading and writing are dominated by pure-Rust work (XML parse/serialize,
deflate), with per-cell allocation kept off the hot paths. For a large sheet a
save is bounded mostly by the compression level; use `CompressionLevel::None`
when speed matters more than file size, or `StreamingWorkbook` when memory does.

## License

MIT. See [LICENSE](LICENSE).

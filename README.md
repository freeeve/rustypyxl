# rustypyxl

[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=freeeve_rustypyxl&metric=alert_status)](https://sonarcloud.io/summary/new_code?id=freeeve_rustypyxl)
[![Maintainability Rating](https://sonarcloud.io/api/project_badges/measure?project=freeeve_rustypyxl&metric=sqale_rating)](https://sonarcloud.io/summary/new_code?id=freeeve_rustypyxl)
[![Security Rating](https://sonarcloud.io/api/project_badges/measure?project=freeeve_rustypyxl&metric=security_rating)](https://sonarcloud.io/summary/new_code?id=freeeve_rustypyxl)
[![Known Vulnerabilities](https://snyk.io/test/github/freeeve/rustypyxl/badge.svg)](https://snyk.io/test/github/freeeve/rustypyxl)

A Rust-powered Excel (XLSX) library for Python with an openpyxl-compatible API.

It is also a standalone Rust crate: the same library is published on crates.io
as [`rustypyxl`](https://crates.io/crates/rustypyxl) for use directly from Rust,
with no Python involved. See [Using from Rust](#using-from-rust).

## Installation

```bash
pip install rustypyxl
```

## Using from Rust

The core is a normal Rust library -- the Python package is a thin binding over it.

```toml
[dependencies]
rustypyxl = "0.5"
```

```rust
use rustypyxl::{Workbook, CellValue};

let mut wb = Workbook::new();
wb.create_sheet(Some("Data".to_string())).unwrap();
wb.set_cell_value_in_sheet("Data", 1, 1, CellValue::from("Hello")).unwrap();
wb.save("output.xlsx").unwrap();
```

## Usage

```python
import rustypyxl

# Create or load a workbook
wb = rustypyxl.Workbook()
ws = wb.create_sheet('Sheet1')          # or: wb = rustypyxl.load_workbook('input.xlsx'); ws = wb.active

# openpyxl-style cell access
ws['A1'] = 'Hello'
ws['A2'] = 42.5
ws['A3'] = '=SUM(A1:A2)'
ws.cell(row=4, column=1).value = 'world'
print(ws['A1'].value)                   # -> "Hello"

# Append rows, merge, freeze, rename
ws.append(['Name', 'Age', 'Score'])
ws.merge_cells('A1:C1')
ws.freeze_panes = 'A2'
ws.title = 'Data'

# Iterate
for row in ws.iter_rows(values_only=True):
    print(row)

wb.save('output.xlsx')
```

### Bulk API (fastest for large grids)

When writing or reading many rows at once, the workbook-level bulk methods avoid
per-cell Python overhead:

```python
wb.write_rows('Data', [
    ['Name', 'Age', 'Score'],
    ['Alice', 30, 95.5],
    ['Bob', 25, 87.3],
])
data = wb.read_rows('Data', min_row=1, max_row=100)
```

## Features

- **openpyxl-compatible API**: Familiar patterns (`ws['A1']`, `ws.cell()`, `ws.append()`, `iter_rows()`) for easy migration
- **Read and write support**: Full round-trip capability
- **Cell values**: Strings, numbers, booleans, dates, formulas
- **Formatting**: Fonts (incl. underline styles), alignment, fills, borders, number formats
- **Workbook features**: Hyperlinks, comments, named ranges, merged cells, freeze panes
- **Sheet protection**: Cell locking and worksheet protection

Not yet supported through the Python API: inserting/deleting rows and columns, charts, and images.
- **Parquet import/export**: Direct Parquet ↔ Excel conversion (bypasses Python FFI)
- **S3 support**: Works with boto3 via bytes I/O
- **Bytes I/O**: Load from bytes or file-like objects, save to bytes
- **Configurable compression**: Trade off speed vs file size

## Parquet Import

Import large Parquet files directly into Excel worksheets. Data flows from Parquet → Rust → Excel without crossing the Python FFI boundary, making it very fast for large datasets.

```python
import rustypyxl

wb = rustypyxl.Workbook()
wb.create_sheet("Data")

# Import parquet file into sheet
result = wb.insert_from_parquet(
    sheet_name="Data",
    path="large_dataset.parquet",
    start_row=1,
    start_col=1,
    include_headers=True,
    column_renames={"old_name": "new_name"},  # optional
    columns=["col1", "col2", "col3"],  # optional: select specific columns
)

print(f"Imported {result['rows_imported']} rows")
print(f"Data range: {result['range']}")

wb.save("output.xlsx")
```

Performance: ~4 seconds for 1M rows × 20 columns on M1 MacBook Pro.

## Parquet Export

Export worksheet data to Parquet format with automatic type inference:

```python
import rustypyxl

wb = rustypyxl.load_workbook("data.xlsx")

# Export entire sheet
result = wb.export_to_parquet(
    sheet_name="Sheet1",
    path="output.parquet",
    has_headers=True,              # first row contains headers
    compression="snappy",          # snappy, zstd, gzip, lz4, none
    column_renames={"old": "new"}, # optional: rename columns
    column_types={"date_col": "datetime"},  # optional: force column types
)

print(f"Exported {result['rows_exported']} rows")
print(f"File size: {result['file_size']} bytes")

# Export specific range
result = wb.export_range_to_parquet(
    sheet_name="Sheet1",
    path="subset.parquet",
    min_row=1, min_col=1,
    max_row=1000, max_col=5,
)
```

Supported column type hints: `string`, `float64`, `int64`, `boolean`, `date`, `datetime`, `auto`.

## Loading from Bytes or File-like Objects

Load workbooks from in-memory bytes or file-like objects:

```python
import rustypyxl
import io

# From bytes
with open("file.xlsx", "rb") as f:
    data = f.read()
wb = rustypyxl.load_workbook(data)

# From file-like object (e.g., BytesIO, HTTP response)
wb = rustypyxl.load_workbook(io.BytesIO(data))

# Save to bytes (for HTTP responses, S3, etc.)
output_bytes = wb.save_to_bytes()
```

## S3 Support

Use `save_to_bytes()` and `load_workbook(bytes)` with boto3 for S3 integration:

```python
import boto3
import rustypyxl

s3 = boto3.client("s3")

# Load from S3
response = s3.get_object(Bucket="my-bucket", Key="path/to/file.xlsx")
wb = rustypyxl.load_workbook(response["Body"].read())

# Save to S3
data = wb.save_to_bytes()
s3.put_object(Bucket="my-bucket", Key="path/to/output.xlsx", Body=data)
```

This works with any S3-compatible service and uses boto3's credential handling (IAM roles, environment variables, etc.).

## Streaming Writes (Low Memory)

For very large files, use `WriteOnlyWorkbook` which streams rows directly to disk:

```python
import rustypyxl

wb = rustypyxl.WriteOnlyWorkbook("large_output.xlsx")
wb.create_sheet("Data")

for i in range(1_000_000):
    wb.append_row([f"Row {i}", i, i * 1.5, i % 2 == 0])

wb.close()  # Must call close() to finalize the file
```

This uses minimal memory regardless of file size, similar to openpyxl's `write_only=True` mode.

## Benchmarks

Apple Silicon, openpyxl 3.1.5. Times are the **minimum** wall-clock over several
runs (the fastest run is the one least disturbed by other processes), measured
on an otherwise-idle machine. Your results will vary with data and hardware.

### Write Performance (1M rows × 20 columns, mixed data)

| Method | Time | vs openpyxl |
|--------|------|-------------|
| rustypyxl `WriteOnlyWorkbook` (streaming) | ~4s | ~27x |
| rustypyxl `write_rows` (build in memory, then save) | ~29s | ~4x |
| openpyxl (`write_only`) | ~112s | — |

The streaming path is the one to use for large writes: it serializes rows
straight to disk and never holds the sheet in memory. `write_rows` is slower at
this scale because all 20M cell values cross the Python/Rust boundary and the
whole workbook is built in memory first; it is convenient for moderate sheets,
not the throughput path. (Loading straight from Parquet with
`insert_from_parquet` avoids the boundary entirely -- see the Parquet section.)

### Read Performance (min wall time)

| Dataset | rustypyxl | calamine | openpyxl |
|---------|-----------|----------|----------|
| 10k × 20 (numeric) | 0.09s | 0.10s | 0.73s |
| 10k × 20 (strings) | 0.11s | 0.11s | 1.55s |
| 100k × 20 (numeric) | 1.02s | 1.01s | 7.16s |
| 100k × 20 (mixed) | 1.25s | 1.16s | 11.9s |

rustypyxl and [calamine](https://github.com/tafia/calamine) (a read-only Rust
reader, via python-calamine) are within noise of each other, both roughly
5-10x faster than openpyxl's read-only mode.

### Memory Usage (Read)

| Dataset | rustypyxl | calamine | openpyxl |
|---------|-----------|----------|----------|
| 10k × 20 | 29 MB | 9 MB | 11 MB |
| 50k × 20 | 62 MB | 48 MB | 53 MB |
| 100k × 20 | 103 MB | 95 MB | 106 MB |

rustypyxl keeps the whole workbook resident (like openpyxl's default mode), so
its read footprint is comparable to openpyxl and above calamine's streaming
reader. For low-memory reads of very large files, the trade-off is CPU vs RAM.

### Memory Usage (Write)

| Dataset | rustypyxl (`write_rows`) | `WriteOnlyWorkbook` | openpyxl (`write_only`) |
|---------|--------------------------|---------------------|-------------------------|
| 10k × 20 | 10 MB | ~0 MB | 0.4 MB |
| 50k × 20 | 50 MB | ~0 MB | 0.4 MB |
| 100k × 20 | 99 MB | ~0 MB | 0.4 MB |

`WriteOnlyWorkbook` streams rows directly to disk, so its memory stays flat
regardless of file size -- the same idea as openpyxl's `write_only` mode.

## Building from Source

```bash
# Install Rust and maturin
curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
pip install maturin

# Build
maturin develop --release
```

## License

MIT

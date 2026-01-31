# rustypyxl

A Rust-powered Excel (XLSX) library for Python with an openpyxl-compatible API.

## Installation

```bash
pip install rustypyxl
```

## Usage

```python
import rustypyxl

# Load a workbook
wb = rustypyxl.load_workbook('input.xlsx')
ws = wb.active

# Read values
value = wb.get_cell_value('Sheet1', 1, 1)

# Write values
wb.set_cell_value('Sheet1', 1, 1, 'Hello')
wb.set_cell_value('Sheet1', 2, 1, 42.5)
wb.set_cell_value('Sheet1', 3, 1, '=SUM(A1:A2)')

# Bulk operations
wb.write_rows('Sheet1', [
    ['Name', 'Age', 'Score'],
    ['Alice', 30, 95.5],
    ['Bob', 25, 87.3],
])

data = wb.read_rows('Sheet1', min_row=1, max_row=100)

# Save
wb.save('output.xlsx')
```

## Features

- **openpyxl-compatible API**: Familiar patterns for easy migration
- **Read and write support**: Full round-trip capability
- **Cell values**: Strings, numbers, booleans, dates, formulas
- **Formatting**: Fonts, alignment, fills, borders, number formats
- **Workbook features**: Comments, hyperlinks, named ranges, merged cells
- **Sheet features**: Protection, data validation, column/row dimensions
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
print(f"Data range: {result['range_with_headers']}")

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

Micro benchmarks on M1 MacBook Pro. Your results may vary depending on data characteristics and hardware.

### Write Performance (1M rows × 20 columns)

| Library | Time |
|---------|------|
| rustypyxl | ~10s |
| openpyxl | ~200s |

### Read Performance

| Dataset | rustypyxl | calamine | openpyxl |
|---------|-----------|----------|----------|
| 10k × 20 (numeric) | 0.13s | 0.21s | 1.16s |
| 10k × 20 (strings) | 0.14s | 0.26s | 2.97s |
| 100k × 20 (numeric) | 0.84s | 1.76s | 11.5s |
| 100k × 20 (mixed) | 1.40s | 2.36s | 32.9s |

[calamine](https://github.com/tafia/calamine) is a Rust Excel reader with Python bindings via python-calamine (read-only).

### Memory Usage (Read)

| Dataset | rustypyxl | calamine | openpyxl |
|---------|-----------|----------|----------|
| 10k × 20 | 29 MB | 9 MB | 11 MB |
| 50k × 20 | 58 MB | 48 MB | 53 MB |
| 100k × 20 | 95 MB | 95 MB | 106 MB |

### Memory Usage (Write)

| Dataset | rustypyxl | WriteOnlyWorkbook | openpyxl (write_only) |
|---------|-----------|-------------------|----------------------|
| 10k × 20 | 10 MB | ~0 MB | 0.4 MB |
| 50k × 20 | 50 MB | ~0 MB | 0.4 MB |
| 100k × 20 | 99 MB | ~0 MB | 0.4 MB |

`WriteOnlyWorkbook` streams rows directly to disk like openpyxl's write_only mode.

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

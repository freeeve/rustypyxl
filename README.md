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
- **Parquet import**: Fast import from Parquet files (bypasses Python FFI)
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

Note: openpyxl's `write_only=True` mode uses minimal memory (~0.4 MB) by streaming rows to disk. rustypyxl currently holds the full workbook in memory.

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

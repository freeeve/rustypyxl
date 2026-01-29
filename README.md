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
- **Configurable compression**: Trade off speed vs file size

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

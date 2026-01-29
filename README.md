# rustypyxl

A high-performance Excel (XLSX) library for Python, written in Rust. Drop-in replacement for openpyxl with 10-20x faster read/write performance.

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

# Bulk operations (fastest)
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

- **Fast**: 10-20x faster than openpyxl for large files
- **Compatible**: Drop-in replacement for common openpyxl patterns
- **Full-featured**:
  - Cell values (string, number, boolean, date, formula)
  - Cell formatting (font, alignment, fill, borders, number formats)
  - Comments and hyperlinks
  - Named ranges
  - Sheet protection
  - Data validation
  - Merged cells
  - Column/row dimensions
  - Configurable compression

## Performance

In micro benchmarks on M1 MacBook (1M rows × 20 columns):

| Operation | rustypyxl | openpyxl | Speedup |
|-----------|-----------|----------|---------|
| Write     | ~10s      | ~200s    | 20x     |
| Read      | ~11s      | ~180s    | 16x     |

### Read Performance vs calamine

rustypyxl is also faster than [calamine](https://github.com/tafia/calamine), a popular Rust Excel reader (with Python bindings via python-calamine):

| Dataset | rustypyxl | calamine | openpyxl | vs calamine | vs openpyxl |
|---------|-----------|----------|----------|-------------|-------------|
| 10k × 20 (numeric) | 0.13s | 0.21s | 1.16s | 1.7x | 9.1x |
| 10k × 20 (strings) | 0.14s | 0.26s | 2.97s | 1.9x | 21x  |
| 100k × 20 (numeric) | 0.84s | 1.76s | 11.5s | 2.1x | 14x  |
| 100k × 20 (mixed) | 1.40s | 2.36s | 32.9s | 1.7x | 24x  |

Unlike calamine (read-only), rustypyxl also supports writing Excel files.

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

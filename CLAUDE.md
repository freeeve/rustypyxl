# rustypyxl

A Rust-powered Excel (XLSX) library for Python with an openpyxl-compatible API.

## Repository Structure

```
rustypyxl/
├── rustypyxl-core/       # Core Rust library (no Python dependencies)
│   ├── src/
│   │   ├── lib.rs        # Crate entry point, re-exports
│   │   ├── workbook.rs   # Workbook struct, load/save, XML parsing
│   │   ├── worksheet.rs  # Worksheet struct, cell storage
│   │   ├── cell.rs       # CellValue enum, InternedString
│   │   ├── writer.rs     # ZIP/XML writing functions
│   │   ├── style.rs      # Font, Fill, Border, Alignment, CellStyle
│   │   ├── utils.rs      # Coordinate parsing, column letters
│   │   ├── error.rs      # Error types
│   │   ├── streaming.rs  # StreamingWorkbook for low-memory writes
│   │   ├── s3.rs         # S3 load/save (optional, behind "s3" feature)
│   │   ├── parquet_import.rs  # Parquet import/export (optional)
│   │   ├── autofilter.rs # AutoFilter support
│   │   ├── conditional.rs # Conditional formatting
│   │   ├── table.rs      # Table/ListObject support
│   │   ├── chart.rs      # Chart support (partial)
│   │   ├── image.rs      # Image embedding (partial)
│   │   └── pagesetup.rs  # Page setup, margins, headers/footers
│   ├── tests/            # Integration tests
│   ├── benches/          # Criterion benchmarks
│   └── fuzz/             # Fuzz testing targets
│
├── rustypyxl-pyo3/       # Python bindings via PyO3
│   └── src/
│       ├── lib.rs        # Python module definition
│       ├── workbook.rs   # PyWorkbook class
│       ├── worksheet.rs  # PyWorksheet class
│       ├── cell.rs       # PyCell class
│       ├── style.rs      # PyFont, PyAlignment, etc.
│       └── streaming.rs  # PyStreamingWorkbook (WriteOnlyWorkbook)
│
├── tests/                # Python pytest tests
├── benchmarks/           # Python benchmark scripts
└── Cargo.toml            # Workspace definition
```

## Architecture

### Two-Crate Design

1. **rustypyxl-core**: Pure Rust library with no Python dependencies. Can be used standalone in Rust projects.
2. **rustypyxl-pyo3**: Thin Python bindings wrapping rustypyxl-core via PyO3.

### Key Data Structures

**Workbook** (`workbook.rs`):
- `worksheets: Vec<Worksheet>` - ordered list of sheets
- `sheet_names: Vec<String>` - parallel to worksheets
- `named_ranges: Vec<NamedRange>` - workbook-level named ranges
- `compression: CompressionLevel` - for save operations

**Worksheet** (`worksheet.rs`):
- `cells: HashMap<u64, CellData>` - sparse cell storage (key = packed row/col)
- Uses `hashbrown::HashMap` with `ahash` for performance (behind `fast-hash` feature)
- Cell key encoding: `(row as u64) << 32 | col as u64`

**CellValue** (`cell.rs`):
```rust
pub enum CellValue {
    Empty,
    String(InternedString),  // Arc<str> for deduplication
    Number(f64),
    Boolean(bool),
    Formula(String),
    Date(String),
}
```

### File I/O

**Loading** (`workbook.rs:parse_workbook`):
1. Open ZIP archive (supports file path or bytes)
2. Parse `xl/workbook.xml` for sheet names and relationships
3. Parse `xl/sharedStrings.xml` into `Vec<InternedString>`
4. Parse `xl/styles.xml` into style lookup table
5. Parse each `xl/worksheets/sheetN.xml` in parallel (Rayon)

**Saving** (`workbook.rs:save_to_writer`, `writer.rs`):
1. Collect shared strings from all worksheets
2. Write ZIP entries: `[Content_Types].xml`, `_rels/.rels`, etc.
3. Write `xl/workbook.xml`, `xl/sharedStrings.xml`, `xl/styles.xml`
4. Write each worksheet XML (parallel row generation for large sheets)

### Feature Flags

**rustypyxl-core**:
- `fast-hash` (default): Use ahash/hashbrown for faster HashMap
- `parquet`: Enable Parquet import/export via arrow/parquet crates
- `s3`: Enable S3 load/save via aws-sdk-s3 (for pure Rust usage)

**rustypyxl-pyo3**:
- `parquet` (default): Enable Parquet methods

Note: For Python S3 support, use `save_to_bytes()`/`load_workbook(bytes)` with boto3 rather than the Rust S3 feature. This avoids extra dependencies and works with boto3's familiar credential handling.

## Development

### Building

```bash
# Rust library only
cargo build -p rustypyxl-core

# Python extension (development)
maturin develop --release

# Python extension (wheel)
maturin build --release
```

### Testing

```bash
# Rust tests
cargo test -p rustypyxl-core

# Rust tests with S3 feature
cargo test -p rustypyxl-core --features s3

# Python tests
pytest tests/

# Fuzz testing
cd rustypyxl-core/fuzz && cargo +nightly fuzz run fuzz_load
```

### Benchmarks

```bash
# Rust benchmarks
cargo bench -p rustypyxl-core

# Python benchmarks
python benchmarks/benchmark_read.py
python benchmarks/benchmark_write.py
```

## Key Implementation Details

### String Interning

Strings use `Arc<str>` (aliased as `InternedString`) to deduplicate repeated values. Shared strings are collected before writing and referenced by index in cell XML.

### Parallel Processing

- **Reading**: Multiple worksheets parsed in parallel via Rayon
- **Writing**: Row XML generation parallelized for sheets >1000 rows (chunked at 5000 rows)

### Coordinate Parsing

`utils.rs` provides optimized byte-level coordinate parsing:
- `parse_coordinate_bytes(&[u8]) -> Option<(u32, u32)>` - zero-allocation parsing
- `parse_u32_bytes(&[u8]) -> Option<u32>` - fast integer parsing
- Overflow protection for row/column limits

### Memory Efficiency

- Sparse cell storage (only non-empty cells stored)
- Streaming writes via `StreamingWorkbook` for constant memory usage
- Pre-allocated buffers based on dimension hints

### Compression Options

```rust
pub enum CompressionLevel {
    None,     // Stored (fastest, largest files)
    Fast,     // Deflate level 1
    Default,  // Deflate level 6
    Best,     // Deflate level 9 (smallest, slowest)
}
```

## Python API Patterns

The Python API mirrors openpyxl where possible:

```python
# Loading
wb = rustypyxl.load_workbook("file.xlsx")  # file path
wb = rustypyxl.Workbook.load(bytes_data)   # bytes
wb = rustypyxl.Workbook.load(file_obj)     # file-like object

# Accessing sheets
ws = wb.active                    # first sheet
ws = wb["SheetName"]              # by name
ws = wb.worksheets[0]             # by index

# Cell access (via worksheet)
cell = ws["A1"]
cell = ws.cell(row=1, column=1)
value = cell.value

# Bulk operations (via workbook, faster)
wb.write_rows("Sheet1", [[1, 2, 3], [4, 5, 6]])
data = wb.read_rows("Sheet1", min_row=1, max_row=100)

# Saving
wb.save("output.xlsx")            # file path
data = wb.save_to_bytes()         # bytes
```

## Error Handling

All Rust errors are `RustypyxlError` enum variants, converted to Python `ValueError` in bindings:

```rust
pub enum RustypyxlError {
    Io(std::io::Error),
    Zip(zip::result::ZipError),
    Xml(quick_xml::Error),
    InvalidCoordinate(String),
    WorksheetNotFound(String),
    WorksheetAlreadyExists(String),
    NoWorksheets,
    InvalidFormat(String),
    ParseError(String),
    S3Error(String),
    Custom(String),
}
```

## XLSX Format Notes

- XLSX is a ZIP archive containing XML files
- Cell values either inline or reference shared strings table
- Styles stored separately and referenced by index (xf)
- Worksheets can use relationship IDs that don't match sheet IDs

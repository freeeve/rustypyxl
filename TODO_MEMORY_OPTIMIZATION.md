# Memory Optimization TODO

## Current State (v0.1.x)

| Dataset | rustypyxl | calamine | openpyxl |
|---------|-----------|----------|----------|
| 10k × 20 | 28.6 MB | 9.4 MB | 10.9 MB |
| 50k × 20 | 58.2 MB | 47.6 MB | 52.9 MB |
| 100k × 20 | 95.2 MB | 95.2 MB | 105.5 MB |

**Problem**: rustypyxl uses ~3x more memory than calamine for small datasets.

---

## Root Causes

### 1. CellData Struct Too Large (~136 bytes per cell)

```rust
pub struct CellData {
    pub value: CellValue,              // 24 bytes
    pub style: Option<Arc<CellStyle>>, // 16 bytes
    pub number_format: Option<String>, // 24 bytes - PROBLEM
    pub data_type: Option<String>,     // 24 bytes - PROBLEM
    pub hyperlink: Option<String>,     // 24 bytes - PROBLEM
    pub comment: Option<String>,       // 24 bytes - PROBLEM
}
```

Most cells have None for hyperlink, comment, number_format, data_type - but we pay 96 bytes of overhead anyway.

### 2. No String Interning for Metadata

- `number_format` repeats same values ("General", "0.00", "mm/dd/yyyy") across thousands of cells
- Each cell allocates a separate String instead of sharing via Arc<str>

### 3. Full CellStyle Clone Per Cell

- Each styled cell carries Arc to full CellStyle
- Font, Fill, Border structs use `Option<String>` instead of `Arc<str>`
- Same font name "Calibri" allocated N times

---

## Optimization Options

### Option A: String Interning (Conservative)
- Wrap metadata strings in `Arc<str>`
- Use string pool for number_format, font names, colors
- **Savings**: ~4 MB (14-21% reduction)
- **Effort**: Low
- **Risk**: Low

### Option B: Compact Cell Enum (Recommended)
- Use `enum Cell { Simple(CellValue), Full(Box<CellData>) }`
- Only use Full variant for cells with metadata
- **Savings**: ~8 MB (25-35% reduction)
- **Effort**: Medium
- **Risk**: Moderate (requires match everywhere)

### Option C: Full Redesign (Aggressive)
- Remove most metadata from CellData
- Store only: value, style_id, hyperlink_id, comment_id
- Use worksheet-level pools for strings
- **Savings**: ~20 MB (65% reduction, calamine parity)
- **Effort**: High
- **Risk**: High (major API changes)

---

## Recommended Implementation Plan

### Phase 1: String Interning
1. Create `StringPool` struct using `Arc<str>`
2. Intern number_format values during parsing
3. Intern font names, colors in style.rs
4. Change `Option<String>` to `Option<Arc<str>>` where appropriate

### Phase 2: Compact Cell Enum
1. Create `SimpleCell` variant with just CellValue
2. Move hyperlink/comment to worksheet-level HashMap
3. Only use full CellData for cells that need it

### Phase 3: Style Index References
1. Store `style_id: Option<u32>` instead of `Arc<CellStyle>`
2. Resolve styles on demand via workbook lookup
3. Share Font, Fill, Border objects across styles

---

## Files to Modify

| File | Changes |
|------|---------|
| `rustypyxl-core/src/worksheet.rs` | CellData struct, cell storage |
| `rustypyxl-core/src/style.rs` | Use Arc<str> for strings |
| `rustypyxl-core/src/workbook.rs` | String pool, style indexing |
| `rustypyxl-core/src/cell.rs` | CellValue::Date could use Arc<str> |

---

## Target Goals

| Dataset | Current | Target | Reduction |
|---------|---------|--------|-----------|
| 10k × 20 | 28.6 MB | ~15 MB | 48% |
| 50k × 20 | 58.2 MB | ~40 MB | 31% |
| 100k × 20 | 95.2 MB | ~80 MB | 16% |

Stretch goal: Match calamine's 9.4 MB for 10k×20 (requires Option C).

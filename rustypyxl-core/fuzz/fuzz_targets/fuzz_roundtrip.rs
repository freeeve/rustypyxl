#![no_main]

use libfuzzer_sys::fuzz_target;
use arbitrary::{Arbitrary, Unstructured};
use rustypyxl_core::{Workbook, CellValue, CellStyle, Font, Fill, Alignment};
use std::sync::Arc;

/// Maximum dimensions to prevent OOM
const MAX_ROWS: u32 = 50;
const MAX_COLS: u32 = 20;
const MAX_SHEETS: usize = 3;
const MAX_CELLS_PER_SHEET: usize = 100;

/// A fuzzable cell value
#[derive(Debug, Clone)]
enum FuzzCellValue {
    Empty,
    String(String),
    Number(f64),
    Boolean(bool),
    Formula(String),
}

impl<'a> Arbitrary<'a> for FuzzCellValue {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        let choice: u8 = u.int_in_range(0..=4)?;
        Ok(match choice {
            0 => FuzzCellValue::Empty,
            1 => {
                let len: usize = u.int_in_range(0..=50)?;
                let s: String = (0..len)
                    .map(|_| {
                        let c: u8 = u.int_in_range(32..=126).unwrap_or(32);
                        c as char
                    })
                    .collect();
                FuzzCellValue::String(s)
            }
            2 => {
                let n: f64 = u.arbitrary()?;
                // Avoid NaN and infinity which don't roundtrip well in XML
                if n.is_nan() || n.is_infinite() {
                    FuzzCellValue::Number(0.0)
                } else {
                    FuzzCellValue::Number(n)
                }
            }
            3 => FuzzCellValue::Boolean(u.arbitrary()?),
            _ => {
                // Simple formulas
                let formulas = ["=A1+B1", "=SUM(A1:A10)", "=1+1", "=IF(A1>0,1,0)", "=NOW()"];
                let idx: usize = u.int_in_range(0..=4)?;
                FuzzCellValue::Formula(formulas[idx].to_string())
            }
        })
    }
}

impl From<FuzzCellValue> for CellValue {
    fn from(v: FuzzCellValue) -> Self {
        match v {
            FuzzCellValue::Empty => CellValue::Empty,
            FuzzCellValue::String(s) => CellValue::String(Arc::from(s.as_str())),
            FuzzCellValue::Number(n) => CellValue::Number(n),
            FuzzCellValue::Boolean(b) => CellValue::Boolean(b),
            FuzzCellValue::Formula(f) => CellValue::Formula(f),
        }
    }
}

/// A cell with position and value
#[derive(Debug, Arbitrary)]
struct FuzzCell {
    row: u8,
    col: u8,
    value: FuzzCellValue,
    has_style: bool,
    bold: bool,
    italic: bool,
    font_size: u8,
    bg_color: Option<[u8; 3]>,
    h_align: u8,
    v_align: u8,
}

/// A sheet with cells
#[derive(Debug)]
struct FuzzSheet {
    name: String,
    cells: Vec<FuzzCell>,
    merged_ranges: Vec<String>,
}

impl<'a> Arbitrary<'a> for FuzzSheet {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        // Generate sheet name (alphanumeric, 1-20 chars)
        let name_len: usize = u.int_in_range(1..=15)?;
        let name: String = (0..name_len)
            .map(|_| {
                let c: u8 = u.int_in_range(0..=35).unwrap_or(0);
                if c < 10 {
                    (b'0' + c) as char
                } else {
                    (b'A' + c - 10) as char
                }
            })
            .collect();

        // Generate cells
        let num_cells: usize = u.int_in_range(0..=MAX_CELLS_PER_SHEET)?;
        let mut cells = Vec::with_capacity(num_cells);
        for _ in 0..num_cells {
            cells.push(u.arbitrary()?);
        }

        // Generate merged ranges (0-3)
        let num_merged: usize = u.int_in_range(0..=3)?;
        let mut merged_ranges = Vec::with_capacity(num_merged);
        for _ in 0..num_merged {
            let r1: u8 = u.int_in_range(1..=5)?;
            let c1: u8 = u.int_in_range(1..=5)?;
            let r2: u8 = u.int_in_range(r1..=r1.saturating_add(2).min(10))?;
            let c2: u8 = u.int_in_range(c1..=c1.saturating_add(2).min(10))?;
            if r1 != r2 || c1 != c2 {
                // Format as Excel range like "A1:B2"
                let col1 = (b'A' + c1 - 1) as char;
                let col2 = (b'A' + c2 - 1) as char;
                merged_ranges.push(format!("{}{}:{}{}", col1, r1, col2, r2));
            }
        }

        Ok(FuzzSheet { name, cells, merged_ranges })
    }
}

/// A complete workbook configuration
#[derive(Debug)]
struct FuzzWorkbook {
    sheets: Vec<FuzzSheet>,
}

impl<'a> Arbitrary<'a> for FuzzWorkbook {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        let num_sheets: usize = u.int_in_range(1..=MAX_SHEETS)?;
        let mut sheets = Vec::with_capacity(num_sheets);
        let mut used_names = std::collections::HashSet::new();

        for i in 0..num_sheets {
            let mut sheet: FuzzSheet = u.arbitrary()?;
            // Ensure unique sheet names
            if used_names.contains(&sheet.name) || sheet.name.is_empty() {
                sheet.name = format!("Sheet{}", i + 1);
            }
            used_names.insert(sheet.name.clone());
            sheets.push(sheet);
        }

        Ok(FuzzWorkbook { sheets })
    }
}

fn build_style(cell: &FuzzCell) -> Option<CellStyle> {
    if !cell.has_style {
        return None;
    }

    let font = Font {
        name: Some("Arial".to_string()),
        size: Some(cell.font_size.max(8).min(72) as f64),
        bold: cell.bold,
        italic: cell.italic,
        ..Default::default()
    };

    let fill = cell.bg_color.map(|rgb| Fill {
        pattern_type: Some("solid".to_string()),
        fg_color: Some(format!("{:02X}{:02X}{:02X}", rgb[0], rgb[1], rgb[2])),
        bg_color: None,
    });

    let h_align = match cell.h_align % 4 {
        0 => None,
        1 => Some("left".to_string()),
        2 => Some("center".to_string()),
        _ => Some("right".to_string()),
    };

    let v_align = match cell.v_align % 4 {
        0 => None,
        1 => Some("top".to_string()),
        2 => Some("center".to_string()),
        _ => Some("bottom".to_string()),
    };

    let alignment = if h_align.is_some() || v_align.is_some() {
        Some(Alignment {
            horizontal: h_align,
            vertical: v_align,
            ..Default::default()
        })
    } else {
        None
    };

    Some(CellStyle {
        font: Some(font),
        fill,
        alignment,
        ..Default::default()
    })
}

fn values_equal(a: &CellValue, b: &CellValue) -> bool {
    match (a, b) {
        (CellValue::Empty, CellValue::Empty) => true,
        (CellValue::String(s1), CellValue::String(s2)) => s1.as_ref() == s2.as_ref(),
        (CellValue::Number(n1), CellValue::Number(n2)) => {
            // Allow small floating point differences
            (n1 - n2).abs() < 1e-10 || (n1.is_nan() && n2.is_nan())
        }
        (CellValue::Boolean(b1), CellValue::Boolean(b2)) => b1 == b2,
        (CellValue::Formula(f1), CellValue::Formula(f2)) => f1 == f2,
        // Empty string and Empty are equivalent
        (CellValue::Empty, CellValue::String(s)) | (CellValue::String(s), CellValue::Empty) => s.is_empty(),
        _ => false,
    }
}

fuzz_target!(|data: &[u8]| {
    let mut u = Unstructured::new(data);
    let Ok(fuzz_wb) = FuzzWorkbook::arbitrary(&mut u) else {
        return;
    };

    // Build the workbook
    let mut wb = Workbook::new();

    // Track what we set for verification
    let mut expected: Vec<(String, Vec<(u32, u32, CellValue)>)> = Vec::new();

    for fuzz_sheet in &fuzz_wb.sheets {
        // Create sheet
        let sheet = match wb.create_sheet(Some(fuzz_sheet.name.clone())) {
            Ok(s) => s,
            Err(_) => continue,
        };

        // Use HashMap to track the LAST value set for each cell (handles overwrites)
        let mut sheet_cells_map: std::collections::HashMap<(u32, u32), CellValue> = std::collections::HashMap::new();

        // Add cells
        for cell in &fuzz_sheet.cells {
            let row = (cell.row as u32 % MAX_ROWS).max(1);
            let col = (cell.col as u32 % MAX_COLS).max(1);
            let value: CellValue = cell.value.clone().into();

            sheet.set_cell_value(row, col, value.clone());
            // Overwrite in map - this is the value that will actually be in the sheet
            sheet_cells_map.insert((row, col), value);

            // Apply style if present
            if let Some(style) = build_style(cell) {
                sheet.set_cell_style(row, col, style);
            }
        }

        // Convert map to vec for verification
        let sheet_cells: Vec<_> = sheet_cells_map.into_iter()
            .map(|((row, col), value)| (row, col, value))
            .collect();

        // Add merged cells
        for range in &fuzz_sheet.merged_ranges {
            sheet.merge_cells(range);
        }

        expected.push((fuzz_sheet.name.clone(), sheet_cells));
    }

    // Save to temp file
    let temp_file = match tempfile::NamedTempFile::new() {
        Ok(f) => f,
        Err(_) => return,
    };
    let temp_path = temp_file.path().to_str().unwrap();

    if wb.save(temp_path).is_err() {
        return;
    }

    // Load back
    let loaded_wb = match Workbook::load(temp_path) {
        Ok(wb) => wb,
        Err(e) => {
            panic!("Failed to load workbook that we just saved: {:?}", e);
        }
    };

    // Verify data
    for (sheet_name, cells) in &expected {
        let sheet = match loaded_wb.get_sheet_by_name(sheet_name) {
            Ok(s) => s,
            Err(_) => {
                panic!("Sheet '{}' missing after roundtrip!", sheet_name);
            }
        };

        for (row, col, expected_value) in cells {
            let actual_value = sheet.get_cell_value(*row, *col).unwrap_or(&CellValue::Empty);

            if !values_equal(expected_value, actual_value) {
                panic!(
                    "Cell {}:{},{} mismatch!\nExpected: {:?}\nActual: {:?}",
                    sheet_name, row, col, expected_value, actual_value
                );
            }
        }
    }
});

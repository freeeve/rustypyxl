use rustypyxl_core::{Workbook, CellValue};
use tempfile::NamedTempFile;

#[test]
fn test_boolean_roundtrip() {
    // Create workbook with Boolean values
    let mut wb = Workbook::new();
    let ws = wb.create_sheet(Some("Test".to_string())).unwrap();
    ws.set_cell_value(1, 1, CellValue::Boolean(false));
    ws.set_cell_value(1, 2, CellValue::Boolean(true));

    // Save to temp file
    let temp_file = NamedTempFile::new().unwrap();
    let temp_path = temp_file.path().to_str().unwrap();
    wb.save(temp_path).unwrap();

    // Load it back
    let loaded_wb = Workbook::load(temp_path).unwrap();
    let loaded_ws = loaded_wb.get_sheet_by_name("Test").unwrap();

    let val1 = loaded_ws.get_cell_value(1, 1);
    let val2 = loaded_ws.get_cell_value(1, 2);

    // Verify
    assert_eq!(val1, Some(&CellValue::Boolean(false)), "Boolean false should roundtrip");
    assert_eq!(val2, Some(&CellValue::Boolean(true)), "Boolean true should roundtrip");
}

#[test]
fn test_boolean_roundtrip_multiple_sheets() {
    // Reproduce the fuzz crash case: sheet "C" with Boolean(false) at (1,1)
    let mut wb = Workbook::new();

    // Create multiple sheets like the fuzzer does
    let sheet_c = wb.create_sheet(Some("C".to_string())).unwrap();
    sheet_c.set_cell_value(1, 1, CellValue::Boolean(false));

    // Save to temp file
    let temp_file = NamedTempFile::new().unwrap();
    let temp_path = temp_file.path().to_str().unwrap();
    wb.save(temp_path).unwrap();

    // Load it back
    let loaded_wb = Workbook::load(temp_path).unwrap();

    // Check sheet names
    println!("Sheet names: {:?}", loaded_wb.sheet_names());

    let loaded_ws = loaded_wb.get_sheet_by_name("C").unwrap();
    let val = loaded_ws.get_cell_value(1, 1);

    println!("Cell C:1,1 value: {:?}", val);

    // Verify
    assert_eq!(val, Some(&CellValue::Boolean(false)), "Boolean false in sheet C should roundtrip");
}

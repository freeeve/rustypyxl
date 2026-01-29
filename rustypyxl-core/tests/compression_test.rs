use std::fs::File;
use std::io::Read;
use zip::ZipArchive;

#[test]
fn test_no_compression() {
    use rustypyxl_core::{Workbook, CellValue};
    
    let mut wb = Workbook::new();
    let _ws = wb.create_sheet(Some("Test".to_string())).unwrap();
    wb.set_cell_value_in_sheet("Test", 1, 1, CellValue::String(std::sync::Arc::from("Hello"))).unwrap();
    
    wb.save("/tmp/test_compression.xlsx").unwrap();
    
    // Check the compression method
    let file = File::open("/tmp/test_compression.xlsx").unwrap();
    let mut archive = ZipArchive::new(file).unwrap();
    
    for i in 0..archive.len() {
        let file = archive.by_index(i).unwrap();
        println!("File: {}, compression: {:?}", file.name(), file.compression());
    }
    
    // Check specifically for the sheet
    let sheet = archive.by_name("xl/worksheets/sheet1.xml").unwrap();
    let compression = sheet.compression();
    println!("Sheet compression: {:?}", compression);
    
    assert_eq!(compression, zip::CompressionMethod::Stored, "Expected Stored compression, got {:?}", compression);
}

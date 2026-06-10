use std::fs::File;
use zip::ZipArchive;

/// The default compression must be Deflate (level 6), as documented; the old
/// Stored default ("for benchmarking") shipped needlessly large files.
#[test]
fn test_default_compression_is_deflate() {
    use rustypyxl_core::{CellValue, Workbook};

    let dir = std::env::temp_dir().join("rustypyxl_tests");
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join("test_compression.xlsx");
    let path_str = path.to_str().unwrap();

    let mut wb = Workbook::new();
    let _ws = wb.create_sheet(Some("Test".to_string())).unwrap();
    wb.set_cell_value_in_sheet(
        "Test",
        1,
        1,
        CellValue::String(std::sync::Arc::from("Hello")),
    )
    .unwrap();

    wb.save(path_str).unwrap();

    let file = File::open(&path).unwrap();
    let mut archive = ZipArchive::new(file).unwrap();
    let sheet = archive.by_name("xl/worksheets/sheet1.xml").unwrap();
    assert_eq!(
        sheet.compression(),
        zip::CompressionMethod::Deflated,
        "default save should be compressed"
    );
    drop(sheet);

    std::fs::remove_file(&path).ok();
}

/// CompressionLevel::None still produces Stored entries when asked for.
#[test]
fn test_no_compression_opt_in() {
    use rustypyxl_core::{CellValue, CompressionLevel, Workbook};

    let dir = std::env::temp_dir().join("rustypyxl_tests");
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join("test_compression_none.xlsx");
    let path_str = path.to_str().unwrap();

    let mut wb = Workbook::new();
    let _ws = wb.create_sheet(Some("Test".to_string())).unwrap();
    wb.set_cell_value_in_sheet(
        "Test",
        1,
        1,
        CellValue::String(std::sync::Arc::from("Hello")),
    )
    .unwrap();
    wb.set_compression(CompressionLevel::None);

    wb.save(path_str).unwrap();

    let file = File::open(&path).unwrap();
    let mut archive = ZipArchive::new(file).unwrap();
    let sheet = archive.by_name("xl/worksheets/sheet1.xml").unwrap();
    assert_eq!(sheet.compression(), zip::CompressionMethod::Stored);
    drop(sheet);

    std::fs::remove_file(&path).ok();
}

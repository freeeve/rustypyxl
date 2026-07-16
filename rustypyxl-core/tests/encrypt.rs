//! Writing password-protected (agile-encrypted) workbooks. The primary check is
//! an internal round-trip: encrypt with rustypyxl, then decrypt with rustypyxl.
//! Cross-validation against msoffcrypto-tool lives in the Python tests.

#![cfg(feature = "encrypt")]

use rustypyxl::{CellValue, Workbook};

fn sample() -> Workbook {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("Secret".to_string())).unwrap();
    wb.set_cell_value_in_sheet("Secret", 1, 1, CellValue::from("hello"))
        .unwrap();
    wb.set_cell_value_in_sheet("Secret", 1, 2, CellValue::Number(42.0))
        .unwrap();
    wb.set_cell_value_in_sheet("Secret", 2, 1, CellValue::from("world"))
        .unwrap();
    wb
}

#[test]
fn encrypt_then_decrypt_round_trip() {
    let enc = sample().save_to_bytes_with_password("s3cret").unwrap();

    // The output is a CFB container, not a ZIP.
    assert_eq!(&enc[..8], &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);

    let wb = Workbook::load_from_bytes_with_password(&enc, "s3cret").unwrap();
    let ws = wb.get_sheet_by_name("Secret").unwrap();
    assert_eq!(ws.get_cell_value(1, 1), Some(&CellValue::from("hello")));
    assert_eq!(ws.get_cell_value(1, 2), Some(&CellValue::Number(42.0)));
    assert_eq!(ws.get_cell_value(2, 1), Some(&CellValue::from("world")));
}

#[test]
fn wrong_password_fails_to_open_our_own_output() {
    let enc = sample().save_to_bytes_with_password("right").unwrap();
    assert!(Workbook::load_from_bytes_with_password(&enc, "wrong").is_err());
}

#[test]
fn larger_workbook_round_trips() {
    // Exercise a package that spans multiple 4096-byte segments and the regular
    // FAT (a big EncryptedPackage stream).
    let mut wb = Workbook::new();
    wb.create_sheet(Some("Data".to_string())).unwrap();
    for row in 1..=500u32 {
        wb.set_cell_value_in_sheet("Data", row, 1, CellValue::from(format!("row-{row}")))
            .unwrap();
        wb.set_cell_value_in_sheet("Data", row, 2, CellValue::Number(row as f64))
            .unwrap();
    }
    let enc = wb.save_to_bytes_with_password("pw").unwrap();
    let back = Workbook::load_from_bytes_with_password(&enc, "pw").unwrap();
    let ws = back.get_sheet_by_name("Data").unwrap();
    assert_eq!(ws.get_cell_value(500, 1), Some(&CellValue::from("row-500")));
    assert_eq!(ws.get_cell_value(250, 2), Some(&CellValue::Number(250.0)));
}

//! Reading a password-protected (agile-encrypted) workbook. The fixture was
//! encrypted by msoffcrypto-tool (an independent implementation), so a
//! successful decrypt cross-validates the crypto against a different codebase.
//!
//! These tests only run with the `decrypt` feature; without it the module is
//! empty.

#![cfg(feature = "decrypt")]

use rustypyxl::{CellValue, Workbook};

/// A workbook whose "Secret" sheet holds [["hello", 42], ["world", 7]],
/// agile-encrypted with the password "s3cret".
const ENCRYPTED: &[u8] = include_bytes!("fixtures/encrypted.xlsx");

#[test]
fn decrypts_agile_encrypted_workbook() {
    let wb = Workbook::load_from_bytes_with_password(ENCRYPTED, "s3cret").unwrap();
    let ws = wb.get_sheet_by_name("Secret").unwrap();
    assert_eq!(ws.get_cell_value(1, 1), Some(&CellValue::from("hello")));
    assert_eq!(ws.get_cell_value(1, 2), Some(&CellValue::Number(42.0)));
    assert_eq!(ws.get_cell_value(2, 1), Some(&CellValue::from("world")));
    assert_eq!(ws.get_cell_value(2, 2), Some(&CellValue::Number(7.0)));
}

#[test]
fn wrong_password_is_an_error() {
    match Workbook::load_from_bytes_with_password(ENCRYPTED, "wrong") {
        Err(e) => assert!(
            format!("{e}").contains("password"),
            "error should mention the password: {e}"
        ),
        Ok(_) => panic!("a wrong password must not decrypt"),
    }
}

#[test]
fn a_non_encrypted_file_loads_normally_through_the_password_loader() {
    // Build a plain workbook and confirm the password loader passes it through.
    let mut wb = Workbook::new();
    wb.create_sheet(Some("S".to_string())).unwrap();
    wb.set_cell_value_in_sheet("S", 1, 1, CellValue::from("plain"))
        .unwrap();
    let bytes = wb.save_to_bytes().unwrap();

    let reloaded = Workbook::load_from_bytes_with_password(&bytes, "ignored").unwrap();
    assert_eq!(
        reloaded
            .get_sheet_by_name("S")
            .unwrap()
            .get_cell_value(1, 1),
        Some(&CellValue::from("plain"))
    );
}

#[test]
fn plain_loader_gives_a_helpful_error_on_an_encrypted_file() {
    match Workbook::load_from_bytes(ENCRYPTED) {
        Err(e) => assert!(
            format!("{e}").contains("encrypted"),
            "error should say the file is encrypted: {e}"
        ),
        Ok(_) => panic!("the plain loader must not open an encrypted file"),
    }
}

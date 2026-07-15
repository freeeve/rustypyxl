//! Auto-fit sizes a column to its content. It is an approximation (character
//! count of the displayed string x padding), so tests assert the width lands in
//! a sensible band and tracks relative content length, not exact pixels.

use rustypyxl::{CellValue, Workbook};

fn sheet_with(cells: &[(u32, u32, CellValue)]) -> Workbook {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("S".to_string())).unwrap();
    let ws = wb.get_sheet_by_name_mut("S").unwrap();
    for (row, col, val) in cells {
        ws.set_cell_value(*row, *col, val.clone());
    }
    wb
}

#[test]
fn width_tracks_the_longest_value() {
    let mut wb = sheet_with(&[
        (1, 1, CellValue::from("hi")),
        (2, 1, CellValue::from("a much longer piece of text")),
        (3, 1, CellValue::from("mid")),
    ]);
    let ws = wb.get_sheet_by_name_mut("S").unwrap();
    let width = ws.auto_fit_column(1).unwrap();

    // "a much longer piece of text" is 27 chars + 2 padding.
    assert!((width - 29.0).abs() < 0.001, "got {width}");
    assert_eq!(ws.get_column_width(1), Some(width));
}

#[test]
fn empty_column_is_left_untouched() {
    let mut wb = sheet_with(&[(1, 1, CellValue::from("x"))]);
    let ws = wb.get_sheet_by_name_mut("S").unwrap();
    // Column 5 has no cells.
    assert_eq!(ws.auto_fit_column(5), None);
    assert_eq!(ws.get_column_width(5), None);
}

#[test]
fn measures_the_formatted_string_not_the_raw_value() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("S".to_string())).unwrap();
    let ws = wb.get_sheet_by_name_mut("S").unwrap();
    // Raw value is short ("0.1"), but under a percent format it displays as
    // "10.00%" (6 chars), which is what should be measured.
    ws.set_cell_value(1, 1, CellValue::Number(0.1));
    ws.set_cell_number_format(1, 1, "0.00%");
    let width = ws.auto_fit_column(1).unwrap();
    assert!((width - 8.0).abs() < 0.001, "got {width}"); // 6 + 2 padding
}

#[test]
fn auto_fit_all_sizes_every_populated_column() {
    let mut wb = sheet_with(&[
        (1, 1, CellValue::from("short")),
        (1, 3, CellValue::from("a longer heading here")),
    ]);
    let ws = wb.get_sheet_by_name_mut("S").unwrap();
    ws.auto_fit_all();

    assert_eq!(ws.get_column_width(1), Some(7.0)); // "short" 5 + 2
    assert_eq!(ws.get_column_width(3), Some(23.0)); // 21 + 2
                                                    // A column with no content is never assigned a width.
    assert_eq!(ws.get_column_width(2), None);
}

#[test]
fn width_never_exceeds_excel_maximum() {
    let long = "x".repeat(400);
    let mut wb = sheet_with(&[(1, 1, CellValue::from(long))]);
    let ws = wb.get_sheet_by_name_mut("S").unwrap();
    assert_eq!(ws.auto_fit_column(1), Some(255.0));
}

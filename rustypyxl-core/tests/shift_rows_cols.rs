//! insert/delete rows and columns: cell shifting plus adjustment of merged
//! ranges, data validations, conditional formatting, and table ranges.

use rustypyxl::conditional::{ConditionalFormatting, ConditionalOperator, ConditionalRule};
use rustypyxl::worksheet::DataValidation;
use rustypyxl::{CellValue, Workbook, Worksheet};

fn num(ws: &Worksheet, row: u32, col: u32) -> Option<f64> {
    match ws.get_cell_value(row, col) {
        Some(CellValue::Number(n)) => Some(*n),
        _ => None,
    }
}

#[test]
fn insert_rows_shifts_cells_down() {
    let mut ws = Worksheet::new("S");
    ws.set_cell_value(1, 1, CellValue::Number(1.0));
    ws.set_cell_value(2, 1, CellValue::Number(2.0));
    ws.set_cell_value(3, 1, CellValue::Number(3.0));

    ws.insert_rows(2, 1); // insert one blank row before row 2

    assert_eq!(num(&ws, 1, 1), Some(1.0), "row above stays");
    assert_eq!(num(&ws, 2, 1), None, "row 2 is now blank");
    assert_eq!(num(&ws, 3, 1), Some(2.0), "old row 2 -> 3");
    assert_eq!(num(&ws, 4, 1), Some(3.0), "old row 3 -> 4");
    assert_eq!(ws.max_row, 4);
}

#[test]
fn delete_rows_removes_and_shifts_up() {
    let mut ws = Worksheet::new("S");
    for r in 1..=4 {
        ws.set_cell_value(r, 1, CellValue::Number(r as f64));
    }

    ws.delete_rows(2, 2); // delete rows 2 and 3

    assert_eq!(num(&ws, 1, 1), Some(1.0));
    assert_eq!(num(&ws, 2, 1), Some(4.0), "old row 4 -> 2");
    assert_eq!(num(&ws, 3, 1), None);
    assert_eq!(ws.max_row, 2);
}

#[test]
fn insert_and_delete_columns() {
    let mut ws = Worksheet::new("S");
    ws.set_cell_value(1, 1, CellValue::Number(1.0));
    ws.set_cell_value(1, 2, CellValue::Number(2.0));
    ws.set_cell_value(1, 3, CellValue::Number(3.0));

    ws.insert_columns(2, 2);
    assert_eq!(num(&ws, 1, 1), Some(1.0));
    assert_eq!(num(&ws, 1, 2), None);
    assert_eq!(num(&ws, 1, 4), Some(2.0), "old col 2 -> 4");
    assert_eq!(num(&ws, 1, 5), Some(3.0));

    ws.delete_columns(2, 2); // undo
    assert_eq!(num(&ws, 1, 2), Some(2.0));
    assert_eq!(num(&ws, 1, 3), Some(3.0));
    assert_eq!(ws.max_column, 3);
}

#[test]
fn merged_range_shifts_grows_and_collapses() {
    // shift entirely below the insert
    let mut ws = Worksheet::new("S");
    ws.merged_cells.push(("A1".into(), "B2".into()));
    ws.insert_rows(1, 1);
    assert_eq!(ws.merged_cells, vec![("A2".to_string(), "B3".to_string())]);

    // insert INSIDE a merge grows it
    let mut ws = Worksheet::new("S");
    ws.merged_cells.push(("A2".into(), "C4".into()));
    ws.insert_rows(3, 1); // row 3 is inside rows 2..=4
    assert_eq!(ws.merged_cells, vec![("A2".to_string(), "C5".to_string())]);

    // deleting the whole merge drops it
    let mut ws = Worksheet::new("S");
    ws.merged_cells.push(("A2".into(), "B3".into()));
    ws.delete_rows(2, 2);
    assert!(ws.merged_cells.is_empty(), "fully-deleted merge is removed");
}

#[test]
fn data_validation_key_and_sqref_shift() {
    let mut ws = Worksheet::new("S");
    let dv = DataValidation {
        validation_type: "list".to_string(),
        formula1: Some("\"a,b\"".to_string()),
        sqref: Some("A5:A10".to_string()),
        ..Default::default()
    };
    ws.add_data_validation(5, 1, dv);

    ws.insert_rows(1, 2); // everything drops by 2

    // keyed cell moved 5 -> 7, and the sqref moved with it
    let (pos, dv) = ws.data_validations.iter().next().expect("validation kept");
    assert_eq!(*pos, (7, 1));
    assert_eq!(dv.sqref.as_deref(), Some("A7:A12"));
}

#[test]
fn conditional_formatting_range_shifts() {
    let mut ws = Worksheet::new("S");
    let mut cf = ConditionalFormatting::new("B2:B10");
    cf.add_rule(ConditionalRule::cell_is(
        ConditionalOperator::GreaterThan,
        "5",
    ));
    ws.add_conditional_formatting(cf);

    ws.insert_columns(1, 1); // push columns right by one

    assert_eq!(ws.conditional_formatting[0].range, "C2:C10");
}

#[test]
fn full_round_trip_after_insert() {
    let mut wb = Workbook::new();
    let ws = wb.create_sheet(Some("Data".to_string())).unwrap();
    ws.set_cell_value(1, 1, CellValue::String("header".into()));
    ws.set_cell_value(2, 1, CellValue::Number(10.0));
    ws.set_cell_value(3, 1, CellValue::Number(20.0));
    ws.merged_cells.push(("A2".into(), "A3".into()));

    ws.insert_rows(2, 1); // blank row under the header

    let reloaded = Workbook::load_from_bytes(&wb.save_to_bytes().unwrap()).unwrap();
    let ws = reloaded.get_sheet_by_name("Data").unwrap();
    assert_eq!(
        ws.get_cell_value(1, 1),
        Some(&CellValue::String("header".into()))
    );
    assert_eq!(ws.get_cell_value(2, 1), None, "inserted blank row");
    assert_eq!(num(ws, 3, 1), Some(10.0));
    assert_eq!(num(ws, 4, 1), Some(20.0));
    assert_eq!(ws.merged_cells, vec![("A3".to_string(), "A4".to_string())]);
}

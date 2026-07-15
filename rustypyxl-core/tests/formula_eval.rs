//! The formula engine evaluates against a workbook's real cell values, including
//! same-sheet and cross-sheet references, chained formula cells (evaluated
//! recursively), and circular references (which yield #REF! rather than looping).

use rustypyxl::{CellValue, FormulaValue, Workbook};

fn wb_with_data() -> Workbook {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("S".to_string())).unwrap();
    wb.set_cell_value_in_sheet("S", 1, 1, CellValue::Number(10.0))
        .unwrap(); // A1
    wb.set_cell_value_in_sheet("S", 2, 1, CellValue::Number(20.0))
        .unwrap(); // A2
    wb.set_cell_value_in_sheet("S", 3, 1, CellValue::Number(30.0))
        .unwrap(); // A3
    wb
}

#[test]
fn evaluates_against_real_cells() {
    let wb = wb_with_data();
    assert_eq!(
        wb.evaluate_formula("S", "=SUM(A1:A3)").unwrap(),
        FormulaValue::Number(60.0)
    );
    assert_eq!(
        wb.evaluate_formula("S", "=A1*A2+A3").unwrap(),
        FormulaValue::Number(230.0)
    );
    assert_eq!(
        wb.evaluate_formula("S", "=IF(A1>5,\"yes\",\"no\")")
            .unwrap(),
        FormulaValue::Text("yes".to_string())
    );
}

#[test]
fn evaluate_cell_computes_formula_cells() {
    let mut wb = wb_with_data();
    // B1 = SUM(A1:A3); B2 = B1 * 2  -> chained formula cells.
    wb.set_cell_value_in_sheet("S", 1, 2, CellValue::Formula("SUM(A1:A3)".to_string()))
        .unwrap();
    wb.set_cell_value_in_sheet("S", 2, 2, CellValue::Formula("B1*2".to_string()))
        .unwrap();

    assert_eq!(
        wb.evaluate_cell("S", 1, 2).unwrap(),
        FormulaValue::Number(60.0)
    );
    // B2 pulls B1's computed value recursively.
    assert_eq!(
        wb.evaluate_cell("S", 2, 2).unwrap(),
        FormulaValue::Number(120.0)
    );
    // A plain value cell just returns its value.
    assert_eq!(
        wb.evaluate_cell("S", 1, 1).unwrap(),
        FormulaValue::Number(10.0)
    );
}

#[test]
fn cross_sheet_reference() {
    let mut wb = wb_with_data();
    wb.create_sheet(Some("Other".to_string())).unwrap();
    wb.set_cell_value_in_sheet("Other", 1, 1, CellValue::Number(7.0))
        .unwrap();
    assert_eq!(
        wb.evaluate_formula("S", "=Other!A1*A1").unwrap(),
        FormulaValue::Number(70.0)
    );
}

#[test]
fn circular_reference_is_ref_error_not_a_hang() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("S".to_string())).unwrap();
    // C1 = C2, C2 = C1 -> a cycle.
    wb.set_cell_value_in_sheet("S", 1, 3, CellValue::Formula("C2".to_string()))
        .unwrap();
    wb.set_cell_value_in_sheet("S", 2, 3, CellValue::Formula("C1".to_string()))
        .unwrap();
    assert_eq!(
        wb.evaluate_cell("S", 1, 3).unwrap(),
        FormulaValue::Error("#REF!".to_string())
    );
}

#[test]
fn calculate_all_writes_cached_values() {
    let mut wb = wb_with_data(); // A1=10, A2=20, A3=30
    wb.set_cell_value_in_sheet("S", 1, 2, CellValue::Formula("SUM(A1:A3)".to_string()))
        .unwrap(); // B1
    wb.set_cell_value_in_sheet("S", 2, 2, CellValue::Formula("B1*2".to_string()))
        .unwrap(); // B2 depends on B1
    wb.set_cell_value_in_sheet("S", 3, 2, CellValue::Formula("1/0".to_string()))
        .unwrap(); // B3 error

    let count = wb.calculate_all();
    assert_eq!(count, 3);

    let ws = wb.get_sheet_by_name("S").unwrap();
    // The cached result is stored on each formula cell.
    assert_eq!(
        ws.get_cell(1, 2).unwrap().cached_formula_value.as_deref(),
        Some("60")
    );
    assert_eq!(
        ws.get_cell(2, 2).unwrap().cached_formula_value.as_deref(),
        Some("120")
    );
    let b3 = ws.get_cell(3, 2).unwrap();
    assert_eq!(b3.cached_formula_value.as_deref(), Some("#DIV/0!"));
    assert_eq!(b3.data_type, Some("e"));

    // The cached values survive a save/load round-trip.
    let out = wb.save_to_bytes().unwrap();
    let reloaded = Workbook::load_from_bytes(&out).unwrap();
    assert_eq!(
        reloaded
            .get_sheet_by_name("S")
            .unwrap()
            .get_cell(1, 2)
            .unwrap()
            .cached_formula_value
            .as_deref(),
        Some("60")
    );
}

#[test]
fn unknown_sheet_and_function_are_errors() {
    let wb = wb_with_data();
    assert!(wb.evaluate_formula("S", "=Missing!A1").unwrap().is_error());
    assert!(wb.evaluate_formula("S", "=NOSUCHFN(1)").unwrap().is_error());
    // An unknown sheet name to the API itself is a distinct error.
    assert!(wb.evaluate_formula("Nope", "=1").is_err());
}

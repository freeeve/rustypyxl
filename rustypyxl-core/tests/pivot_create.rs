//! A pivot table created from a source range is written on save, wired to a
//! generated cache, and reads back through the phase-2 model -- proving the
//! generated parts are structurally consistent.

use rustypyxl::{CellValue, Workbook};
use std::io::{Cursor, Read};
use zip::ZipArchive;

fn part_exists(bytes: &[u8], name: &str) -> bool {
    ZipArchive::new(Cursor::new(bytes.to_vec()))
        .unwrap()
        .by_name(name)
        .is_ok()
}

fn read_part(bytes: &[u8], name: &str) -> String {
    let mut zip = ZipArchive::new(Cursor::new(bytes.to_vec())).unwrap();
    let mut f = zip.by_name(name).unwrap();
    let mut s = String::new();
    f.read_to_string(&mut s).unwrap();
    s
}

fn workbook_with_source() -> Workbook {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("Sales".to_string())).unwrap();
    // Header row.
    for (col, name) in ["Region", "Product", "Amount"].iter().enumerate() {
        wb.set_cell_value_in_sheet("Sales", 1, col as u32 + 1, CellValue::from(*name))
            .unwrap();
    }
    // A few data rows.
    let data = [
        ("East", "Widget", 100.0),
        ("West", "Widget", 150.0),
        ("East", "Gadget", 200.0),
    ];
    for (i, (region, product, amount)) in data.iter().enumerate() {
        let row = i as u32 + 2;
        wb.set_cell_value_in_sheet("Sales", row, 1, CellValue::from(*region))
            .unwrap();
        wb.set_cell_value_in_sheet("Sales", row, 2, CellValue::from(*product))
            .unwrap();
        wb.set_cell_value_in_sheet("Sales", row, 3, CellValue::Number(*amount))
            .unwrap();
    }
    wb
}

#[test]
fn created_pivot_emits_parts_and_wiring() {
    let mut wb = workbook_with_source();
    wb.add_pivot_table(
        "Sales",
        "A1:C4",
        "Sales",
        "F1",
        &["Region".to_string()],
        &["Product".to_string()],
        &[("Amount".to_string(), "sum".to_string())],
        Some("SalesByRegion"),
    )
    .unwrap();

    let out = wb.save_to_bytes().unwrap();

    assert!(part_exists(&out, "xl/pivotCache/pivotCacheDefinition1.xml"));
    assert!(part_exists(&out, "xl/pivotCache/pivotCacheRecords1.xml"));
    assert!(part_exists(&out, "xl/pivotTables/pivotTable1.xml"));
    assert!(part_exists(
        &out,
        "xl/pivotTables/_rels/pivotTable1.xml.rels"
    ));

    // The workbook declares the cache and the sheet links the pivot.
    let wbxml = read_part(&out, "xl/workbook.xml");
    assert!(wbxml.contains("<pivotCaches>"));
    let sheetrels = read_part(&out, "xl/worksheets/_rels/sheet1.xml.rels");
    assert!(sheetrels.contains("../pivotTables/pivotTable1.xml"));

    // Content types declare the new parts.
    let ct = read_part(&out, "[Content_Types].xml");
    assert!(ct.contains("/xl/pivotTables/pivotTable1.xml"));
}

#[test]
fn created_pivot_reads_back_through_the_model() {
    let mut wb = workbook_with_source();
    wb.add_pivot_table(
        "Sales",
        "A1:C4",
        "Sales",
        "F1",
        &["Region".to_string()],
        &[],
        &[("Amount".to_string(), "sum".to_string())],
        Some("SalesByRegion"),
    )
    .unwrap();

    // Round-trip through save + load, then read with the phase-2 model.
    let out = wb.save_to_bytes().unwrap();
    let reloaded = Workbook::load_from_bytes(&out).unwrap();
    let pivots = reloaded.pivot_tables();
    assert_eq!(pivots.len(), 1);
    let p = &pivots[0];
    assert_eq!(p.name, "SalesByRegion");
    assert_eq!(p.source_sheet.as_deref(), Some("Sales"));
    assert_eq!(p.source_ref.as_deref(), Some("A1:C4"));
    assert_eq!(p.cache_fields, vec!["Region", "Product", "Amount"]);
    assert_eq!(p.row_fields, vec!["Region"]);
    assert_eq!(p.data_fields.len(), 1);
    assert_eq!(p.data_fields[0].source_field, "Amount");
    assert_eq!(p.data_fields[0].subtotal, "sum");
}

#[test]
fn unknown_field_is_an_error() {
    let mut wb = workbook_with_source();
    let err = wb.add_pivot_table(
        "Sales",
        "A1:C4",
        "Sales",
        "F1",
        &["NoSuchField".to_string()],
        &[],
        &[("Amount".to_string(), "sum".to_string())],
        None,
    );
    assert!(err.is_err());
}

#[test]
fn two_pivots_get_distinct_part_numbers() {
    let mut wb = workbook_with_source();
    wb.add_pivot_table(
        "Sales",
        "A1:C4",
        "Sales",
        "F1",
        &["Region".to_string()],
        &[],
        &[("Amount".to_string(), "sum".to_string())],
        None,
    )
    .unwrap();
    wb.add_pivot_table(
        "Sales",
        "A1:C4",
        "Sales",
        "F20",
        &["Product".to_string()],
        &[],
        &[("Amount".to_string(), "sum".to_string())],
        None,
    )
    .unwrap();

    let out = wb.save_to_bytes().unwrap();
    assert!(part_exists(&out, "xl/pivotTables/pivotTable1.xml"));
    assert!(part_exists(&out, "xl/pivotTables/pivotTable2.xml"));
    assert!(part_exists(&out, "xl/pivotCache/pivotCacheDefinition2.xml"));

    // Both pivots read back.
    let reloaded = Workbook::load_from_bytes(&out).unwrap();
    assert_eq!(reloaded.pivot_tables().len(), 2);
}

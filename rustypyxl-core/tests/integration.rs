//! Integration tests for rustypyxl-core.

use rustypyxl_core::{Workbook, CellValue};
use rustypyxl_core::autofilter::{AutoFilter, FilterColumn, CustomFilter, FilterOperator};
use rustypyxl_core::conditional::{ConditionalFormatting, ConditionalRule, ConditionalFormatType, ColorScale, DataBar, ConditionalColor};
use rustypyxl_core::table::{Table, TableStyle, TableColumn, TotalsRowFunction};
use rustypyxl_core::pagesetup::{PageSetup, PaperSize, Orientation, PageMargins, HeaderFooterSection};
use std::fs;

fn temp_file(name: &str) -> String {
    let dir = std::env::temp_dir().join("rustypyxl_tests");
    fs::create_dir_all(&dir).unwrap();
    dir.join(name).to_string_lossy().to_string()
}

#[test]
fn test_create_and_save_workbook() {
    let mut wb = Workbook::new();
    let ws = wb.create_sheet(Some("Test".to_string())).unwrap();
    assert_eq!(ws.title(), "Test");

    wb.set_cell_value_in_sheet("Test", 1, 1, CellValue::from("Hello")).unwrap();
    wb.set_cell_value_in_sheet("Test", 1, 2, CellValue::Number(42.0)).unwrap();
    wb.set_cell_value_in_sheet("Test", 1, 3, CellValue::Boolean(true)).unwrap();

    let path = temp_file("test_create.xlsx");
    wb.save(&path).unwrap();

    assert!(std::path::Path::new(&path).exists());
    fs::remove_file(&path).ok();
}

#[test]
fn test_roundtrip_basic() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("Data".to_string())).unwrap();

    // Add various cell types
    wb.set_cell_value_in_sheet("Data", 1, 1, CellValue::from("Name")).unwrap();
    wb.set_cell_value_in_sheet("Data", 1, 2, CellValue::from("Value")).unwrap();
    wb.set_cell_value_in_sheet("Data", 2, 1, CellValue::from("Item A")).unwrap();
    wb.set_cell_value_in_sheet("Data", 2, 2, CellValue::Number(100.0)).unwrap();
    wb.set_cell_value_in_sheet("Data", 3, 1, CellValue::from("Item B")).unwrap();
    wb.set_cell_value_in_sheet("Data", 3, 2, CellValue::Number(200.0)).unwrap();

    let path = temp_file("test_roundtrip.xlsx");
    wb.save(&path).unwrap();

    // Reload and verify
    let wb2 = Workbook::load(&path).unwrap();
    let ws2 = wb2.get_sheet_by_name("Data").unwrap();

    if let Some(cell) = ws2.get_cell(1, 1) {
        match &cell.value {
            CellValue::String(s) => assert_eq!(s.as_ref(), "Name"),
            _ => panic!("Expected string value"),
        }
    }

    if let Some(cell) = ws2.get_cell(2, 2) {
        match &cell.value {
            CellValue::Number(n) => assert!((n - 100.0).abs() < 0.001),
            _ => panic!("Expected number value"),
        }
    }

    fs::remove_file(&path).ok();
}

#[test]
fn test_formula_roundtrip() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("Formulas".to_string())).unwrap();

    wb.set_cell_value_in_sheet("Formulas", 1, 1, CellValue::Number(10.0)).unwrap();
    wb.set_cell_value_in_sheet("Formulas", 2, 1, CellValue::Number(20.0)).unwrap();
    wb.set_cell_value_in_sheet("Formulas", 3, 1, CellValue::Formula("SUM(A1:A2)".to_string())).unwrap();

    let path = temp_file("test_formula.xlsx");
    wb.save(&path).unwrap();

    let wb2 = Workbook::load(&path).unwrap();
    let ws2 = wb2.get_sheet_by_name("Formulas").unwrap();

    if let Some(cell) = ws2.get_cell(3, 1) {
        match &cell.value {
            CellValue::Formula(f) => assert_eq!(f, "SUM(A1:A2)"),
            _ => panic!("Expected formula value"),
        }
    }

    fs::remove_file(&path).ok();
}

#[test]
fn test_autofilter_creation() {
    let mut af = AutoFilter::new("A1:D100");
    assert_eq!(af.range, "A1:D100");
    assert!(af.columns.is_empty());

    // Add value filter
    af.add_filter(FilterColumn::values(0, vec!["Apple".to_string(), "Orange".to_string()]));
    assert_eq!(af.columns.len(), 1);

    // Add custom filter
    let custom = CustomFilter::new(FilterOperator::GreaterThan, "100")
        .and(FilterOperator::LessThan, "500");
    af.add_filter(FilterColumn::custom(1, custom));
    assert_eq!(af.columns.len(), 2);

    // Set sort
    af.sort_by(2, true);
    assert_eq!(af.sort_column, Some(2));
    assert!(af.sort_descending);
}

#[test]
fn test_conditional_formatting_color_scale() {
    let cs = ColorScale::two_color(
        ConditionalColor::rgb("FF0000"),
        ConditionalColor::rgb("00FF00"),
    );

    assert_eq!(cs.min_type, "min");
    assert_eq!(cs.max_type, "max");

    if let Some(ref rgb) = cs.min_color.rgb {
        assert_eq!(rgb, "FF0000");
    }
}

#[test]
fn test_conditional_formatting_data_bar() {
    let db = DataBar::new();

    assert!(db.show_value);
    assert!(db.gradient);
    assert_eq!(db.min_type, "min");
    assert_eq!(db.max_type, "max");
}

#[test]
fn test_conditional_formatting_rule() {
    let rule = ConditionalRule::cell_is(
        rustypyxl_core::conditional::ConditionalOperator::GreaterThan,
        "100",
    );

    assert_eq!(rule.rule_type, ConditionalFormatType::CellIs);
    assert!(rule.operator.is_some());
    assert_eq!(rule.formula1, Some("100".to_string()));
}

#[test]
fn test_conditional_formatting_creation() {
    let mut cf = ConditionalFormatting::new("A1:A100");
    assert_eq!(cf.range, "A1:A100");
    assert!(cf.rules.is_empty());

    let rule = ConditionalRule::cell_is(
        rustypyxl_core::conditional::ConditionalOperator::GreaterThan,
        "50",
    );
    cf.add_rule(rule);
    assert_eq!(cf.rules.len(), 1);
}

#[test]
fn test_table_creation() {
    let table = Table::new(1, "SalesData", "A1:D100");

    assert_eq!(table.id, 1);
    assert_eq!(table.name, "SalesData");
    assert_eq!(table.range, "A1:D100");
    assert!(table.header_row);
    assert!(!table.totals_row);
    assert!(table.auto_filter);
}

#[test]
fn test_table_with_headers() {
    let table = Table::with_headers(
        1,
        "Products",
        "A1:C10",
        &["Name", "Price", "Quantity"],
    );

    assert_eq!(table.columns.len(), 3);
    assert_eq!(table.columns[0].name, "Name");
    assert_eq!(table.columns[1].name, "Price");
    assert_eq!(table.columns[2].name, "Quantity");
}

#[test]
fn test_table_with_totals() {
    let mut table = Table::with_headers(1, "Data", "A1:B10", &["Item", "Value"]);
    table = table.with_totals_row();
    table.set_column_totals("Value", TotalsRowFunction::Sum);

    assert!(table.totals_row);
    assert_eq!(table.columns[1].totals_row_function, TotalsRowFunction::Sum);
}

#[test]
fn test_table_style() {
    let table = Table::new(1, "Test", "A1:B10")
        .with_style(TableStyle::blue());

    assert_eq!(table.style.style_name(), "TableStyleMedium2");

    let green = Table::new(2, "Test2", "A1:B10")
        .with_style(TableStyle::green());
    assert_eq!(green.style.style_name(), "TableStyleMedium7");
}

#[test]
fn test_table_structured_reference() {
    let table = Table::new(1, "Sales", "A1:C10");

    let ref1 = table.structured_ref(Some("Amount"), None);
    assert_eq!(ref1, "Sales[[Amount]]");

    let ref2 = table.structured_ref(Some("Amount"), Some("#Data"));
    assert_eq!(ref2, "Sales[[#Data],[Amount]]");

    let ref3 = table.structured_ref(None, Some("#Totals"));
    assert_eq!(ref3, "Sales[[#Totals]]");
}

#[test]
fn test_table_column_formula() {
    let col = TableColumn::new(1, "Total")
        .with_formula("[@Price]*[@Quantity]")
        .with_totals_function(TotalsRowFunction::Sum);

    assert_eq!(col.calculated_column_formula, Some("[@Price]*[@Quantity]".to_string()));
    assert_eq!(col.totals_row_function, TotalsRowFunction::Sum);
}

#[test]
fn test_page_setup_default() {
    let ps = PageSetup::new();

    assert_eq!(ps.paper_size, PaperSize::Letter);
    assert_eq!(ps.orientation, Orientation::Portrait);
    assert_eq!(ps.scale, 100);
    assert!(!ps.print_gridlines);
}

#[test]
fn test_page_setup_builder() {
    let ps = PageSetup::new()
        .with_paper_size(PaperSize::A4)
        .with_orientation(Orientation::Landscape)
        .with_scale(80)
        .print_gridlines()
        .center_on_page();

    assert_eq!(ps.paper_size, PaperSize::A4);
    assert_eq!(ps.orientation, Orientation::Landscape);
    assert_eq!(ps.scale, 80);
    assert!(ps.print_gridlines);
    assert!(ps.center_horizontally);
    assert!(ps.center_vertically);
}

#[test]
fn test_page_margins() {
    let margins = PageMargins::narrow();
    assert_eq!(margins.left, 0.25);
    assert_eq!(margins.right, 0.25);

    let wide = PageMargins::wide();
    assert_eq!(wide.left, 1.0);
    assert_eq!(wide.right, 1.0);

    let uniform = PageMargins::uniform(0.5);
    assert_eq!(uniform.left, 0.5);
    assert_eq!(uniform.top, 0.5);
}

#[test]
fn test_header_footer() {
    let section = HeaderFooterSection::new()
        .with_left("Page &P of &N")
        .with_center("Report Title")
        .with_right("&D");

    let s = section.to_string();
    assert!(s.contains("&L"));
    assert!(s.contains("&C"));
    assert!(s.contains("&R"));
    assert!(s.contains("&P"));
    assert!(s.contains("&D"));
}

#[test]
fn test_paper_size_codes() {
    assert_eq!(PaperSize::Letter.code(), 1);
    assert_eq!(PaperSize::A4.code(), 9);
    assert_eq!(PaperSize::Legal.code(), 5);
    assert_eq!(PaperSize::Tabloid.code(), 3);
}

#[test]
fn test_worksheet_autofilter() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("Test".to_string())).unwrap();

    // Get mutable reference to worksheet and set autofilter
    let ws = wb.get_sheet_by_name_mut("Test").unwrap();
    ws.set_auto_filter(AutoFilter::new("A1:D100"));

    assert!(ws.auto_filter.is_some());
}

#[test]
fn test_worksheet_conditional_formatting() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("Test".to_string())).unwrap();

    let ws = wb.get_sheet_by_name_mut("Test").unwrap();

    let mut cf = ConditionalFormatting::new("A1:A100");
    cf.add_rule(ConditionalRule::cell_is(
        rustypyxl_core::conditional::ConditionalOperator::GreaterThan,
        "50",
    ));
    ws.add_conditional_formatting(cf);

    assert_eq!(ws.conditional_formatting.len(), 1);
}

#[test]
fn test_worksheet_table() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("Test".to_string())).unwrap();

    let ws = wb.get_sheet_by_name_mut("Test").unwrap();

    let table = Table::with_headers(1, "TestTable", "A1:C10", &["Name", "Price", "Qty"]);
    ws.add_table(table);

    assert_eq!(ws.tables.len(), 1);
    assert_eq!(ws.tables[0].name, "TestTable");
}

#[test]
fn test_worksheet_page_setup() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("Test".to_string())).unwrap();

    let ws = wb.get_sheet_by_name_mut("Test").unwrap();

    let ps = PageSetup::new()
        .with_paper_size(PaperSize::A4)
        .with_orientation(Orientation::Landscape);
    ws.set_page_setup(ps);

    assert!(ws.page_setup.is_some());
}

#[test]
fn test_multiple_sheets() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("Sheet1".to_string())).unwrap();
    wb.create_sheet(Some("Sheet2".to_string())).unwrap();
    wb.create_sheet(Some("Sheet3".to_string())).unwrap();

    let names = wb.sheet_names();
    assert_eq!(names.len(), 3);
    assert!(names.contains(&"Sheet1".to_string()));
    assert!(names.contains(&"Sheet2".to_string()));
    assert!(names.contains(&"Sheet3".to_string()));

    // Save and reload
    let path = temp_file("test_multi_sheets.xlsx");
    wb.save(&path).unwrap();

    let wb2 = Workbook::load(&path).unwrap();
    let names2 = wb2.sheet_names();
    assert_eq!(names2.len(), 3);

    fs::remove_file(&path).ok();
}

#[test]
fn test_named_ranges() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("Data".to_string())).unwrap();

    wb.create_named_range("MyRange".to_string(), "Data!$A$1:$C$10".to_string()).unwrap();

    let range = wb.get_named_range("MyRange");
    assert!(range.is_some());
    assert_eq!(range.unwrap(), "Data!$A$1:$C$10");
}

#[test]
fn test_merged_cells() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("Test".to_string())).unwrap();

    let ws = wb.get_sheet_by_name_mut("Test").unwrap();
    ws.merge_cells("A1:C1");
    ws.merge_cells("A2:A5");

    assert_eq!(ws.merged_cells.len(), 2);

    let path = temp_file("test_merged.xlsx");
    wb.save(&path).unwrap();

    let wb2 = Workbook::load(&path).unwrap();
    let ws2 = wb2.get_sheet_by_name("Test").unwrap();
    assert_eq!(ws2.merged_cells.len(), 2);

    fs::remove_file(&path).ok();
}

#[test]
fn test_column_dimensions() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("Test".to_string())).unwrap();

    let ws = wb.get_sheet_by_name_mut("Test").unwrap();
    ws.set_column_width(1, 20.0);
    ws.set_column_width(2, 15.5);

    assert_eq!(ws.column_dimensions.get(&1), Some(&20.0));
    assert_eq!(ws.column_dimensions.get(&2), Some(&15.5));
}

#[test]
fn test_row_dimensions() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("Test".to_string())).unwrap();

    let ws = wb.get_sheet_by_name_mut("Test").unwrap();
    ws.set_row_height(1, 30.0);
    ws.set_row_height(5, 45.0);

    assert_eq!(ws.row_dimensions.get(&1), Some(&30.0));
    assert_eq!(ws.row_dimensions.get(&5), Some(&45.0));
}

/// Saving a sheet that uses protection, merges, validations, hyperlinks, and
/// page setup must emit worksheet children in CT_Worksheet schema order, or
/// Excel prompts to "repair" the file and strips the offending elements.
#[test]
fn test_worksheet_element_order_follows_schema() {
    use rustypyxl_core::DataValidation;
    use std::io::Read;

    let mut wb = Workbook::new();
    wb.create_sheet(Some("Test".to_string())).unwrap();
    wb.set_cell_value_in_sheet("Test", 1, 1, CellValue::from("data")).unwrap();
    wb.set_cell_hyperlink(2, 1, "#Test!A1".to_string()).unwrap();

    let ws = wb.get_sheet_by_name_mut("Test").unwrap();
    ws.enable_protection(Some("secret".to_string()));
    ws.merge_cells("B1:C1");
    ws.add_data_validation(3, 1, DataValidation::default());
    let mut ps = PageSetup::new();
    ps.print_gridlines = true;
    ws.set_page_setup(ps);

    let path = temp_file("test_element_order.xlsx");
    wb.save(&path).unwrap();

    let file = fs::File::open(&path).unwrap();
    let mut zip = zip::ZipArchive::new(file).unwrap();
    let mut sheet_xml = String::new();
    zip.by_name("xl/worksheets/sheet1.xml")
        .unwrap()
        .read_to_string(&mut sheet_xml)
        .unwrap();

    let order = [
        "<sheetData",
        "<sheetProtection",
        "<mergeCells",
        "<dataValidations",
        "<hyperlinks",
        "<printOptions",
        "<pageMargins",
        "<pageSetup ",
    ];
    let positions: Vec<usize> = order
        .iter()
        .map(|tag| {
            sheet_xml
                .find(tag)
                .unwrap_or_else(|| panic!("missing element {} in {}", tag, sheet_xml))
        })
        .collect();
    let mut sorted = positions.clone();
    sorted.sort_unstable();
    assert_eq!(positions, sorted, "worksheet elements out of schema order: {}", sheet_xml);

    // The password attribute must hold the legacy verifier hash, not plaintext.
    assert!(!sheet_xml.contains("password=\"secret\""));
    assert!(sheet_xml.contains("password=\"DAA7\""), "expected hashed password in {}", sheet_xml);

    fs::remove_file(&path).ok();
}

/// A cell with t="str" (cached formula string result) holds literal text, not
/// a shared-string index; matching only the first byte of the type attribute
/// used to resolve "123" against the shared-strings table.
#[test]
fn test_t_str_cells_are_literal_text_not_shared_index() {
    use std::io::Write;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1">
<c r="A1" t="s"><v>0</v></c>
<c r="B1" t="str"><v>0</v></c>
</row></sheetData></worksheet>"#;
    let shared_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>real shared</t></si></sst>"#;
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
    let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;
    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let mut zip_buf = std::io::Cursor::new(Vec::new());
    {
        let mut zw = zip::ZipWriter::new(&mut zip_buf);
        let opts: zip::write::FileOptions<'_, zip::write::ExtendedFileOptions> =
            zip::write::FileOptions::default();
        for (path, content) in [
            ("[Content_Types].xml", content_types),
            ("_rels/.rels", root_rels),
            ("xl/workbook.xml", workbook_xml),
            ("xl/_rels/workbook.xml.rels", rels_xml),
            ("xl/sharedStrings.xml", shared_xml),
            ("xl/worksheets/sheet1.xml", sheet_xml),
        ] {
            zw.start_file(path, opts.clone()).unwrap();
            zw.write_all(content.as_bytes()).unwrap();
        }
        zw.finish().unwrap();
    }

    let wb = Workbook::load_from_bytes(zip_buf.get_ref()).unwrap();
    let ws = wb.get_sheet_by_name("S").unwrap();

    match &ws.get_cell(1, 1).expect("A1 missing").value {
        CellValue::String(s) => assert_eq!(s.as_ref(), "real shared"),
        other => panic!("A1: expected shared string, got {:?}", other),
    }
    match &ws.get_cell(1, 2).expect("B1 missing").value {
        CellValue::String(s) => assert_eq!(s.as_ref(), "0"),
        other => panic!("B1: expected literal \"0\", got {:?}", other),
    }
}

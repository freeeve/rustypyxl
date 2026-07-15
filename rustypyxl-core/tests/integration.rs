//! Integration tests for rustypyxl.

use rustypyxl::autofilter::{AutoFilter, CustomFilter, FilterColumn, FilterOperator};
use rustypyxl::conditional::{
    ColorScale, ConditionalColor, ConditionalFormatType, ConditionalFormatting,
    ConditionalOperator, ConditionalRule, DataBar,
};
use rustypyxl::pagesetup::{
    HeaderFooterSection, Orientation, PageMargins, PageSetup, PaperSize,
};
use rustypyxl::table::{Table, TableColumn, TableStyle, TotalsRowFunction};
use rustypyxl::{CellValue, Workbook};
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

    wb.set_cell_value_in_sheet("Test", 1, 1, CellValue::from("Hello"))
        .unwrap();
    wb.set_cell_value_in_sheet("Test", 1, 2, CellValue::Number(42.0))
        .unwrap();
    wb.set_cell_value_in_sheet("Test", 1, 3, CellValue::Boolean(true))
        .unwrap();

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
    wb.set_cell_value_in_sheet("Data", 1, 1, CellValue::from("Name"))
        .unwrap();
    wb.set_cell_value_in_sheet("Data", 1, 2, CellValue::from("Value"))
        .unwrap();
    wb.set_cell_value_in_sheet("Data", 2, 1, CellValue::from("Item A"))
        .unwrap();
    wb.set_cell_value_in_sheet("Data", 2, 2, CellValue::Number(100.0))
        .unwrap();
    wb.set_cell_value_in_sheet("Data", 3, 1, CellValue::from("Item B"))
        .unwrap();
    wb.set_cell_value_in_sheet("Data", 3, 2, CellValue::Number(200.0))
        .unwrap();

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

    wb.set_cell_value_in_sheet("Formulas", 1, 1, CellValue::Number(10.0))
        .unwrap();
    wb.set_cell_value_in_sheet("Formulas", 2, 1, CellValue::Number(20.0))
        .unwrap();
    wb.set_cell_value_in_sheet(
        "Formulas",
        3,
        1,
        CellValue::Formula("SUM(A1:A2)".to_string()),
    )
    .unwrap();

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
    af.add_filter(FilterColumn::values(
        0,
        vec!["Apple".to_string(), "Orange".to_string()],
    ));
    assert_eq!(af.columns.len(), 1);

    // Add custom filter
    let custom =
        CustomFilter::new(FilterOperator::GreaterThan, "100").and(FilterOperator::LessThan, "500");
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
        rustypyxl::conditional::ConditionalOperator::GreaterThan,
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
        rustypyxl::conditional::ConditionalOperator::GreaterThan,
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
    let table = Table::with_headers(1, "Products", "A1:C10", &["Name", "Price", "Quantity"]);

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
    let table = Table::new(1, "Test", "A1:B10").with_style(TableStyle::blue());

    assert_eq!(table.style.style_name(), "TableStyleMedium2");

    let green = Table::new(2, "Test2", "A1:B10").with_style(TableStyle::green());
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

    assert_eq!(
        col.calculated_column_formula,
        Some("[@Price]*[@Quantity]".to_string())
    );
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
        rustypyxl::conditional::ConditionalOperator::GreaterThan,
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

    wb.create_named_range("MyRange".to_string(), "Data!$A$1:$C$10".to_string())
        .unwrap();

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
    use rustypyxl::DataValidation;
    use std::io::Read;

    let mut wb = Workbook::new();
    wb.create_sheet(Some("Test".to_string())).unwrap();
    wb.set_cell_value_in_sheet("Test", 1, 1, CellValue::from("data"))
        .unwrap();
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
    assert_eq!(
        positions, sorted,
        "worksheet elements out of schema order: {}",
        sheet_xml
    );

    // The password attribute must hold the legacy verifier hash, not plaintext.
    assert!(!sheet_xml.contains("password=\"secret\""));
    assert!(
        sheet_xml.contains("password=\"DAA7\""),
        "expected hashed password in {}",
        sheet_xml
    );

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

/// Load+save must preserve workbook structure: hidden sheets, active tab,
/// freeze panes, autofilter, data validations, page setup, hyperlinks, and
/// tables. Before this round-trip support existed, save silently stripped
/// all of these from any loaded file.
#[test]
fn test_roundtrip_preserves_structure() {
    use rustypyxl::pagesetup::PageMargins;
    use rustypyxl::{DataValidation, SheetVisibility};

    let mut wb = Workbook::new();
    wb.create_sheet(Some("Main".to_string())).unwrap();
    wb.create_sheet(Some("Secret".to_string())).unwrap();
    wb.set_cell_value_in_sheet("Main", 1, 1, CellValue::from("data"))
        .unwrap();
    wb.set_cell_value_in_sheet("Secret", 1, 1, CellValue::from("hidden data"))
        .unwrap();
    wb.active_sheet = 0;

    {
        let ws = wb.get_sheet_by_name_mut("Secret").unwrap();
        ws.visibility = SheetVisibility::Hidden;
    }
    {
        let ws = wb.get_sheet_by_name_mut("Main").unwrap();
        ws.set_freeze_panes(Some("B2".to_string()));
        ws.set_auto_filter(AutoFilter::new("A1:C10"));

        let dv = DataValidation {
            validation_type: "list".to_string(),
            formula1: Some("\"Yes,No\"".to_string()),
            sqref: Some("D2:D10".to_string()),
            ..Default::default()
        };
        ws.add_data_validation(2, 4, dv);

        let mut ps = PageSetup::new();
        ps.orientation = Orientation::Landscape;
        ps.paper_size = PaperSize::A4;
        ps.scale = 80;
        ps.print_gridlines = true;
        ps.margins = PageMargins {
            left: 1.5,
            right: 0.25,
            top: 2.0,
            bottom: 0.5,
            header: 0.1,
            footer: 0.9,
        };
        ws.set_page_setup(ps);

        ws.set_cell_hyperlink(3, 1, "https://example.com/page?a=1&b=2".to_string());
        ws.set_cell_hyperlink(4, 1, "#Secret!A1".to_string());

        let table = Table::with_headers(1, "MyTable", "F1:G3", &["Alpha", "Beta"]);
        ws.add_table(table);
    }

    let path = temp_file("test_roundtrip_structure.xlsx");
    wb.save(&path).unwrap();

    let wb2 = Workbook::load(&path).unwrap();
    assert_eq!(wb2.active_sheet, 0);

    let secret = wb2.get_sheet_by_name("Secret").unwrap();
    assert_eq!(
        secret.visibility,
        SheetVisibility::Hidden,
        "hidden sheet became visible"
    );

    let main = wb2.get_sheet_by_name("Main").unwrap();
    assert_eq!(main.visibility, SheetVisibility::Visible);
    assert_eq!(
        main.freeze_panes.as_deref(),
        Some("B2"),
        "freeze panes lost"
    );

    let af = main.auto_filter.as_ref().expect("autofilter lost");
    assert_eq!(af.range, "A1:C10");

    let dv = main
        .data_validations
        .get(&(2, 4))
        .expect("data validation lost");
    assert_eq!(dv.validation_type, "list");
    assert_eq!(dv.formula1.as_deref(), Some("\"Yes,No\""));
    assert_eq!(
        dv.sqref.as_deref(),
        Some("D2:D10"),
        "validation range narrowed"
    );

    let ps = main.page_setup.as_ref().expect("page setup lost");
    assert_eq!(ps.orientation, Orientation::Landscape);
    assert_eq!(ps.paper_size, PaperSize::A4);
    assert_eq!(ps.scale, 80);
    assert!(ps.print_gridlines);
    assert_eq!(ps.margins.left, 1.5);
    assert_eq!(ps.margins.top, 2.0);

    let link = main.get_cell(3, 1).and_then(|c| c.hyperlink.clone());
    assert_eq!(
        link.as_deref(),
        Some("https://example.com/page?a=1&b=2"),
        "external hyperlink lost"
    );
    let internal = main.get_cell(4, 1).and_then(|c| c.hyperlink.clone());
    assert_eq!(
        internal.as_deref(),
        Some("#Secret!A1"),
        "internal hyperlink lost"
    );

    assert_eq!(main.tables.len(), 1, "table lost");
    let t = &main.tables[0];
    assert_eq!(t.name, "MyTable");
    assert_eq!(t.range, "F1:G3");
    assert_eq!(t.columns.len(), 2);
    assert_eq!(t.columns[0].name, "Alpha");
    assert!(t.header_row);

    // A second save+load cycle must not degrade anything further.
    let path2 = temp_file("test_roundtrip_structure2.xlsx");
    wb2.save(&path2).unwrap();
    let wb3 = Workbook::load(&path2).unwrap();
    assert_eq!(
        wb3.get_sheet_by_name("Secret").unwrap().visibility,
        SheetVisibility::Hidden
    );
    assert_eq!(wb3.get_sheet_by_name("Main").unwrap().tables.len(), 1);
    assert_eq!(
        wb3.get_sheet_by_name("Main")
            .unwrap()
            .get_cell(3, 1)
            .and_then(|c| c.hyperlink.clone())
            .as_deref(),
        Some("https://example.com/page?a=1&b=2")
    );

    fs::remove_file(&path).ok();
    fs::remove_file(&path2).ok();
}

/// Conditional formatting must round-trip with its differential formats
/// (dxfs): previously dxfId was never written, so rules matched but applied
/// no formatting, and nothing was parsed back on load.
#[test]
fn test_conditional_formatting_roundtrip_with_dxfs() {
    use rustypyxl::conditional::{ConditionalFormat, IconSet, IconSetStyle};
    use std::io::Read;

    let mut wb = Workbook::new();
    wb.create_sheet(Some("CF".to_string())).unwrap();
    for r in 1..=10 {
        wb.set_cell_value_in_sheet("CF", r, 1, CellValue::Number(r as f64 * 10.0))
            .unwrap();
    }

    {
        let ws = wb.get_sheet_by_name_mut("CF").unwrap();

        let mut cf1 = ConditionalFormatting::new("A1:A10");
        cf1.add_rule(
            ConditionalRule::cell_is(ConditionalOperator::GreaterThan, "50").with_format(
                ConditionalFormat::new()
                    .with_fill(ConditionalColor::rgb("FFFFC7CE"))
                    .with_font_color(ConditionalColor::rgb("FF9C0006"))
                    .with_bold(true),
            ),
        );
        cf1.add_rule(
            ConditionalRule::with_color_scale(ColorScale::red_yellow_green()).with_priority(2),
        );
        ws.add_conditional_formatting(cf1);

        let mut cf2 = ConditionalFormatting::new("B1:B10");
        cf2.add_rule(ConditionalRule::with_data_bar(
            DataBar::new().with_color(ConditionalColor::blue()),
        ));
        // Empty thresholds previously produced schema-invalid <iconSet/>
        cf2.add_rule(
            ConditionalRule::with_icon_set(IconSet::new(IconSetStyle::ThreeArrows))
                .with_priority(2),
        );
        cf2.add_rule(
            ConditionalRule::contains_text("err")
                .with_priority(3)
                .with_format(ConditionalFormat::new().with_fill(ConditionalColor::yellow())),
        );
        ws.add_conditional_formatting(cf2);
    }

    let path = temp_file("test_cf_roundtrip.xlsx");
    wb.save(&path).unwrap();

    // The sheet XML must wire rules to dxfs via dxfId, and styles.xml must
    // actually contain the dxf entries.
    {
        let file = fs::File::open(&path).unwrap();
        let mut zip = zip::ZipArchive::new(file).unwrap();
        let mut sheet_xml = String::new();
        zip.by_name("xl/worksheets/sheet1.xml")
            .unwrap()
            .read_to_string(&mut sheet_xml)
            .unwrap();
        assert!(
            sheet_xml.contains("dxfId=\"0\""),
            "missing dxfId in {}",
            sheet_xml
        );
        assert!(
            sheet_xml.contains("SEARCH("),
            "missing implied text-rule formula"
        );
        // 3 icons -> three cfvo thresholds at 0/33/66 percent
        assert!(sheet_xml.contains(r#"<iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="66"/></iconSet>"#),
            "icon set missing default thresholds: {}", sheet_xml);

        let mut styles_xml = String::new();
        zip.by_name("xl/styles.xml")
            .unwrap()
            .read_to_string(&mut styles_xml)
            .unwrap();
        assert!(
            styles_xml.contains("<dxfs count=\"2\">"),
            "dxfs missing in {}",
            styles_xml
        );
        assert!(styles_xml.contains("FFFFC7CE"));
    }

    // Load back and verify the model round-trips
    let wb2 = Workbook::load(&path).unwrap();
    let ws2 = wb2.get_sheet_by_name("CF").unwrap();
    assert_eq!(
        ws2.conditional_formatting.len(),
        2,
        "CF blocks lost on load"
    );

    let cf1 = &ws2.conditional_formatting[0];
    assert_eq!(cf1.range, "A1:A10");
    assert_eq!(cf1.rules.len(), 2);
    let rule = &cf1.rules[0];
    assert_eq!(rule.rule_type, ConditionalFormatType::CellIs);
    assert_eq!(rule.operator, Some(ConditionalOperator::GreaterThan));
    assert_eq!(rule.formula1.as_deref(), Some("50"));
    let fmt = rule.format.as_ref().expect("dxf format lost on load");
    assert_eq!(
        fmt.fill_color.as_ref().and_then(|c| c.rgb.as_deref()),
        Some("FFFFC7CE")
    );
    assert_eq!(
        fmt.font_color.as_ref().and_then(|c| c.rgb.as_deref()),
        Some("FF9C0006")
    );
    assert_eq!(fmt.bold, Some(true));

    let scale_rule = &cf1.rules[1];
    let cs = scale_rule.color_scale.as_ref().expect("color scale lost");
    assert!(cs.mid_color.is_some());
    assert_eq!(cs.min_type, "min");
    assert_eq!(cs.max_type, "max");

    let cf2 = &ws2.conditional_formatting[1];
    assert_eq!(cf2.rules.len(), 3);
    let db = cf2.rules[0].data_bar.as_ref().expect("data bar lost");
    assert_eq!(db.fill_color.rgb.as_deref(), Some("FF0000FF"));
    let is = cf2.rules[1].icon_set.as_ref().expect("icon set lost");
    assert_eq!(is.thresholds.len(), 3);
    let text_rule = &cf2.rules[2];
    assert_eq!(text_rule.rule_type, ConditionalFormatType::ContainsText);
    assert_eq!(text_rule.text.as_deref(), Some("err"));
    assert!(text_rule.format.is_some());

    // Second save+load cycle must be stable
    let path2 = temp_file("test_cf_roundtrip2.xlsx");
    wb2.save(&path2).unwrap();
    let wb3 = Workbook::load(&path2).unwrap();
    let ws3 = wb3.get_sheet_by_name("CF").unwrap();
    assert_eq!(ws3.conditional_formatting.len(), 2);
    assert_eq!(
        ws3.conditional_formatting[0].rules[0]
            .format
            .as_ref()
            .and_then(|f| f.fill_color.as_ref())
            .and_then(|c| c.rgb.as_deref()),
        Some("FFFFC7CE")
    );

    fs::remove_file(&path).ok();
    fs::remove_file(&path2).ok();
}

/// Defined-name scope, cached formula values, header/footer sections, and
/// the comments VML part must survive save (+ load where applicable).
#[test]
fn test_roundtrip_names_formulas_headers_comments() {
    use rustypyxl::pagesetup::HeaderFooterSection;
    use std::io::Read;

    let mut wb = Workbook::new();
    wb.create_sheet(Some("S1".to_string())).unwrap();
    wb.create_sheet(Some("S2".to_string())).unwrap();

    // Sheet-scoped + hidden defined names
    wb.create_named_range("GlobalName".to_string(), "S1!$A$1:$B$2".to_string())
        .unwrap();
    wb.named_ranges.push(rustypyxl::NamedRange {
        name: "LocalName".to_string(),
        range: "S2!$C$1".to_string(),
        local_sheet_id: Some(1),
        hidden: true,
    });

    {
        let ws = wb.get_sheet_by_name_mut("S1").unwrap();
        ws.set_cell_value(1, 1, CellValue::Number(2.0));
        ws.set_cell_value(1, 2, CellValue::Number(3.0));
        // Formula with a cached numeric result
        let cell = ws.get_or_create_cell_mut(1, 3);
        cell.value = CellValue::Formula("A1+B1".to_string());
        cell.cached_formula_value = Some("5".to_string());
        // Formula with a cached string result
        let cell = ws.get_or_create_cell_mut(2, 3);
        cell.value = CellValue::Formula("CONCATENATE(\"a\",\"b\")".to_string());
        cell.cached_formula_value = Some("ab".to_string());
        cell.data_type = Some("str");

        ws.set_cell_comment(3, 1, "a comment".to_string());

        let mut ps = PageSetup::new();
        ps.header_footer.odd_header = Some(
            HeaderFooterSection::new()
                .with_left("Lft")
                .with_center("Ctr & Co")
                .with_right("Rgt"),
        );
        ps.header_footer.odd_footer = Some(HeaderFooterSection::new().with_center("Page"));
        ws.set_page_setup(ps);
    }

    let path = temp_file("test_names_formulas.xlsx");
    wb.save(&path).unwrap();

    // Package-level checks: cached <v>, VML part, legacyDrawing wiring
    {
        let file = fs::File::open(&path).unwrap();
        let mut zip = zip::ZipArchive::new(file).unwrap();
        let mut sheet_xml = String::new();
        zip.by_name("xl/worksheets/sheet1.xml")
            .unwrap()
            .read_to_string(&mut sheet_xml)
            .unwrap();
        assert!(
            sheet_xml.contains("<f>A1+B1</f><v>5</v>"),
            "cached formula value not written: {}",
            sheet_xml
        );
        assert!(
            sheet_xml.contains("t=\"str\""),
            "cached string type missing"
        );
        assert!(
            sheet_xml.contains("<legacyDrawing r:id=\"rIdVml\"/>"),
            "legacyDrawing missing"
        );

        let mut vml = String::new();
        zip.by_name("xl/drawings/vmlDrawing1.vml")
            .unwrap()
            .read_to_string(&mut vml)
            .unwrap();
        assert!(vml.contains("ObjectType=\"Note\""));
        assert!(
            vml.contains("<x:Row>2</x:Row>"),
            "comment anchor row wrong: {}",
            vml
        );

        let mut ct = String::new();
        zip.by_name("[Content_Types].xml")
            .unwrap()
            .read_to_string(&mut ct)
            .unwrap();
        assert!(ct.contains("Extension=\"vml\""));
        assert!(ct.contains("comments+xml"));

        let mut wbxml = String::new();
        zip.by_name("xl/workbook.xml")
            .unwrap()
            .read_to_string(&mut wbxml)
            .unwrap();
        assert!(
            wbxml.contains("localSheetId=\"1\""),
            "sheet scope lost: {}",
            wbxml
        );
        assert!(wbxml.contains("hidden=\"1\""), "hidden flag lost");
    }

    // Model round-trip
    let wb2 = Workbook::load(&path).unwrap();
    let local = wb2
        .named_ranges
        .iter()
        .find(|nr| nr.name == "LocalName")
        .expect("scoped name lost");
    assert_eq!(local.local_sheet_id, Some(1));
    assert!(local.hidden);
    let global = wb2
        .named_ranges
        .iter()
        .find(|nr| nr.name == "GlobalName")
        .unwrap();
    assert_eq!(global.local_sheet_id, None);

    let ws2 = wb2.get_sheet_by_name("S1").unwrap();
    let c = ws2.get_cell(1, 3).expect("formula cell lost");
    assert!(matches!(&c.value, CellValue::Formula(f) if f == "A1+B1"));
    let cached = c.cached_formula_value.as_deref();
    assert!(
        cached == Some("5") || cached == Some("5.0"),
        "cached numeric formula value lost: {:?}",
        cached
    );
    let cs = ws2.get_cell(2, 3).expect("string formula cell lost");
    assert_eq!(cs.cached_formula_value.as_deref(), Some("ab"));

    let ps2 = ws2.page_setup.as_ref().expect("page setup lost");
    let hdr = ps2.header_footer.odd_header.as_ref().expect("header lost");
    assert_eq!(hdr.left.as_deref(), Some("Lft"));
    assert_eq!(hdr.center.as_deref(), Some("Ctr & Co"));
    assert_eq!(hdr.right.as_deref(), Some("Rgt"));
    let ftr = ps2.header_footer.odd_footer.as_ref().expect("footer lost");
    assert_eq!(ftr.center.as_deref(), Some("Page"));

    assert_eq!(
        ws2.get_cell(3, 1)
            .and_then(|c| c.comment.clone())
            .as_deref(),
        Some("a comment")
    );

    fs::remove_file(&path).ok();
}

/// Styles applied through the core Rust API must reach the saved file.
/// Previously the writer emitted s= only from style_index, which nothing in
/// core populated, so every core styling call produced unstyled output.
#[test]
fn test_core_styling_api_reaches_saved_file() {
    use rustypyxl::{Alignment, CellStyle, Fill, Font};
    use std::io::Read;

    let mut wb = Workbook::new();
    wb.create_sheet(Some("S".to_string())).unwrap();
    wb.set_cell_value_in_sheet("S", 1, 1, CellValue::from("styled"))
        .unwrap();
    wb.set_cell_value_in_sheet("S", 2, 1, CellValue::Number(0.5))
        .unwrap();
    wb.set_cell_value_in_sheet("S", 3, 1, CellValue::Number(0.25))
        .unwrap();

    {
        let ws = wb.get_sheet_by_name_mut("S").unwrap();
        let style = CellStyle::new()
            .with_font(Font::new().with_bold(true).with_color("FF0000"))
            .with_fill(Fill::solid("FFFF00"));
        ws.set_cell_style(1, 1, style);
        ws.set_cell_number_format(2, 1, "0.000%"); // custom (not a builtin id)
                                                   // Built-in date/time format must survive the id round-trip
        ws.set_cell_number_format(3, 1, "h:mm");
        ws.set_cell_alignment(3, 1, Alignment::new().with_horizontal("center"));
    }

    let path = temp_file("test_core_styles.xlsx");
    wb.save(&path).unwrap();

    {
        let file = fs::File::open(&path).unwrap();
        let mut zip = zip::ZipArchive::new(file).unwrap();
        let mut sheet_xml = String::new();
        zip.by_name("xl/worksheets/sheet1.xml")
            .unwrap()
            .read_to_string(&mut sheet_xml)
            .unwrap();
        assert!(
            sheet_xml.contains(r#"<c r="A1" s="#),
            "styled cell has no s= attribute: {}",
            sheet_xml
        );
        let mut styles_xml = String::new();
        zip.by_name("xl/styles.xml")
            .unwrap()
            .read_to_string(&mut styles_xml)
            .unwrap();
        assert!(
            styles_xml.contains("FFFF00"),
            "fill color missing from styles.xml"
        );
        assert!(
            styles_xml.contains("0.000%"),
            "custom number format missing from styles.xml"
        );
        assert!(
            styles_xml.contains(r#"numFmtId="20""#),
            "builtin h:mm xf missing"
        );
    }

    let wb2 = Workbook::load(&path).unwrap();
    let ws2 = wb2.get_sheet_by_name("S").unwrap();

    let c1 = ws2.get_cell(1, 1).expect("A1 missing");
    let style = c1.style.as_ref().expect("style lost on round-trip");
    let font = style.font.as_ref().expect("font lost");
    assert!(font.bold, "bold lost");
    let fill = style.fill.as_ref().expect("fill lost");
    assert!(
        format!("{:?}", fill).contains("FFFF00"),
        "fill color lost: {:?}",
        fill
    );

    let c2 = ws2.get_cell(2, 1).expect("A2 missing");
    let fmt = c2
        .number_format
        .clone()
        .or_else(|| c2.style.as_ref().and_then(|s| s.number_format.clone()));
    assert_eq!(fmt.as_deref(), Some("0.000%"), "number format lost");

    let c3 = ws2.get_cell(3, 1).expect("A3 missing");
    let fmt3 = c3
        .number_format
        .clone()
        .or_else(|| c3.style.as_ref().and_then(|s| s.number_format.clone()));
    assert_eq!(
        fmt3.as_deref(),
        Some("h:mm"),
        "builtin date format lost (id 20)"
    );
    let align = c3
        .style
        .as_ref()
        .and_then(|s| s.alignment.clone())
        .expect("alignment lost");
    assert_eq!(align.horizontal.as_deref(), Some("center"));

    fs::remove_file(&path).ok();
}

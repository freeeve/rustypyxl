//! Parser conformance tests for spec-legal XLSX shapes that are not what
//! Excel itself emits: cells with no `r` attribute, multi-run inline strings,
//! namespace-prefixed elements, and the 1904 date system.

use std::io::{Cursor, Write};

use rustypyxl_core::{CellValue, Workbook};
use zip::{write::SimpleFileOptions, ZipWriter};

const DEFAULT_WORKBOOK_PR: &str = "<workbookPr/>";

/// Assemble a minimal single-sheet xlsx around the given sheet XML body.
fn build_xlsx(sheet_xml: &str, shared_strings: Option<&str>, workbook_pr: &str) -> Vec<u8> {
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default();

    let workbook_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  {workbook_pr}
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#
    );

    let rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let content_types = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let mut add = |name: &str, body: &str| {
        zip.start_file(name, options).unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    };

    add("[Content_Types].xml", content_types);
    add("_rels/.rels", root_rels);
    add("xl/workbook.xml", &workbook_xml);
    add("xl/_rels/workbook.xml.rels", rels);
    add("xl/worksheets/sheet1.xml", sheet_xml);
    if let Some(sst) = shared_strings {
        add("xl/sharedStrings.xml", sst);
    }

    zip.finish().unwrap().into_inner()
}

fn load_sheet_xml(sheet_xml: &str) -> Workbook {
    Workbook::load_from_bytes(&build_xlsx(sheet_xml, None, DEFAULT_WORKBOOK_PR)).unwrap()
}

/// `r` is optional on `<c>`: the column is implied by the cell's index within
/// the row. Without inference both cells collapse onto one key.
#[test]
fn cells_without_r_attribute_take_successive_columns() {
    let wb = load_sheet_xml(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2"><c><v>1</v></c><c><v>2</v></c><c><v>3</v></c></row>
  </sheetData>
</worksheet>"#,
    );
    let ws = wb.get_sheet_by_name("Sheet1").unwrap();

    assert_eq!(ws.get_cell_value(2, 1), Some(&CellValue::Number(1.0)));
    assert_eq!(ws.get_cell_value(2, 2), Some(&CellValue::Number(2.0)));
    assert_eq!(ws.get_cell_value(2, 3), Some(&CellValue::Number(3.0)));
}

/// A cell that does carry `r` re-anchors the implied position for the cells
/// after it, so gaps in a row are honoured.
#[test]
fn explicit_r_reanchors_the_implied_column() {
    let wb = load_sheet_xml(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c><v>1</v></c><c r="D1"><v>4</v></c><c><v>5</v></c></row>
  </sheetData>
</worksheet>"#,
    );
    let ws = wb.get_sheet_by_name("Sheet1").unwrap();

    assert_eq!(ws.get_cell_value(1, 1), Some(&CellValue::Number(1.0)));
    assert_eq!(ws.get_cell_value(1, 4), Some(&CellValue::Number(4.0)));
    assert_eq!(ws.get_cell_value(1, 5), Some(&CellValue::Number(5.0)));
    assert_eq!(ws.get_cell_value(1, 2), None);
}

/// `r` is optional on `<row>` too.
#[test]
fn rows_without_r_attribute_take_successive_rows() {
    let wb = load_sheet_xml(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row><c><v>1</v></c></row>
    <row><c><v>2</v></c></row>
  </sheetData>
</worksheet>"#,
    );
    let ws = wb.get_sheet_by_name("Sheet1").unwrap();

    assert_eq!(ws.get_cell_value(1, 1), Some(&CellValue::Number(1.0)));
    assert_eq!(ws.get_cell_value(2, 1), Some(&CellValue::Number(2.0)));
}

/// Self-closing cells advance the implied column like any other cell, instead
/// of being dropped.
#[test]
fn self_closing_cells_without_r_hold_their_column() {
    let wb = load_sheet_xml(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c/><c><v>2</v></c></row>
  </sheetData>
</worksheet>"#,
    );
    let ws = wb.get_sheet_by_name("Sheet1").unwrap();

    assert_eq!(ws.get_cell_value(1, 1), Some(&CellValue::Empty));
    assert_eq!(ws.get_cell_value(1, 2), Some(&CellValue::Number(2.0)));
}

/// Rich-text inline strings split their text across `<r><t>` runs; every run
/// contributes, not just the last.
#[test]
fn inline_string_runs_are_concatenated() {
    let wb = load_sheet_xml(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><r><t>Hello </t></r><r><t>World</t></r></is></c>
      <c r="B1" t="inlineStr"><is><t>Plain</t></is></c>
    </row>
  </sheetData>
</worksheet>"#,
    );
    let ws = wb.get_sheet_by_name("Sheet1").unwrap();

    assert_eq!(
        ws.get_cell_value(1, 1),
        Some(&CellValue::String("Hello World".into()))
    );
    assert_eq!(
        ws.get_cell_value(1, 2),
        Some(&CellValue::String("Plain".into()))
    );
}

/// Shared strings are also rich text; runs there already concatenated, and the
/// inline fix must not regress that.
#[test]
fn shared_string_runs_are_concatenated() {
    let sst = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><r><t>Rich </t></r><r><t>Text</t></r></si>
</sst>"#;
    let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>
</worksheet>"#;

    let wb = Workbook::load_from_bytes(&build_xlsx(sheet, Some(sst), DEFAULT_WORKBOOK_PR)).unwrap();
    let ws = wb.get_sheet_by_name("Sheet1").unwrap();

    assert_eq!(
        ws.get_cell_value(1, 1),
        Some(&CellValue::String("Rich Text".into()))
    );
}

/// Attribute order is not significant in XML: `ht` before `r` must still apply
/// the height to the right row.
#[test]
fn row_height_is_independent_of_attribute_order() {
    let wb = load_sheet_xml(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row ht="42.5" r="1" customHeight="1"><c r="A1"><v>1</v></c></row>
    <row r="2" ht="21.5" customHeight="1"><c r="A2"><v>2</v></c></row>
    <row r="3" ht="15.5" customHeight="1"/>
  </sheetData>
</worksheet>"#,
    );
    let ws = wb.get_sheet_by_name("Sheet1").unwrap();

    assert_eq!(ws.get_row_height(1), Some(42.5));
    assert_eq!(ws.get_row_height(2), Some(21.5));
    assert_eq!(ws.get_row_height(3), Some(15.5));
}

/// Producers may namespace-prefix every element; the local name is what counts.
#[test]
fn namespace_prefixed_elements_are_parsed() {
    let sst = r#"<?xml version="1.0" encoding="UTF-8"?>
<x:sst xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <x:si><x:t>Shared</x:t></x:si>
</x:sst>"#;
    let sheet = r#"<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:sheetData>
    <x:row r="1">
      <x:c r="A1" t="s"><x:v>0</x:v></x:c>
      <x:c r="B1"><x:v>7</x:v></x:c>
    </x:row>
  </x:sheetData>
</x:worksheet>"#;

    let wb = Workbook::load_from_bytes(&build_xlsx(sheet, Some(sst), DEFAULT_WORKBOOK_PR)).unwrap();
    let ws = wb.get_sheet_by_name("Sheet1").unwrap();

    assert_eq!(
        ws.get_cell_value(1, 1),
        Some(&CellValue::String("Shared".into()))
    );
    assert_eq!(ws.get_cell_value(1, 2), Some(&CellValue::Number(7.0)));
}

/// The cached result of a formula is echoed back on save; it must not be
/// reformatted (an integer `5` would come back out as `5.0`).
#[test]
fn cached_formula_value_keeps_its_source_text() {
    let wb = load_sheet_xml(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><f>2+3</f><v>5</v></c></row>
  </sheetData>
</worksheet>"#,
    );
    let ws = wb.get_sheet_by_name("Sheet1").unwrap();
    let cell = ws.get_cell(1, 1).unwrap();

    assert_eq!(cell.value, CellValue::Formula("2+3".to_string()));
    assert_eq!(cell.cached_formula_value.as_deref(), Some("5"));

    // And it survives the round-trip rather than turning into "5.0".
    let reloaded = Workbook::load_from_bytes(&wb.save_to_bytes().unwrap()).unwrap();
    let reloaded_cell = reloaded
        .get_sheet_by_name("Sheet1")
        .unwrap()
        .get_cell(1, 1)
        .unwrap();
    assert_eq!(reloaded_cell.cached_formula_value.as_deref(), Some("5"));
}

/// A formula whose cached result is a shared string still resolves through the
/// shared-strings table.
#[test]
fn cached_formula_value_resolves_shared_strings() {
    let sst = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Cached</t></si>
</sst>"#;
    let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="s"><f>A2</f><v>0</v></c></row></sheetData>
</worksheet>"#;

    let wb = Workbook::load_from_bytes(&build_xlsx(sheet, Some(sst), DEFAULT_WORKBOOK_PR)).unwrap();
    let cell = wb
        .get_sheet_by_name("Sheet1")
        .unwrap()
        .get_cell(1, 1)
        .unwrap();

    assert_eq!(cell.cached_formula_value.as_deref(), Some("Cached"));
}

/// Serials in a 1904-system file mean a different date than in a 1900-system
/// file, so the flag has to survive a round-trip.
#[test]
fn date1904_flag_round_trips() {
    let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>"#;

    let wb = Workbook::load_from_bytes(&build_xlsx(sheet, None, r#"<workbookPr date1904="1"/>"#))
        .unwrap();
    assert!(wb.date1904);

    let reloaded = Workbook::load_from_bytes(&wb.save_to_bytes().unwrap()).unwrap();
    assert!(reloaded.date1904, "date1904 must survive a save/load cycle");

    // The default 1900 system stays off.
    let wb_1900 = Workbook::load_from_bytes(&build_xlsx(sheet, None, DEFAULT_WORKBOOK_PR)).unwrap();
    assert!(!wb_1900.date1904);
    let reloaded_1900 = Workbook::load_from_bytes(&wb_1900.save_to_bytes().unwrap()).unwrap();
    assert!(!reloaded_1900.date1904);
}

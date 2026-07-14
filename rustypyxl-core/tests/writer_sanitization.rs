//! The writer must never emit XML 1.0 control characters, whatever the caller
//! puts in a string. quick-xml escapes `< > & " '` but passes control chars
//! through, and Excel rejects the resulting file as corrupt -- so every
//! user-supplied string has to be stripped on the way out.

use std::io::Read;

use rustypyxl::conditional::{
    ConditionalColor, ConditionalFormat, ConditionalFormatting, ConditionalOperator,
    ConditionalRule,
};
use rustypyxl::pagesetup::{HeaderFooterSection, PageSetup};
use rustypyxl::table::{Table, TableColumn};
use rustypyxl::worksheet::DataValidation;
use rustypyxl::{CellValue, NamedRange, Workbook};
use zip::ZipArchive;

/// A C0 control char that is illegal in XML 1.0 even when escaped.
const DIRTY: &str = "bad\u{1}text";
const CLEAN: &str = "badtext";

/// Every XML part in the archive, as text.
fn xml_parts(bytes: &[u8]) -> Vec<(String, String)> {
    let mut archive = ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut parts = Vec::new();
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let name = file.name().to_string();
        let mut body = String::new();
        if file.read_to_string(&mut body).is_ok() {
            parts.push((name, body));
        }
    }
    parts
}

/// A workbook exercising every part that takes a user-supplied string.
fn dirty_workbook() -> Workbook {
    let mut wb = Workbook::new();
    wb.named_ranges.push(NamedRange {
        name: format!("Name{}", DIRTY),
        range: format!("Sheet1!A1:B2{}", DIRTY),
        local_sheet_id: None,
        hidden: false,
    });

    let ws = wb.create_sheet(Some("Sheet1".to_string())).unwrap();
    ws.set_cell_value(1, 1, CellValue::String(DIRTY.into()));
    ws.set_cell_comment(1, 1, DIRTY.to_string());

    let validation = DataValidation {
        validation_type: "list".to_string(),
        formula1: Some(format!("\"{}\"", DIRTY)),
        formula2: Some(DIRTY.to_string()),
        ..Default::default()
    };
    ws.add_data_validation(2, 1, validation);

    let mut cf = ConditionalFormatting::new("A1:A10");
    cf.add_rule(
        ConditionalRule::cell_is(ConditionalOperator::GreaterThan, DIRTY)
            .with_format(ConditionalFormat::new().with_fill(ConditionalColor::green())),
    );
    ws.add_conditional_formatting(cf);

    let mut table = Table::new(1, "Table1", "A1:B2");
    table.name = format!("T{}", DIRTY);
    table.display_name = format!("T{}", DIRTY);
    table.columns = vec![
        TableColumn::new(1, &format!("Col{}", DIRTY)),
        TableColumn::new(2, "Total").with_formula(format!("SUM({})", DIRTY)),
    ];
    ws.add_table(table);

    let mut page_setup = PageSetup::new();
    page_setup.header_footer.odd_header = Some(HeaderFooterSection {
        left: Some(DIRTY.to_string()),
        center: None,
        right: None,
    });
    page_setup.header_footer.odd_footer = Some(HeaderFooterSection {
        left: None,
        center: Some(DIRTY.to_string()),
        right: None,
    });
    ws.set_page_setup(page_setup);

    wb
}

/// No part of the archive may contain an illegal control character, and the
/// file it produces must still load.
#[test]
fn control_characters_are_stripped_from_every_part() {
    let wb = dirty_workbook();
    let bytes = wb.save_to_bytes().unwrap();

    for (name, body) in xml_parts(&bytes) {
        let illegal: Vec<char> = body
            .chars()
            .filter(|&c| (c as u32) < 0x20 && !matches!(c, '\t' | '\n' | '\r'))
            .collect();
        assert!(
            illegal.is_empty(),
            "{} contains illegal XML control chars: {:?}",
            name,
            illegal
        );
    }

    // And the stripped text is what survives, rather than the value being lost.
    let reloaded = Workbook::load_from_bytes(&bytes).unwrap();
    let ws = reloaded.get_sheet_by_name("Sheet1").unwrap();
    assert_eq!(
        ws.get_cell_value(1, 1),
        Some(&CellValue::String(CLEAN.into()))
    );
    assert_eq!(reloaded.named_ranges[0].name, format!("Name{}", CLEAN));
}

/// The specific parts named in the task each carry their sanitized string.
#[test]
fn sanitized_text_reaches_each_part() {
    let bytes = dirty_workbook().save_to_bytes().unwrap();
    let parts = xml_parts(&bytes);
    let part = |needle: &str| -> String {
        parts
            .iter()
            .find(|(name, _)| name.contains(needle))
            .unwrap_or_else(|| panic!("part {} missing", needle))
            .1
            .clone()
    };

    assert!(part("comments/comment1.xml").contains(CLEAN));
    assert!(part("tables/table1.xml").contains(&format!("Col{}", CLEAN)));
    assert!(part("tables/table1.xml").contains(&format!("SUM({})", CLEAN)));
    assert!(part("workbook.xml").contains(&format!("Name{}", CLEAN)));

    let sheet = part("worksheets/sheet1.xml");
    assert!(
        sheet.contains("formula1"),
        "data validation formula written"
    );
    assert!(sheet.contains(CLEAN));
}

/// `<t>` needs xml:space="preserve" or a conforming consumer may trim the
/// significant whitespace away.
#[test]
fn significant_whitespace_is_marked_preserve() {
    let mut wb = Workbook::new();
    let ws = wb.create_sheet(Some("Sheet1".to_string())).unwrap();
    ws.set_cell_value(1, 1, CellValue::String("  padded  ".into()));
    ws.set_cell_value(2, 1, CellValue::String("unpadded".into()));

    let bytes = wb.save_to_bytes().unwrap();
    let parts = xml_parts(&bytes);
    let sst = &parts
        .iter()
        .find(|(name, _)| name.contains("sharedStrings"))
        .unwrap()
        .1;

    assert!(
        sst.contains(r#"<t xml:space="preserve">  padded  </t>"#),
        "padded string must carry xml:space=preserve, got: {}",
        sst
    );
    assert!(
        sst.contains("<t>unpadded</t>"),
        "unpadded string should not carry the attribute"
    );

    let reloaded = Workbook::load_from_bytes(&bytes).unwrap();
    assert_eq!(
        reloaded
            .get_sheet_by_name("Sheet1")
            .unwrap()
            .get_cell_value(1, 1),
        Some(&CellValue::String("  padded  ".into()))
    );
}

/// sst `count` is the number of references; `uniqueCount` the table size.
#[test]
fn shared_string_count_is_total_references() {
    let mut wb = Workbook::new();
    let ws = wb.create_sheet(Some("Sheet1".to_string())).unwrap();
    // Three cells, two distinct strings.
    ws.set_cell_value(1, 1, CellValue::String("a".into()));
    ws.set_cell_value(2, 1, CellValue::String("a".into()));
    ws.set_cell_value(3, 1, CellValue::String("b".into()));

    let bytes = wb.save_to_bytes().unwrap();
    let parts = xml_parts(&bytes);
    let sst = &parts
        .iter()
        .find(|(name, _)| name.contains("sharedStrings"))
        .unwrap()
        .1;

    assert!(sst.contains(r#"count="3""#), "count is total refs: {}", sst);
    assert!(
        sst.contains(r#"uniqueCount="2""#),
        "uniqueCount is table size: {}",
        sst
    );
}

/// dxf number formats share the workbook's numFmtId space, so they must be
/// allocated above the custom formats already in use rather than from a fixed
/// 200+ range that a loaded file may already occupy.
#[test]
fn dxf_num_fmt_ids_do_not_collide_with_custom_formats() {
    let mut wb = Workbook::new();
    // A loaded file whose custom formats already occupy the 200+ range the
    // dxf writer used to hardcode.
    wb.styles.num_fmts.push((200, "0.000".to_string()));
    wb.styles.num_fmts.push((201, "0.0000".to_string()));

    let ws = wb.create_sheet(Some("Sheet1".to_string())).unwrap();
    let mut cf = ConditionalFormatting::new("A1:A10");
    let mut format = ConditionalFormat::new();
    format.number_format = Some("0.00%".to_string());
    cf.add_rule(
        ConditionalRule::cell_is(ConditionalOperator::GreaterThan, "1").with_format(format),
    );
    ws.add_conditional_formatting(cf);

    let bytes = wb.save_to_bytes().unwrap();
    let parts = xml_parts(&bytes);
    let styles = &parts
        .iter()
        .find(|(name, _)| name.contains("styles.xml"))
        .unwrap()
        .1;

    // Every numFmtId in the file must be distinct.
    let ids: Vec<&str> = styles
        .match_indices("numFmtId=\"")
        .map(|(i, m)| {
            let rest = &styles[i + m.len()..];
            &rest[..rest.find('"').unwrap()]
        })
        .filter(|id| *id != "0")
        .collect();
    let mut sorted = ids.clone();
    sorted.sort_unstable();
    sorted.dedup();
    assert_eq!(
        sorted.len(),
        ids.len(),
        "numFmtId collision between dxfs and custom formats: {:?}",
        ids
    );
}

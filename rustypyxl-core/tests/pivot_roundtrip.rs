//! Pivot tables are preserved verbatim across a load/save round-trip. rustypyxl
//! does not model pivot tables; it captures their parts, the workbook
//! `<pivotCaches>` element, and the relationships tying them together, and
//! re-emits them on save so they are not silently dropped.

use rustypyxl::Workbook;
use std::io::{Cursor, Read, Write};
use zip::write::{FileOptions, ZipWriter};
use zip::{CompressionMethod, ZipArchive};

/// Build a minimal xlsx that contains one sheet and one pivot table wired to a
/// pivot cache. Each pivot part carries a unique marker so the test can prove it
/// survived byte-for-byte.
fn xlsx_with_pivot() -> Vec<u8> {
    let parts: &[(&str, &str)] = &[
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
<Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
</Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
        ),
        (
            "xl/workbook.xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
<pivotCaches><pivotCache cacheId="1" r:id="rId2"/></pivotCaches>
</workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>"#,
        ),
        (
            "xl/styles.xml",
            r#"<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#,
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/></worksheet>"#,
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>"#,
        ),
        (
            "xl/pivotCache/pivotCacheDefinition1.xml",
            r#"<?xml version="1.0"?><pivotCacheDefinition marker="CACHE_DEF_MARKER"/>"#,
        ),
        (
            "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
            r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>"#,
        ),
        (
            "xl/pivotCache/pivotCacheRecords1.xml",
            r#"<?xml version="1.0"?><pivotCacheRecords marker="CACHE_REC_MARKER"/>"#,
        ),
        (
            "xl/pivotTables/pivotTable1.xml",
            r#"<?xml version="1.0"?><pivotTableDefinition cacheId="1" marker="PIVOT_TABLE_MARKER"/>"#,
        ),
        (
            "xl/pivotTables/_rels/pivotTable1.xml.rels",
            r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>"#,
        ),
    ];

    let mut cursor = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut cursor);
        let opts: FileOptions<'_, ()> =
            FileOptions::default().compression_method(CompressionMethod::Deflated);
        for (name, content) in parts {
            zip.start_file(*name, opts).unwrap();
            zip.write_all(content.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
    }
    cursor.into_inner()
}

fn part(bytes: &[u8], name: &str) -> Option<String> {
    let mut zip = ZipArchive::new(Cursor::new(bytes.to_vec())).unwrap();
    let mut f = zip.by_name(name).ok()?;
    let mut s = String::new();
    f.read_to_string(&mut s).unwrap();
    Some(s)
}

#[test]
fn pivot_parts_are_captured_on_load() {
    let wb = Workbook::load_from_bytes(&xlsx_with_pivot()).unwrap();

    assert!(!wb.pivots.is_empty(), "pivot artifacts captured");
    // All five pivot XML/rels parts are captured.
    assert_eq!(wb.pivots.parts.len(), 5);
    assert!(wb
        .pivots
        .workbook_caches_xml
        .as_deref()
        .unwrap()
        .contains("pivotCache"));
    assert_eq!(wb.pivots.workbook_rels.len(), 1);

    // The sheet's pivotTable relationship is preserved on the worksheet.
    let ws = wb.get_sheet_by_name("Data").unwrap();
    assert_eq!(ws.pivot_rels.len(), 1);
    assert!(ws.pivot_rels[0].1.ends_with("/pivotTable"));
}

#[test]
fn pivot_parts_survive_save() {
    let wb = Workbook::load_from_bytes(&xlsx_with_pivot()).unwrap();
    let out = wb.save_to_bytes().unwrap();

    // Every pivot part is re-emitted with its content intact.
    assert!(part(&out, "xl/pivotCache/pivotCacheDefinition1.xml")
        .unwrap()
        .contains("CACHE_DEF_MARKER"));
    assert!(part(&out, "xl/pivotCache/pivotCacheRecords1.xml")
        .unwrap()
        .contains("CACHE_REC_MARKER"));
    assert!(part(&out, "xl/pivotTables/pivotTable1.xml")
        .unwrap()
        .contains("PIVOT_TABLE_MARKER"));
    assert!(part(&out, "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels").is_some());
    assert!(part(&out, "xl/pivotTables/_rels/pivotTable1.xml.rels").is_some());

    // The workbook still declares the pivot cache, with a renumbered rel id.
    let wbxml = part(&out, "xl/workbook.xml").unwrap();
    assert!(wbxml.contains("<pivotCaches>"));
    assert!(
        wbxml.contains(r#"r:id="rIdPivotCache1""#),
        "rel id renumbered"
    );

    let wbrels = part(&out, "xl/_rels/workbook.xml.rels").unwrap();
    assert!(wbrels.contains(r#"Id="rIdPivotCache1""#));
    assert!(wbrels.contains("pivotCache/pivotCacheDefinition1.xml"));

    // The sheet keeps its pivotTable relationship.
    let sheetrels = part(&out, "xl/worksheets/_rels/sheet1.xml.rels").unwrap();
    assert!(sheetrels.contains("../pivotTables/pivotTable1.xml"));

    // Content types declare the pivot parts.
    let ct = part(&out, "[Content_Types].xml").unwrap();
    assert!(ct.contains("/xl/pivotTables/pivotTable1.xml"));
    assert!(ct.contains("pivotCacheDefinition+xml"));
}

#[test]
fn pivot_survives_a_double_round_trip() {
    let once = Workbook::load_from_bytes(&xlsx_with_pivot())
        .unwrap()
        .save_to_bytes()
        .unwrap();
    // Load the saved file again: the pivot must still be captured.
    let twice = Workbook::load_from_bytes(&once).unwrap();
    assert_eq!(twice.pivots.parts.len(), 5);
    assert_eq!(twice.get_sheet_by_name("Data").unwrap().pivot_rels.len(), 1);
}

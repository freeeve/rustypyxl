//! The read-only pivot model exposes a loaded pivot table's source range, cache
//! fields, and row/column/data field placements (openpyxl-level read support).

use rustypyxl::pivot::PivotDataField;
use rustypyxl::Workbook;
use std::io::{Cursor, Write};
use zip::write::{FileOptions, ZipWriter};
use zip::CompressionMethod;

/// A minimal xlsx whose sheet "Sales" (A1:C6, columns Region/Product/Amount)
/// backs one pivot table: Region in rows, Product in columns, Sum of Amount in
/// the values area.
fn xlsx_with_real_pivot() -> Vec<u8> {
    let parts: &[(&str, &str)] = &[
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
</Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
        ),
        (
            "xl/workbook.xml",
            r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sales" sheetId="1" r:id="rId1"/></sheets>
<pivotCaches><pivotCache cacheId="1" r:id="rId2"/></pivotCaches>
</workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>"#,
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#,
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>"#,
        ),
        (
            "xl/pivotCache/pivotCacheDefinition1.xml",
            r#"<?xml version="1.0"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" recordCount="5">
<cacheSource type="worksheet"><worksheetSource ref="A1:C6" sheet="Sales"/></cacheSource>
<cacheFields count="3">
<cacheField name="Region" numFmtId="0"><sharedItems/></cacheField>
<cacheField name="Product" numFmtId="0"><sharedItems/></cacheField>
<cacheField name="Amount" numFmtId="0"><sharedItems containsNumber="1"/></cacheField>
</cacheFields>
</pivotCacheDefinition>"#,
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
            r#"<?xml version="1.0"?><pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>"#,
        ),
        (
            "xl/pivotTables/pivotTable1.xml",
            r#"<?xml version="1.0"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="SalesPivot" cacheId="1" dataOnRows="0">
<location ref="A3:D10" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
<pivotFields count="3">
<pivotField axis="axisRow" showAll="0"/>
<pivotField axis="axisCol" showAll="0"/>
<pivotField dataField="1" showAll="0"/>
</pivotFields>
<rowFields count="1"><field x="0"/></rowFields>
<colFields count="1"><field x="1"/></colFields>
<dataFields count="1"><dataField name="Sum of Amount" fld="2" baseField="0" baseItem="0"/></dataFields>
</pivotTableDefinition>"#,
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

#[test]
fn reads_pivot_definition() {
    let wb = Workbook::load_from_bytes(&xlsx_with_real_pivot()).unwrap();
    let pivots = wb.pivot_tables();
    assert_eq!(pivots.len(), 1);
    let p = &pivots[0];

    assert_eq!(p.name, "SalesPivot");
    assert_eq!(p.cache_id, Some(1));
    assert_eq!(p.location.as_deref(), Some("A3:D10"));
    assert_eq!(p.source_sheet.as_deref(), Some("Sales"));
    assert_eq!(p.source_ref.as_deref(), Some("A1:C6"));
    assert_eq!(p.cache_fields, vec!["Region", "Product", "Amount"]);
    assert_eq!(p.row_fields, vec!["Region"]);
    assert_eq!(p.col_fields, vec!["Product"]);
    assert!(p.page_fields.is_empty());
    assert_eq!(
        p.data_fields,
        vec![PivotDataField {
            name: "Sum of Amount".to_string(),
            source_field: "Amount".to_string(),
            subtotal: "sum".to_string(),
        }]
    );
}

#[test]
fn survives_round_trip_and_still_reads() {
    // Save the loaded file, reload it, and confirm the model still parses -- the
    // preserved parts must remain parseable, not just present.
    let saved = Workbook::load_from_bytes(&xlsx_with_real_pivot())
        .unwrap()
        .save_to_bytes()
        .unwrap();
    let wb = Workbook::load_from_bytes(&saved).unwrap();
    let pivots = wb.pivot_tables();
    assert_eq!(pivots.len(), 1);
    assert_eq!(pivots[0].source_ref.as_deref(), Some("A1:C6"));
    assert_eq!(pivots[0].row_fields, vec!["Region"]);
}

#[test]
fn no_pivots_is_empty() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("S".to_string())).unwrap();
    assert!(wb.pivot_tables().is_empty());
}

#![no_main]

//! Fuzz target for workbook loading in rustypyxl-core.
//!
//! Drives the real `Workbook::load_from_bytes` entry point so that panics in
//! the library's own parsing (ZIP handling, XML parsing, shared-string and
//! style resolution, dimension hints) are caught — not just panics in the
//! zip/quick-xml dependencies.

use libfuzzer_sys::fuzz_target;
use std::io::{Cursor, Write};

/// Wrap small inputs as the workbook.xml of a minimal ZIP package so the
/// fuzzer can reach the XML parsing layers without first having to invent a
/// valid ZIP container byte-by-byte.
fn fuzz_as_workbook_xml(data: &[u8]) {
    if data.len() >= 1000 {
        return;
    }
    let mut zip_buffer = Vec::new();
    {
        let cursor = Cursor::new(&mut zip_buffer);
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Stored);

        if zip.start_file("xl/workbook.xml", options).is_ok() {
            let _ = zip.write_all(data);
        }
        if zip.start_file("[Content_Types].xml", options).is_ok() {
            let content_types = br#"<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="xml" ContentType="application/xml"/>
</Types>"#;
            let _ = zip.write_all(content_types);
        }
        if zip.start_file("_rels/.rels", options).is_ok() {
            let rels = br#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;
            let _ = zip.write_all(rels);
        }
        let _ = zip.finish();
    }
    if !zip_buffer.is_empty() {
        let _ = rustypyxl::Workbook::load_from_bytes(&zip_buffer);
    }
}

/// Same wrapping, but with the fuzz input as a worksheet so the cell/row
/// parsing paths get coverage too.
fn fuzz_as_sheet_xml(data: &[u8]) {
    if data.len() >= 1000 {
        return;
    }
    let mut zip_buffer = Vec::new();
    {
        let cursor = Cursor::new(&mut zip_buffer);
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Stored);

        if zip.start_file("xl/workbook.xml", options).is_ok() {
            let workbook = br#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
            let _ = zip.write_all(workbook);
        }
        if zip.start_file("xl/_rels/workbook.xml.rels", options).is_ok() {
            let rels = br#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;
            let _ = zip.write_all(rels);
        }
        if zip.start_file("xl/worksheets/sheet1.xml", options).is_ok() {
            let _ = zip.write_all(data);
        }
        let _ = zip.finish();
    }
    if !zip_buffer.is_empty() {
        let _ = rustypyxl::Workbook::load_from_bytes(&zip_buffer);
    }
}

fuzz_target!(|data: &[u8]| {
    // Limit input size to keep iterations fast
    if data.len() > 1024 * 1024 {
        return;
    }

    // Raw bytes through the real loader (covers ZIP container handling)
    let _ = rustypyxl::Workbook::load_from_bytes(data);

    // Structured variants that reach the XML layers
    fuzz_as_workbook_xml(data);
    fuzz_as_sheet_xml(data);
});

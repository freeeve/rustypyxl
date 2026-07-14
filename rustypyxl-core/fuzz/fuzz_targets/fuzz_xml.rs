#![no_main]

//! Fuzz target for the worksheet, shared-string, and style XML parsers.
//!
//! This target used to re-implement quick-xml event loops over the fuzz input,
//! which only ever exercised quick-xml -- a bug in rustypyxl's own parsing
//! could not be found by it. Instead, wrap the input as each of the XML parts
//! of a minimal but valid ZIP package and drive `Workbook::load_from_bytes`, so
//! the fuzzer reaches the real parsers with arbitrary bytes.
//!
//! `fuzz_load` covers xl/workbook.xml; this covers the parts underneath it,
//! where the parsing is far more involved: implied cell positions, inline rich
//! text, shared-string indices, and style/xf resolution.

use libfuzzer_sys::fuzz_target;
use std::io::{Cursor, Write};

const WORKBOOK_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#;

const WORKBOOK_RELS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

const CONTENT_TYPES: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"#;

const DEFAULT_SHEET: &[u8] =
    br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>
</worksheet>"#;

/// The parts of a minimal package. The one under test is replaced by the fuzz
/// input; the rest stay well-formed, so the loader gets far enough to actually
/// reach the parser being fuzzed.
const FUZZED_PARTS: [&str; 3] = [
    "xl/worksheets/sheet1.xml",
    "xl/sharedStrings.xml",
    "xl/styles.xml",
];

fn package_with(part: &str, data: &[u8]) -> Option<Vec<u8>> {
    let mut buffer = Vec::new();
    {
        let mut zip = zip::ZipWriter::new(Cursor::new(&mut buffer));
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Stored);

        let mut write = |name: &str, body: &[u8]| -> Option<()> {
            zip.start_file(name, options).ok()?;
            zip.write_all(body).ok()?;
            Some(())
        };

        write("[Content_Types].xml", CONTENT_TYPES)?;
        write("xl/workbook.xml", WORKBOOK_XML)?;
        write("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)?;
        write(
            "xl/worksheets/sheet1.xml",
            if part == "xl/worksheets/sheet1.xml" {
                data
            } else {
                DEFAULT_SHEET
            },
        )?;

        // sharedStrings.xml and styles.xml are optional parts; include only the
        // one being fuzzed, so each run exercises exactly one parser with the
        // arbitrary bytes.
        if part == "xl/sharedStrings.xml" || part == "xl/styles.xml" {
            write(part, data)?;
        }

        zip.finish().ok()?;
    }
    Some(buffer)
}

fuzz_target!(|data: &[u8]| {
    // Keep inputs small: the point is deep coverage of the parsers, not of the
    // ZIP layer, and huge inputs only slow the fuzzer down.
    if data.len() > 4096 {
        return;
    }

    for part in FUZZED_PARTS {
        if let Some(package) = package_with(part, data) {
            // A parse error is fine; a panic is not. Anything that does load
            // must also survive being written back out.
            if let Ok(workbook) = rustypyxl::Workbook::load_from_bytes(&package) {
                let _ = workbook.save_to_bytes();
            }
        }
    }
});

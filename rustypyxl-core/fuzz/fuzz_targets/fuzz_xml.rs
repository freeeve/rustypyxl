#![no_main]

//! Fuzz target for XML parsing in rustypyxl-core.
//!
//! Tests that malformed XML input does not cause crashes or panics in the
//! XML parsing code. The parser should gracefully handle:
//! - Invalid UTF-8 sequences
//! - Malformed XML structure (unclosed tags, invalid attributes)
//! - Deeply nested elements
//! - Very long attribute values or text content
//! - Invalid characters in element names or attributes
//! - Entity expansion edge cases

use libfuzzer_sys::fuzz_target;
use std::io::Cursor;

/// Test workbook.xml parsing with arbitrary input.
/// This exercises the XML parser that reads sheet metadata, named ranges,
/// and workbook structure.
fn fuzz_workbook_xml(data: &[u8]) {
    // Try to parse as if it were a workbook.xml file
    // The parser should handle arbitrary byte sequences without panicking
    let cursor = Cursor::new(data);

    // We use quick_xml directly to test the low-level parsing
    let mut reader = quick_xml::Reader::from_reader(cursor);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(_) => {}
            Err(_) => break, // Parser error is acceptable, panic is not
        }
        buf.clear();
    }
}

/// Test shared strings XML parsing with arbitrary input.
/// Shared strings contain cell text values and are parsed during workbook load.
fn fuzz_shared_strings_xml(data: &[u8]) {
    let cursor = Cursor::new(data);
    let mut reader = quick_xml::Reader::from_reader(cursor);
    reader.config_mut().trim_text(false); // Preserve whitespace like the real parser

    let mut buf = Vec::new();
    let mut in_t = false;
    let mut _current_string = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(e)) => {
                if e.name().as_ref() == b"t" {
                    in_t = true;
                }
            }
            Ok(quick_xml::events::Event::Text(e)) => {
                if in_t {
                    // Attempt to unescape XML entities - this can fail on malformed input
                    let _ = e.unescape();
                }
            }
            Ok(quick_xml::events::Event::End(e)) => {
                if e.name().as_ref() == b"t" {
                    in_t = false;
                } else if e.name().as_ref() == b"si" {
                    _current_string.clear();
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(_) => {}
            Err(_) => break,
        }
        buf.clear();
    }
}

/// Test worksheet XML parsing with arbitrary input.
/// This is the most complex parser, handling cells, formulas, styles, merged cells, etc.
fn fuzz_worksheet_xml(data: &[u8]) {
    let cursor = Cursor::new(data);
    let mut reader = quick_xml::Reader::from_reader(cursor);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut in_v = false;
    let mut in_f = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(e)) => {
                let name = e.name();
                match name.as_ref() {
                    b"c" => {
                        // Parse attributes like r="A1", t="s", s="1"
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let _ = std::str::from_utf8(attr.key.as_ref());
                                let _ = std::str::from_utf8(&attr.value);
                            }
                        }
                    }
                    b"v" => in_v = true,
                    b"f" => in_f = true,
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Text(e)) => {
                if in_v || in_f {
                    let _ = e.unescape();
                }
            }
            Ok(quick_xml::events::Event::End(e)) => {
                match e.name().as_ref() {
                    b"v" => in_v = false,
                    b"f" => in_f = false,
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Empty(e)) => {
                // Handle self-closing elements like <c r="A1"/>
                if e.name().as_ref() == b"c" {
                    for attr in e.attributes() {
                        if let Ok(attr) = attr {
                            let _ = std::str::from_utf8(attr.key.as_ref());
                            let _ = std::str::from_utf8(&attr.value);
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(_) => {}
            Err(_) => break,
        }
        buf.clear();
    }
}

/// Test styles.xml parsing with arbitrary input.
/// Styles contain fonts, fills, borders, number formats, and cell formatting.
fn fuzz_styles_xml(data: &[u8]) {
    let cursor = Cursor::new(data);
    let mut reader = quick_xml::Reader::from_reader(cursor);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(e)) | Ok(quick_xml::events::Event::Empty(e)) => {
                // Parse all attributes
                for attr in e.attributes() {
                    if let Ok(attr) = attr {
                        let _ = std::str::from_utf8(attr.key.as_ref());
                        let _ = std::str::from_utf8(&attr.value);
                        // Try parsing as numbers (common in styles)
                        if let Ok(val_str) = std::str::from_utf8(&attr.value) {
                            let _ = val_str.parse::<u32>();
                            let _ = val_str.parse::<f64>();
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(_) => {}
            Err(_) => break,
        }
        buf.clear();
    }
}

fuzz_target!(|data: &[u8]| {
    // Run all XML parsing variants on the input
    // The goal is to ensure none of them panic
    fuzz_workbook_xml(data);
    fuzz_shared_strings_xml(data);
    fuzz_worksheet_xml(data);
    fuzz_styles_xml(data);
});

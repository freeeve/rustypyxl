use crate::cell::InternedString;
use crate::worksheet::{cell_key, decode_cell_key, SheetVisibility, Worksheet, CellData};
use crate::cell::CellValue;
use crate::utils::column_to_letter;
use crate::error::Result;
use crate::autofilter::FilterType;
use crate::conditional::{ConditionalColor, ConditionalFormat, ConditionalFormatType};
use crate::pagesetup::Orientation;
use crate::style::StyleRegistry;
use zip::write::{FileOptions, ExtendedFileOptions};
use zip::ZipWriter;
use quick_xml::Writer;
use quick_xml::events::{BytesStart, BytesEnd, BytesText, Event};
use std::io::{Write, Cursor, Seek};
use std::collections::HashMap;
use rayon::prelude::*;

/// Returns true for C0 control characters that are illegal in XML 1.0
/// (everything below 0x20 except tab, line feed, and carriage return).
#[inline]
fn is_illegal_xml_char(b: u8) -> bool {
    b < 0x20 && !matches!(b, b'\t' | b'\n' | b'\r')
}

/// Strip control characters that are illegal in XML 1.0 without escaping.
/// Used before handing text to quick-xml, which performs entity escaping itself.
#[inline]
pub(crate) fn strip_illegal_xml_chars(s: &str) -> std::borrow::Cow<'_, str> {
    if s.bytes().any(is_illegal_xml_char) {
        std::borrow::Cow::Owned(
            s.chars()
                .filter(|&c| (c as u32) >= 0x20 || !is_illegal_xml_char(c as u8))
                .collect(),
        )
    } else {
        std::borrow::Cow::Borrowed(s)
    }
}

/// Compute the legacy 16-bit Excel sheet-protection password verifier
/// (CreatePasswordVerifier_Method1 from MS-XLS 2.2.9), as stored in the
/// `password` attribute of `sheetProtection`.
pub(crate) fn legacy_password_hash(password: &str) -> u16 {
    let bytes = password.as_bytes();
    let mut verifier: u16 = 0;
    for &b in bytes.iter().rev().chain(std::iter::once(&(bytes.len() as u8))) {
        verifier = ((verifier >> 14) & 0x0001) | ((verifier << 1) & 0x7fff);
        verifier ^= b as u16;
    }
    verifier ^ 0xCE4B
}

/// Escape XML special characters in text content.
/// Control characters that are illegal in XML 1.0 are stripped, since
/// emitting them produces files Excel refuses to open.
#[inline]
pub fn escape_xml(s: &str) -> std::borrow::Cow<'_, str> {
    if s.bytes()
        .any(|b| matches!(b, b'<' | b'>' | b'&' | b'"' | b'\'') || is_illegal_xml_char(b))
    {
        let mut escaped = String::with_capacity(s.len() + 8);
        for c in s.chars() {
            match c {
                '<' => escaped.push_str("&lt;"),
                '>' => escaped.push_str("&gt;"),
                '&' => escaped.push_str("&amp;"),
                '"' => escaped.push_str("&quot;"),
                '\'' => escaped.push_str("&apos;"),
                c if (c as u32) < 0x20 && is_illegal_xml_char(c as u8) => {}
                _ => escaped.push(c),
            }
        }
        std::borrow::Cow::Owned(escaped)
    } else {
        std::borrow::Cow::Borrowed(s)
    }
}

/// Format a cell value directly to a string buffer (for streaming writes).
/// Uses inline strings instead of shared strings for simplicity.
#[inline]
pub fn format_cell_value(buf: &mut String, coord: &str, value: &CellValue) {
    match value {
        CellValue::String(s) => {
            let escaped = escape_xml(s.as_ref());
            buf.push_str("<c r=\"");
            buf.push_str(coord);
            buf.push_str("\" t=\"inlineStr\"><is><t>");
            buf.push_str(&escaped);
            buf.push_str("</t></is></c>");
        }
        CellValue::Number(n) => {
            if !n.is_finite() {
                // NaN/Infinity are not valid SpreadsheetML numbers; emit an error cell.
                buf.push_str("<c r=\"");
                buf.push_str(coord);
                buf.push_str("\" t=\"e\"><v>#NUM!</v></c>");
                return;
            }
            buf.push_str("<c r=\"");
            buf.push_str(coord);
            buf.push_str("\"><v>");
            buf.push_str(ryu::Buffer::new().format(*n));
            buf.push_str("</v></c>");
        }
        CellValue::Boolean(b) => {
            buf.push_str("<c r=\"");
            buf.push_str(coord);
            buf.push_str("\" t=\"b\"><v>");
            buf.push_str(if *b { "1" } else { "0" });
            buf.push_str("</v></c>");
        }
        CellValue::Formula(f) => {
            let escaped = escape_xml(f);
            buf.push_str("<c r=\"");
            buf.push_str(coord);
            buf.push_str("\"><f>");
            buf.push_str(&escaped);
            buf.push_str("</f></c>");
        }
        CellValue::Date(d) => {
            let escaped = escape_xml(d);
            buf.push_str("<c r=\"");
            buf.push_str(coord);
            buf.push_str("\" t=\"d\"><v>");
            buf.push_str(&escaped);
            buf.push_str("</v></c>");
        }
        CellValue::Empty => {
            // Skip empty cells in streaming mode
        }
    }
}

/// Write cell data directly to a string buffer (fast path, no quick_xml overhead).
/// Uses itoa/ryu for fast number formatting.
#[inline]
fn write_cell_direct(
    buf: &mut String,
    coord: &str,
    cell_data: &CellData,
    style_index: Option<u32>,
    shared_string_map: &HashMap<InternedString, usize>,
) {
    // Helper to write style attribute
    let style_attr = style_index.map(|s| {
        let mut attr = String::with_capacity(10);
        attr.push_str(" s=\"");
        attr.push_str(itoa::Buffer::new().format(s));
        attr.push('"');
        attr
    });
    let style_str = style_attr.as_deref().unwrap_or("");

    match &cell_data.value {
        CellValue::String(s) => {
            if let Some(&idx) = shared_string_map.get(s) {
                // Shared string reference - use itoa for fast integer formatting
                buf.push_str("<c r=\"");
                buf.push_str(coord);
                buf.push('"');
                buf.push_str(style_str);
                buf.push_str(" t=\"s\"><v>");
                buf.push_str(itoa::Buffer::new().format(idx));
                buf.push_str("</v></c>");
            } else {
                // Inline string
                let escaped = escape_xml(s.as_ref());
                buf.push_str("<c r=\"");
                buf.push_str(coord);
                buf.push('"');
                buf.push_str(style_str);
                buf.push_str(" t=\"inlineStr\"><is><t>");
                buf.push_str(&escaped);
                buf.push_str("</t></is></c>");
            }
        }
        CellValue::Number(n) => {
            if !n.is_finite() {
                // NaN/Infinity are not valid SpreadsheetML numbers; emit an error cell.
                buf.push_str("<c r=\"");
                buf.push_str(coord);
                buf.push('"');
                buf.push_str(style_str);
                buf.push_str(" t=\"e\"><v>#NUM!</v></c>");
                return;
            }
            // Use ryu for fast float formatting
            buf.push_str("<c r=\"");
            buf.push_str(coord);
            buf.push('"');
            buf.push_str(style_str);
            buf.push_str("><v>");
            buf.push_str(ryu::Buffer::new().format(*n));
            buf.push_str("</v></c>");
        }
        CellValue::Boolean(b) => {
            buf.push_str("<c r=\"");
            buf.push_str(coord);
            buf.push('"');
            buf.push_str(style_str);
            buf.push_str(" t=\"b\"><v>");
            buf.push_str(if *b { "1" } else { "0" });
            buf.push_str("</v></c>");
        }
        CellValue::Formula(f) => {
            let escaped = escape_xml(f);
            buf.push_str("<c r=\"");
            buf.push_str(coord);
            buf.push('"');
            buf.push_str(style_str);
            // The cached result's type rides on the t attribute (numeric when absent)
            if cell_data.cached_formula_value.is_some() {
                if let Some(t) = cell_data.data_type.as_deref() {
                    if matches!(t, "str" | "b" | "e") {
                        buf.push_str(" t=\"");
                        buf.push_str(t);
                        buf.push('"');
                    }
                }
            }
            buf.push_str("><f>");
            buf.push_str(&escaped);
            buf.push_str("</f>");
            if let Some(ref cached) = cell_data.cached_formula_value {
                buf.push_str("<v>");
                buf.push_str(&escape_xml(cached));
                buf.push_str("</v>");
            }
            buf.push_str("</c>");
        }
        CellValue::Date(d) => {
            let escaped = escape_xml(d);
            buf.push_str("<c r=\"");
            buf.push_str(coord);
            buf.push('"');
            buf.push_str(style_str);
            buf.push_str(" t=\"d\"><v>");
            buf.push_str(&escaped);
            buf.push_str("</v></c>");
        }
        CellValue::Empty => {
            // Skip empty cells without styles, but include if there's a style
            if style_str.is_empty() {
                return; // Skip completely empty cells
            }
            buf.push_str("<c r=\"");
            buf.push_str(coord);
            buf.push('"');
            buf.push_str(style_str);
            buf.push_str("/>");
        }
    }
}

pub fn write_content_types<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &FileOptions<'static, ExtendedFileOptions>,
    sheet_count: usize,
    has_shared_strings: bool,
    comment_sheet_ids: &[u32],
    table_count: usize,
) -> Result<()> {
    zip.start_file("[Content_Types].xml", options.clone())?;

    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut types_start = BytesStart::new("Types");
    types_start.push_attribute(("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types"));
    writer.write_event(quick_xml::events::Event::Start(types_start))?;

    // Default overrides
    let mut default1 = BytesStart::new("Default");
    default1.push_attribute(("Extension", "rels"));
    default1.push_attribute(("ContentType", "application/vnd.openxmlformats-package.relationships+xml"));
    writer.write_event(quick_xml::events::Event::Empty(default1))?;

    let mut default2 = BytesStart::new("Default");
    default2.push_attribute(("Extension", "xml"));
    default2.push_attribute(("ContentType", "application/xml"));
    writer.write_event(quick_xml::events::Event::Empty(default2))?;

    if !comment_sheet_ids.is_empty() {
        // Comment boxes are anchored by legacy VML drawings
        let mut default3 = BytesStart::new("Default");
        default3.push_attribute(("Extension", "vml"));
        default3.push_attribute(("ContentType", "application/vnd.openxmlformats-officedocument.vmlDrawing"));
        writer.write_event(quick_xml::events::Event::Empty(default3))?;
    }

    // Overrides
    let mut override1 = BytesStart::new("Override");
    override1.push_attribute(("PartName", "/xl/workbook.xml"));
    override1.push_attribute(("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"));
    writer.write_event(quick_xml::events::Event::Empty(override1))?;

    for i in 1..=sheet_count {
        let part_name = format!("/xl/worksheets/sheet{}.xml", i);
        let mut override_elem = BytesStart::new("Override");
        override_elem.push_attribute(("PartName", part_name.as_str()));
        override_elem.push_attribute(("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
        writer.write_event(quick_xml::events::Event::Empty(override_elem))?;
    }

    // Only include sharedStrings if there are strings
    if has_shared_strings {
        let mut override2 = BytesStart::new("Override");
        override2.push_attribute(("PartName", "/xl/sharedStrings.xml"));
        override2.push_attribute(("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"));
        writer.write_event(quick_xml::events::Event::Empty(override2))?;
    }

    let mut override3 = BytesStart::new("Override");
    override3.push_attribute(("PartName", "/xl/styles.xml"));
    override3.push_attribute(("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"));
    writer.write_event(quick_xml::events::Event::Empty(override3))?;

    for sheet_id in comment_sheet_ids {
        let part_name = format!("/xl/comments/comment{}.xml", sheet_id);
        let mut override_elem = BytesStart::new("Override");
        override_elem.push_attribute(("PartName", part_name.as_str()));
        override_elem.push_attribute(("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"));
        writer.write_event(quick_xml::events::Event::Empty(override_elem))?;
    }

    for table_id in 1..=table_count {
        let part_name = format!("/xl/tables/table{}.xml", table_id);
        let mut override_elem = BytesStart::new("Override");
        override_elem.push_attribute(("PartName", part_name.as_str()));
        override_elem.push_attribute(("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"));
        writer.write_event(quick_xml::events::Event::Empty(override_elem))?;
    }

    writer.write_event(quick_xml::events::Event::End(BytesEnd::new("Types")))?;

    let result = writer.into_inner().into_inner();
    zip.write_all(&result)?;
    Ok(())
}

pub fn write_rels<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &FileOptions<'static, ExtendedFileOptions>,
) -> Result<()> {
    zip.start_file("_rels/.rels", options.clone())?;
    
    let content = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"#;
    
    zip.write_all(content.as_bytes())?;
    Ok(())
}

pub fn write_doc_props<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &FileOptions<'static, ExtendedFileOptions>,
) -> Result<()> {
    // Write docProps/core.xml
    zip.start_file("docProps/core.xml", options.clone())?;
    let core_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
</cp:coreProperties>"#;
    zip.write_all(core_xml.as_bytes())?;
    
    // Write docProps/app.xml
    zip.start_file("docProps/app.xml", options.clone())?;
    let app_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>RustyPyXL</Application>
</Properties>"#;
    zip.write_all(app_xml.as_bytes())?;
    
    Ok(())
}

pub fn write_workbook_xml<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &FileOptions<'static, ExtendedFileOptions>,
    sheets: &[(String, SheetVisibility)],
    named_ranges: &[crate::workbook::NamedRange],
    active_tab: usize,
) -> Result<()> {
    zip.start_file("xl/workbook.xml", options.clone())?;
    
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut workbook_start = BytesStart::new("workbook");
    workbook_start.push_attribute(("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    workbook_start.push_attribute(("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"));
    writer.write_event(quick_xml::events::Event::Start(workbook_start))?;
    
    // workbookPr
    writer.write_event(quick_xml::events::Event::Empty(BytesStart::new("workbookPr")))?;
    
    // bookViews
    writer.write_event(quick_xml::events::Event::Start(BytesStart::new("bookViews")))?;
    let mut view = BytesStart::new("workbookView");
    view.push_attribute(("visibility", "visible"));
    view.push_attribute(("minimized", "0"));
    view.push_attribute(("showHorizontalScroll", "1"));
    view.push_attribute(("showVerticalScroll", "1"));
    view.push_attribute(("showSheetTabs", "1"));
    view.push_attribute(("tabRatio", "600"));
    view.push_attribute(("firstSheet", "0"));
    let active_tab = active_tab.min(sheets.len().saturating_sub(1));
    view.push_attribute(("activeTab", active_tab.to_string().as_str()));
    writer.write_event(quick_xml::events::Event::Empty(view))?;
    writer.write_event(quick_xml::events::Event::End(BytesEnd::new("bookViews")))?;
    
    // sheets
    writer.write_event(quick_xml::events::Event::Start(BytesStart::new("sheets")))?;
    for (idx, (name, visibility)) in sheets.iter().enumerate() {
        let sheet_id = (idx + 1) as u32;
        let r_id = format!("rId{}", idx + 1);
        let mut sheet = BytesStart::new("sheet");
        sheet.push_attribute(("name", name.as_str()));
        sheet.push_attribute(("sheetId", sheet_id.to_string().as_str()));
        sheet.push_attribute(("state", visibility.as_str()));
        sheet.push_attribute(("r:id", r_id.as_str()));
        writer.write_event(quick_xml::events::Event::Empty(sheet))?;
    }
    writer.write_event(quick_xml::events::Event::End(BytesEnd::new("sheets")))?;
    
    // definedNames (named ranges), preserving sheet scope and visibility
    if !named_ranges.is_empty() {
        writer.write_event(quick_xml::events::Event::Start(BytesStart::new("definedNames")))?;
        for nr in named_ranges {
            let mut defined_name = BytesStart::new("definedName");
            defined_name.push_attribute(("name", nr.name.as_str()));
            if let Some(sheet_id) = nr.local_sheet_id {
                defined_name.push_attribute(("localSheetId", sheet_id.to_string().as_str()));
            }
            if nr.hidden {
                defined_name.push_attribute(("hidden", "1"));
            }
            writer.write_event(quick_xml::events::Event::Start(defined_name))?;
            writer.write_event(quick_xml::events::Event::Text(BytesText::new(&nr.range)))?;
            writer.write_event(quick_xml::events::Event::End(BytesEnd::new("definedName")))?;
        }
        writer.write_event(quick_xml::events::Event::End(BytesEnd::new("definedNames")))?;
    }
    
    writer.write_event(quick_xml::events::Event::End(BytesEnd::new("workbook")))?;
    
    let result = writer.into_inner().into_inner();
    zip.write_all(&result)?;
    Ok(())
}

pub fn write_workbook_rels<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &FileOptions<'static, ExtendedFileOptions>,
    sheet_count: usize,
    has_shared_strings: bool,
) -> Result<()> {
    zip.start_file("xl/_rels/workbook.xml.rels", options.clone())?;

    let mut content = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
"#);

    for i in 1..=sheet_count {
        content.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml"/>
"#,
            i, i
        ));
    }

    // Only include sharedStrings if there are strings
    if has_shared_strings {
        content.push_str(r#"<Relationship Id="rIdSharedStrings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
"#);
    }

    content.push_str(r#"<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#);

    zip.write_all(content.as_bytes())?;
    Ok(())
}

/// Returns (ordered list of strings, map from string -> index for O(1) lookup)
pub fn collect_shared_strings(
    worksheets: &[Worksheet],
) -> (Vec<InternedString>, HashMap<InternedString, usize>) {
    // Estimate capacity: count string cells across all worksheets
    let estimated_strings: usize = worksheets
        .iter()
        .map(|ws| ws.cells.values().filter(|c| matches!(c.value, CellValue::String(_))).count())
        .sum();

    let mut strings = Vec::with_capacity(estimated_strings);
    let mut string_map = HashMap::with_capacity(estimated_strings);

    for worksheet in worksheets {
        for cell_data in worksheet.cells.values() {
            if let CellValue::String(s) = &cell_data.value {
                if !string_map.contains_key(s) {
                    string_map.insert(s.clone(), strings.len());
                    strings.push(s.clone());
                }
            }
        }
    }

    (strings, string_map)
}

pub fn write_shared_strings<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &FileOptions<'static, ExtendedFileOptions>,
    strings: &[InternedString],
) -> Result<()> {
    zip.start_file("xl/sharedStrings.xml", options.clone())?;

    // Pre-allocate buffer: ~50 bytes per string for XML overhead
    let estimated_size = strings.len() * 50 + 200;
    let mut writer = Writer::new(Cursor::new(Vec::with_capacity(estimated_size)));
    let mut sst = BytesStart::new("sst");
    sst.push_attribute(("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    let count_str = strings.len().to_string();
    sst.push_attribute(("count", count_str.as_str()));
    sst.push_attribute(("uniqueCount", count_str.as_str()));
    writer.write_event(quick_xml::events::Event::Start(sst))?;
    
    for s in strings {
        writer.write_event(quick_xml::events::Event::Start(BytesStart::new("si")))?;
        writer.write_event(quick_xml::events::Event::Start(BytesStart::new("t")))?;
        writer.write_event(quick_xml::events::Event::Text(BytesText::new(
            &strip_illegal_xml_chars(s.as_ref()),
        )))?;
        writer.write_event(quick_xml::events::Event::End(BytesEnd::new("t")))?;
        writer.write_event(quick_xml::events::Event::End(BytesEnd::new("si")))?;
    }
    
    writer.write_event(quick_xml::events::Event::End(BytesEnd::new("sst")))?;
    
    let result = writer.into_inner().into_inner();
    zip.write_all(&result)?;
    Ok(())
}

/// Write a color element to the XML string, handling theme and RGB colors.
///
/// Colors are stored internally as:
/// - `"theme:N"` for theme colors
/// - `"#AARRGGBB"` or `"AARRGGBB"` (8-char aRGB from XML roundtrip)
/// - `"#RRGGBB"` or `"RRGGBB"` (6-char RGB, needs FF alpha prefix)
fn write_color_attr(xml: &mut String, element: &str, color: &str) {
    if let Some(theme) = color.strip_prefix("theme:") {
        xml.push_str(&format!(r#"<{} theme="{}"/>"#, element, theme));
    } else {
        let hex = color.strip_prefix('#').unwrap_or(color);
        if hex.len() >= 8 {
            xml.push_str(&format!(r#"<{} rgb="{}"/>"#, element, hex));
        } else {
            xml.push_str(&format!(r#"<{} rgb="FF{}"/>"#, element, hex));
        }
    }
}

/// Write a single font element to the XML string.
fn write_font_xml(xml: &mut String, font: &crate::style::Font) {
    xml.push_str("<font>");
    if font.bold {
        xml.push_str("<b/>");
    }
    if font.italic {
        xml.push_str("<i/>");
    }
    if let Some(ref u) = font.underline {
        if u == "single" {
            xml.push_str("<u/>");
        } else {
            xml.push_str(&format!(r#"<u val="{}"/>"#, u));
        }
    }
    if font.strike {
        xml.push_str("<strike/>");
    }
    if let Some(ref va) = font.vert_align {
        xml.push_str(&format!(r#"<vertAlign val="{}"/>"#, va));
    }
    if let Some(size) = font.size {
        xml.push_str(&format!(r#"<sz val="{}"/>"#, size));
    }
    if let Some(ref color) = font.color {
        write_color_attr(xml, "color", color);
    } else {
        xml.push_str(r#"<color theme="1"/>"#);
    }
    if let Some(ref name) = font.name {
        xml.push_str(&format!(r#"<name val="{}"/>"#, escape_xml(name)));
    } else {
        xml.push_str(r#"<name val="Calibri"/>"#);
    }
    xml.push_str(r#"<family val="2"/>"#);
    xml.push_str("</font>");
}

/// Write a single fill element to the XML string.
fn write_fill_xml(xml: &mut String, fill: &crate::style::Fill) {
    xml.push_str("<fill>");
    if let Some(ref pattern) = fill.pattern_type {
        xml.push_str(&format!(r#"<patternFill patternType="{}">"#, pattern));
        if let Some(ref fg) = fill.fg_color {
            write_color_attr(xml, "fgColor", fg);
        }
        if let Some(ref bg) = fill.bg_color {
            write_color_attr(xml, "bgColor", bg);
        }
        xml.push_str("</patternFill>");
    } else {
        xml.push_str("<patternFill/>");
    }
    xml.push_str("</fill>");
}

/// Write alignment element to the XML string.
fn write_alignment_xml(xml: &mut String, align: &crate::style::Alignment) {
    xml.push_str("<alignment");
    if let Some(ref h) = align.horizontal {
        xml.push_str(&format!(r#" horizontal="{}""#, h));
    }
    if let Some(ref v) = align.vertical {
        xml.push_str(&format!(r#" vertical="{}""#, v));
    }
    if align.wrap_text {
        xml.push_str(r#" wrapText="1""#);
    }
    if let Some(rot) = align.text_rotation {
        xml.push_str(&format!(r#" textRotation="{}""#, rot));
    }
    if align.shrink_to_fit {
        xml.push_str(r#" shrinkToFit="1""#);
    }
    xml.push_str("/>");
}

/// Write protection element to the XML string.
fn write_protection_xml(xml: &mut String, prot: &crate::style::Protection) {
    xml.push_str("<protection");
    xml.push_str(&format!(r#" locked="{}""#, if prot.locked { "1" } else { "0" }));
    if prot.hidden {
        xml.push_str(r#" hidden="1""#);
    }
    xml.push_str("/>");
}

/// Write a single cellXf element to the XML string.
fn write_cell_xf_xml(xml: &mut String, xf: &crate::style::CellXf) {
    xml.push_str(&format!(
        r#"<xf numFmtId="{}" fontId="{}" fillId="{}" borderId="{}""#,
        xf.num_fmt_id, xf.font_id, xf.fill_id, xf.border_id
    ));
    if xf.apply_font {
        xml.push_str(r#" applyFont="1""#);
    }
    if xf.apply_fill {
        xml.push_str(r#" applyFill="1""#);
    }
    if xf.apply_border {
        xml.push_str(r#" applyBorder="1""#);
    }
    if xf.apply_number_format {
        xml.push_str(r#" applyNumberFormat="1""#);
    }
    if xf.alignment.is_some() {
        xml.push_str(r#" applyAlignment="1""#);
    }
    if xf.protection.is_some() {
        xml.push_str(r#" applyProtection="1""#);
    }

    let has_children = xf.alignment.is_some() || xf.protection.is_some();
    if has_children {
        xml.push('>');
        if let Some(ref align) = xf.alignment {
            write_alignment_xml(xml, align);
        }
        if let Some(ref prot) = xf.protection {
            write_protection_xml(xml, prot);
        }
        xml.push_str("</xf>");
    } else {
        xml.push_str("/>");
    }
}

/// Collect the deduplicated differential formats (dxf entries) used by all
/// conditional-formatting rules, in deterministic order. The index of a
/// format in this list is its dxfId, shared between styles.xml and each
/// worksheet's cfRule elements.
pub fn collect_dxfs(worksheets: &[Worksheet]) -> Vec<ConditionalFormat> {
    let mut dxfs: Vec<ConditionalFormat> = Vec::new();
    for ws in worksheets {
        for cf in &ws.conditional_formatting {
            for rule in &cf.rules {
                if let Some(fmt) = &rule.format {
                    if !dxfs.contains(fmt) {
                        dxfs.push(fmt.clone());
                    }
                }
            }
        }
    }
    dxfs
}

/// Write one color element for a dxf child (font color / bgColor / border color).
fn write_dxf_color(xml: &mut String, element: &str, color: &ConditionalColor) {
    xml.push('<');
    xml.push_str(element);
    if let Some(ref rgb) = color.rgb {
        let hex = rgb.strip_prefix('#').unwrap_or(rgb);
        if hex.len() >= 8 {
            xml.push_str(&format!(r#" rgb="{}""#, escape_xml(hex)));
        } else {
            xml.push_str(&format!(r#" rgb="FF{}""#, escape_xml(hex)));
        }
    } else if let Some(theme) = color.theme {
        xml.push_str(&format!(r#" theme="{}""#, theme));
    }
    if let Some(tint) = color.tint {
        xml.push_str(&format!(r#" tint="{}""#, tint));
    }
    xml.push_str("/>");
}

/// Write the dxfs (differential formats) section referenced by cfRule dxfId.
/// CT_Dxf child order: font, numFmt, fill, alignment, border, protection.
fn write_dxfs_xml(xml: &mut String, dxfs: &[ConditionalFormat]) {
    if dxfs.is_empty() {
        xml.push_str(r#"<dxfs count="0"/>"#);
        return;
    }
    xml.push_str(&format!(r#"<dxfs count="{}">"#, dxfs.len()));
    for (idx, fmt) in dxfs.iter().enumerate() {
        xml.push_str("<dxf>");

        let has_font = fmt.font_color.is_some()
            || fmt.bold.is_some()
            || fmt.italic.is_some()
            || fmt.underline.is_some()
            || fmt.strikethrough.is_some();
        if has_font {
            xml.push_str("<font>");
            if let Some(b) = fmt.bold {
                xml.push_str(if b { "<b/>" } else { r#"<b val="0"/>"# });
            }
            if let Some(i) = fmt.italic {
                xml.push_str(if i { "<i/>" } else { r#"<i val="0"/>"# });
            }
            if let Some(st) = fmt.strikethrough {
                xml.push_str(if st { "<strike/>" } else { r#"<strike val="0"/>"# });
            }
            if let Some(u) = fmt.underline {
                xml.push_str(if u { "<u/>" } else { r#"<u val="none"/>"# });
            }
            if let Some(ref color) = fmt.font_color {
                write_dxf_color(xml, "color", color);
            }
            xml.push_str("</font>");
        }

        if let Some(ref code) = fmt.number_format {
            // dxf numFmt ids only need to be unique among dxfs; 200+ avoids
            // the builtin (0-163) and workbook custom (164+) ranges in use
            xml.push_str(&format!(
                r#"<numFmt numFmtId="{}" formatCode="{}"/>"#,
                200 + idx,
                escape_xml(code)
            ));
        }

        if let Some(ref fill) = fmt.fill_color {
            // In a dxf, the highlight color of a solid pattern goes in bgColor
            xml.push_str("<fill><patternFill>");
            write_dxf_color(xml, "bgColor", fill);
            xml.push_str("</patternFill></fill>");
        }

        if let Some(ref border) = fmt.border_color {
            xml.push_str("<border>");
            for side in ["left", "right", "top", "bottom"] {
                xml.push_str(&format!(r#"<{} style="thin">"#, side));
                write_dxf_color(xml, "color", border);
                xml.push_str(&format!("</{}>", side));
            }
            xml.push_str("</border>");
        }

        xml.push_str("</dxf>");
    }
    xml.push_str("</dxfs>");
}

pub fn write_styles_xml<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &FileOptions<'static, ExtendedFileOptions>,
    styles: &StyleRegistry,
    dxfs: &[ConditionalFormat],
) -> Result<()> {
    zip.start_file("xl/styles.xml", options.clone())?;

    let mut xml = String::with_capacity(4096);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#);

    // Number formats
    if !styles.num_fmts.is_empty() {
        xml.push_str(&format!(r#"<numFmts count="{}">"#, styles.num_fmts.len()));
        for (id, code) in &styles.num_fmts {
            xml.push_str(&format!(r#"<numFmt numFmtId="{}" formatCode="{}"/>"#, id, escape_xml(code)));
        }
        xml.push_str("</numFmts>");
    } else {
        xml.push_str(r#"<numFmts count="0"/>"#);
    }

    // Fonts
    xml.push_str(&format!(r#"<fonts count="{}">"#, styles.fonts.len()));
    for font in &styles.fonts {
        write_font_xml(&mut xml, font);
    }
    xml.push_str("</fonts>");

    // Fills
    xml.push_str(&format!(r#"<fills count="{}">"#, styles.fills.len()));
    for fill in &styles.fills {
        write_fill_xml(&mut xml, fill);
    }
    xml.push_str("</fills>");

    // Borders
    xml.push_str(&format!(r#"<borders count="{}">"#, styles.borders.len()));
    for border in &styles.borders {
        xml.push_str("<border>");
        write_border_side(&mut xml, "left", &border.left);
        write_border_side(&mut xml, "right", &border.right);
        write_border_side(&mut xml, "top", &border.top);
        write_border_side(&mut xml, "bottom", &border.bottom);
        write_border_side(&mut xml, "diagonal", &border.diagonal);
        xml.push_str("</border>");
    }
    xml.push_str("</borders>");

    // Cell style XFs (just one default)
    xml.push_str(r#"<cellStyleXfs count="1">"#);
    xml.push_str(r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>"#);
    xml.push_str("</cellStyleXfs>");

    // Cell XFs
    xml.push_str(&format!(r#"<cellXfs count="{}">"#, styles.cell_xfs.len()));
    for xf in &styles.cell_xfs {
        write_cell_xf_xml(&mut xml, xf);
    }
    xml.push_str("</cellXfs>");

    // Cell styles (just Normal)
    xml.push_str(r#"<cellStyles count="1">"#);
    xml.push_str(r#"<cellStyle name="Normal" xfId="0" builtinId="0"/>"#);
    xml.push_str("</cellStyles>");

    // Differential formats for conditional formatting (referenced by dxfId)
    write_dxfs_xml(&mut xml, dxfs);

    xml.push_str("</styleSheet>");

    zip.write_all(xml.as_bytes())?;
    Ok(())
}

/// Helper to write a border side element.
fn write_border_side(xml: &mut String, name: &str, side: &Option<crate::style::BorderStyle>) {
    if let Some(ref s) = side {
        xml.push_str(&format!(r#"<{} style="{}">"#, name, s.style));
        if let Some(ref color) = s.color {
            write_color_attr(xml, "color", color);
        }
        xml.push_str(&format!("</{}>", name));
    } else {
        xml.push_str(&format!("<{}/>", name));
    }
}

/// External (URL) hyperlinks in deterministic cell order. The position in
/// this list defines the relationship id (`rIdHL{i+1}`) shared between the
/// worksheet XML and its .rels part, so both must derive it from here.
pub fn collect_external_hyperlinks(worksheet: &Worksheet) -> Vec<((u32, u32), String)> {
    let mut links: Vec<((u32, u32), String)> = worksheet
        .cells
        .iter()
        .filter_map(|(key, cd)| {
            cd.hyperlink
                .as_ref()
                .filter(|url| !url.starts_with('#'))
                .map(|url| (decode_cell_key(*key), url.clone()))
        })
        .collect();
    links.sort_by_key(|(coord, _)| *coord);
    links
}

#[allow(clippy::too_many_arguments)]
pub fn write_worksheet_xml<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &FileOptions<'static, ExtendedFileOptions>,
    worksheet: &Worksheet,
    sheet_id: u32,
    shared_string_map: &HashMap<InternedString, usize>,
    table_rel_ids: &[String],
    dxfs: &[ConditionalFormat],
    has_comments: bool,
    style_overrides: &HashMap<u64, u32>,
) -> Result<()> {
    let path = format!("xl/worksheets/sheet{}.xml", sheet_id);
    zip.start_file(&path, options.clone())?;

    // Pre-allocate buffer based on estimated size (rough estimate: 100 bytes per cell)
    let estimated_size = worksheet.cells.len() * 100;
    let mut writer = Writer::new(Cursor::new(Vec::with_capacity(estimated_size)));
    let mut worksheet_start = BytesStart::new("worksheet");
    worksheet_start.push_attribute(("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    worksheet_start.push_attribute(("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"));
    writer.write_event(quick_xml::events::Event::Start(worksheet_start))?;
    
    // sheetPr
    writer.write_event(quick_xml::events::Event::Start(BytesStart::new("sheetPr")))?;
    let mut outline = BytesStart::new("outlinePr");
    outline.push_attribute(("summaryBelow", "1"));
    outline.push_attribute(("summaryRight", "1"));
    writer.write_event(quick_xml::events::Event::Empty(outline))?;
    writer.write_event(quick_xml::events::Event::Empty(BytesStart::new("pageSetUpPr")))?;
    writer.write_event(quick_xml::events::Event::End(BytesEnd::new("sheetPr")))?;
    
    // dimension (if we have cells)
    if worksheet.max_row > 0 && worksheet.max_column > 0 {
        let start = "A1";
        let end = format!("{}{}", column_to_letter(worksheet.max_column), worksheet.max_row);
        let mut dim = BytesStart::new("dimension");
        dim.push_attribute(("ref", format!("{}:{}", start, end).as_str()));
        writer.write_event(quick_xml::events::Event::Empty(dim))?;
    }
    
    // sheetViews
    writer.write_event(quick_xml::events::Event::Start(BytesStart::new("sheetViews")))?;
    let frozen = worksheet
        .freeze_panes
        .as_deref()
        .and_then(|cell| crate::utils::parse_coordinate(cell).ok().map(|(row, col)| (cell, row, col)))
        .filter(|&(_, row, col)| row > 1 || col > 1);
    if let Some((cell, row, col)) = frozen {
        let x_split = col - 1;
        let y_split = row - 1;
        let active_pane = if x_split > 0 && y_split > 0 {
            "bottomRight"
        } else if y_split > 0 {
            "bottomLeft"
        } else {
            "topRight"
        };
        let mut view = BytesStart::new("sheetView");
        view.push_attribute(("workbookViewId", "0"));
        writer.write_event(quick_xml::events::Event::Start(view))?;
        let mut pane = BytesStart::new("pane");
        if x_split > 0 {
            pane.push_attribute(("xSplit", x_split.to_string().as_str()));
        }
        if y_split > 0 {
            pane.push_attribute(("ySplit", y_split.to_string().as_str()));
        }
        pane.push_attribute(("topLeftCell", cell));
        pane.push_attribute(("activePane", active_pane));
        pane.push_attribute(("state", "frozen"));
        writer.write_event(quick_xml::events::Event::Empty(pane))?;
        let mut selection = BytesStart::new("selection");
        selection.push_attribute(("pane", active_pane));
        selection.push_attribute(("activeCell", cell));
        selection.push_attribute(("sqref", cell));
        writer.write_event(quick_xml::events::Event::Empty(selection))?;
        writer.write_event(quick_xml::events::Event::End(BytesEnd::new("sheetView")))?;
    } else {
        let mut view = BytesStart::new("sheetView");
        view.push_attribute(("workbookViewId", "0"));
        writer.write_event(quick_xml::events::Event::Empty(view))?;
    }
    writer.write_event(quick_xml::events::Event::End(BytesEnd::new("sheetViews")))?;
    
    // sheetFormatPr
    let mut format_pr = BytesStart::new("sheetFormatPr");
    format_pr.push_attribute(("baseColWidth", "8"));
    format_pr.push_attribute(("defaultRowHeight", "15"));
    writer.write_event(quick_xml::events::Event::Empty(format_pr))?;
    
    // cols (column dimensions)
    if !worksheet.column_dimensions.is_empty() {
        writer.write_event(quick_xml::events::Event::Start(BytesStart::new("cols")))?;
        for (&col, &width) in &worksheet.column_dimensions {
            let mut col_elem = BytesStart::new("col");
            col_elem.push_attribute(("min", col.to_string().as_str()));
            col_elem.push_attribute(("max", col.to_string().as_str()));
            col_elem.push_attribute(("width", width.to_string().as_str()));
            col_elem.push_attribute(("customWidth", "1"));
            writer.write_event(quick_xml::events::Event::Empty(col_elem))?;
        }
        writer.write_event(quick_xml::events::Event::End(BytesEnd::new("cols")))?;
    }
    
    // sheetData
    writer.write_event(quick_xml::events::Event::Start(BytesStart::new("sheetData")))?;

    // Group cells by row - pre-allocate based on max_row
    let estimated_rows = worksheet.max_row as usize;
    type RowCells<'a> = HashMap<u32, Vec<((u32, u32), &'a CellData)>>;
    let mut rows: RowCells = HashMap::with_capacity(estimated_rows);
    for (key, cell_data) in &worksheet.cells {
        let (row, col) = decode_cell_key(*key);
        rows.entry(row).or_default().push(((row, col), cell_data));
    }

    // Write rows in order
    let mut row_numbers: Vec<u32> = rows.keys().copied().collect();
    row_numbers.sort();

    // Use Rayon to generate XML for rows in parallel
    // Each row is processed independently, then results are concatenated in order
    let cell_buf: String = if row_numbers.len() > 1000 {
        // For large worksheets, use parallel processing
        // Process in chunks to balance parallelism overhead vs benefit
        const CHUNK_SIZE: usize = 5000;

        let chunks: Vec<_> = row_numbers.chunks(CHUNK_SIZE).collect();
        let chunk_results: Vec<String> = chunks
            .par_iter()
            .map(|chunk| {
                let mut buf = String::with_capacity(chunk.len() * 200);
                let mut itoa_buf = itoa::Buffer::new();
                let mut ryu_buf = ryu::Buffer::new();

                for &row_num in *chunk {
                    let cells = rows.get(&row_num).unwrap();

                    // Sort cells by column (need to clone since we're in parallel)
                    let mut sorted_cells: Vec<_> = cells.to_vec();
                    sorted_cells.sort_by_key(|((_, col), _)| *col);

                    // Write row start
                    if let Some(height) = worksheet.row_dimensions.get(&row_num) {
                        buf.push_str("<row r=\"");
                        buf.push_str(itoa_buf.format(row_num));
                        buf.push_str("\" ht=\"");
                        buf.push_str(ryu_buf.format(*height));
                        buf.push_str("\" customHeight=\"1\">");
                    } else {
                        buf.push_str("<row r=\"");
                        buf.push_str(itoa_buf.format(row_num));
                        buf.push_str("\">");
                    }

                    // Write cells
                    for &((row, col), cell_data) in &sorted_cells {
                        let coord = format!("{}{}", column_to_letter(col), row);
                        let style_index = cell_data
                            .style_index
                            .or_else(|| style_overrides.get(&cell_key(row, col)).copied());
                        write_cell_direct(&mut buf, &coord, cell_data, style_index, shared_string_map);
                    }

                    buf.push_str("</row>");
                }
                buf
            })
            .collect();

        // Concatenate all chunks in order
        let total_len: usize = chunk_results.iter().map(|s| s.len()).sum();
        let mut result = String::with_capacity(total_len);
        for chunk in chunk_results {
            result.push_str(&chunk);
        }
        result
    } else {
        // For small worksheets, use sequential processing (less overhead)
        let mut buf = String::with_capacity(worksheet.cells.len() * 40);
        let mut itoa_buf = itoa::Buffer::new();
        let mut ryu_buf = ryu::Buffer::new();

        for row_num in row_numbers {
            let cells = rows.get_mut(&row_num).unwrap();
            cells.sort_by_key(|((_, col), _)| *col);

            if let Some(height) = worksheet.row_dimensions.get(&row_num) {
                buf.push_str("<row r=\"");
                buf.push_str(itoa_buf.format(row_num));
                buf.push_str("\" ht=\"");
                buf.push_str(ryu_buf.format(*height));
                buf.push_str("\" customHeight=\"1\">");
            } else {
                buf.push_str("<row r=\"");
                buf.push_str(itoa_buf.format(row_num));
                buf.push_str("\">");
            }

            for &((row, col), cell_data) in cells.iter() {
                let coord = format!("{}{}", column_to_letter(col), row);
                let style_index = cell_data
                    .style_index
                    .or_else(|| style_overrides.get(&cell_key(row, col)).copied());
                write_cell_direct(&mut buf, &coord, cell_data, style_index, shared_string_map);
            }

            buf.push_str("</row>");
        }
        buf
    };

    // Write the cell buffer to the XML writer
    writer.get_mut().write_all(cell_buf.as_bytes())?;

    writer.write_event(Event::End(BytesEnd::new("sheetData")))?;

    // sheetProtection (per CT_Worksheet schema order: directly after sheetData)
    if let Some(ref protection) = worksheet.protection {
        if protection.sheet {
            let mut sheet_protection = BytesStart::new("sheetProtection");
            sheet_protection.push_attribute(("sheet", "1"));
            sheet_protection.push_attribute(("selectLockedCells", if protection.select_locked_cells { "1" } else { "0" }));
            sheet_protection.push_attribute(("selectUnlockedCells", if protection.select_unlocked_cells { "1" } else { "0" }));
            sheet_protection.push_attribute(("formatCells", if protection.format_cells { "1" } else { "0" }));
            sheet_protection.push_attribute(("formatColumns", if protection.format_columns { "1" } else { "0" }));
            sheet_protection.push_attribute(("formatRows", if protection.format_rows { "1" } else { "0" }));
            sheet_protection.push_attribute(("insertColumns", if protection.insert_columns { "1" } else { "0" }));
            sheet_protection.push_attribute(("insertRows", if protection.insert_rows { "1" } else { "0" }));
            sheet_protection.push_attribute(("insertHyperlinks", if protection.insert_hyperlinks { "1" } else { "0" }));
            sheet_protection.push_attribute(("deleteColumns", if protection.delete_columns { "1" } else { "0" }));
            sheet_protection.push_attribute(("deleteRows", if protection.delete_rows { "1" } else { "0" }));
            sheet_protection.push_attribute(("sort", if protection.sort { "1" } else { "0" }));
            sheet_protection.push_attribute(("autoFilter", if protection.auto_filter { "1" } else { "0" }));
            sheet_protection.push_attribute(("pivotTables", if protection.pivot_tables { "1" } else { "0" }));
            sheet_protection.push_attribute(("objects", if protection.objects { "1" } else { "0" }));
            sheet_protection.push_attribute(("scenarios", if protection.scenarios { "1" } else { "0" }));
            // The password attribute holds the legacy 16-bit verifier hash, never
            // the plaintext. A value loaded from an existing file is already hashed.
            if let Some(ref hash) = protection.password_hash {
                sheet_protection.push_attribute(("password", hash.as_str()));
            } else if let Some(ref pwd) = protection.password {
                let hash = format!("{:04X}", legacy_password_hash(pwd));
                sheet_protection.push_attribute(("password", hash.as_str()));
            }
            writer.write_event(quick_xml::events::Event::Empty(sheet_protection))?;
        }
    }

    // autoFilter
    if let Some(ref auto_filter) = worksheet.auto_filter {
        write_auto_filter(&mut writer, auto_filter)?;
    }

    // mergeCells
    if !worksheet.merged_cells.is_empty() {
        let mut merge_cells = BytesStart::new("mergeCells");
        merge_cells.push_attribute(("count", worksheet.merged_cells.len().to_string().as_str()));
        writer.write_event(quick_xml::events::Event::Start(merge_cells))?;
        for (start, end) in &worksheet.merged_cells {
            let mut merge_cell = BytesStart::new("mergeCell");
            merge_cell.push_attribute(("ref", format!("{}:{}", start, end).as_str()));
            writer.write_event(quick_xml::events::Event::Empty(merge_cell))?;
        }
        writer.write_event(Event::End(BytesEnd::new("mergeCells")))?;
    }

    // conditionalFormatting
    if !worksheet.conditional_formatting.is_empty() {
        for cf in &worksheet.conditional_formatting {
            write_conditional_formatting(&mut writer, cf, dxfs)?;
        }
    }

    // dataValidations (per schema order: after conditionalFormatting, before hyperlinks)
    if !worksheet.data_validations.is_empty() {
        let mut data_validations = BytesStart::new("dataValidations");
        data_validations.push_attribute(("count", worksheet.data_validations.len().to_string().as_str()));
        writer.write_event(quick_xml::events::Event::Start(data_validations))?;

        for ((row, col), validation) in &worksheet.data_validations {
            let coord = format!("{}{}", column_to_letter(*col), row);
            let mut dv = BytesStart::new("dataValidation");
            dv.push_attribute(("type", validation.validation_type.as_str()));
            dv.push_attribute(("allowBlank", if validation.allow_blank { "1" } else { "0" }));
            dv.push_attribute(("showErrorMessage", if validation.show_error { "1" } else { "0" }));
            dv.push_attribute(("showInputMessage", if validation.show_input { "1" } else { "0" }));
            // A loaded rule may span multiple cells; fall back to the key cell
            dv.push_attribute(("sqref", validation.sqref.as_deref().unwrap_or(coord.as_str())));
            writer.write_event(quick_xml::events::Event::Start(dv))?;
            if let Some(ref f1) = validation.formula1 {
                writer.write_event(quick_xml::events::Event::Start(BytesStart::new("formula1")))?;
                writer.write_event(quick_xml::events::Event::Text(BytesText::new(f1)))?;
                writer.write_event(quick_xml::events::Event::End(BytesEnd::new("formula1")))?;
            }
            if let Some(ref f2) = validation.formula2 {
                writer.write_event(quick_xml::events::Event::Start(BytesStart::new("formula2")))?;
                writer.write_event(quick_xml::events::Event::Text(BytesText::new(f2)))?;
                writer.write_event(quick_xml::events::Event::End(BytesEnd::new("formula2")))?;
            }
            writer.write_event(quick_xml::events::Event::End(BytesEnd::new("dataValidation")))?;
        }
        writer.write_event(quick_xml::events::Event::End(BytesEnd::new("dataValidations")))?;
    }

    // hyperlinks: external URLs reference the sheet rels via r:id; internal
    // links (stored with a '#' prefix) use the location attribute.
    let external_links = collect_external_hyperlinks(worksheet);
    let mut internal_links: Vec<((u32, u32), String)> = worksheet
        .cells
        .iter()
        .filter_map(|(key, cd)| {
            cd.hyperlink
                .as_ref()
                .filter(|url| url.starts_with('#'))
                .map(|url| (decode_cell_key(*key), url.clone()))
        })
        .collect();
    internal_links.sort_by_key(|(coord, _)| *coord);

    if !external_links.is_empty() || !internal_links.is_empty() {
        // CT_Hyperlinks has no count attribute, unlike mergeCells/dataValidations
        let hyperlinks = BytesStart::new("hyperlinks");
        writer.write_event(quick_xml::events::Event::Start(hyperlinks))?;

        for (i, ((row, col), _url)) in external_links.iter().enumerate() {
            let coord = format!("{}{}", column_to_letter(*col), row);
            let rel_id = format!("rIdHL{}", i + 1);
            let mut hyperlink = BytesStart::new("hyperlink");
            hyperlink.push_attribute(("ref", coord.as_str()));
            hyperlink.push_attribute(("r:id", rel_id.as_str()));
            writer.write_event(quick_xml::events::Event::Empty(hyperlink))?;
        }

        for ((row, col), url) in internal_links {
            let coord = format!("{}{}", column_to_letter(col), row);
            let mut hyperlink = BytesStart::new("hyperlink");
            hyperlink.push_attribute(("ref", coord.as_str()));
            // The location attribute holds the anchor without the '#' prefix
            hyperlink.push_attribute(("location", url.trim_start_matches('#')));
            writer.write_event(quick_xml::events::Event::Empty(hyperlink))?;
        }

        writer.write_event(quick_xml::events::Event::End(BytesEnd::new("hyperlinks")))?;
    }
    
    // printOptions (per schema order: after hyperlinks, before pageMargins)
    if let Some(ref ps) = worksheet.page_setup {
        write_print_options(&mut writer, ps)?;
    }

    // pageMargins - use PageSetup values if available
    let mut margins = BytesStart::new("pageMargins");
    if let Some(ref ps) = worksheet.page_setup {
        margins.push_attribute(("left", ps.margins.left.to_string().as_str()));
        margins.push_attribute(("right", ps.margins.right.to_string().as_str()));
        margins.push_attribute(("top", ps.margins.top.to_string().as_str()));
        margins.push_attribute(("bottom", ps.margins.bottom.to_string().as_str()));
        margins.push_attribute(("header", ps.margins.header.to_string().as_str()));
        margins.push_attribute(("footer", ps.margins.footer.to_string().as_str()));
    } else {
        margins.push_attribute(("left", "0.75"));
        margins.push_attribute(("right", "0.75"));
        margins.push_attribute(("top", "1"));
        margins.push_attribute(("bottom", "1"));
        margins.push_attribute(("header", "0.5"));
        margins.push_attribute(("footer", "0.5"));
    }
    writer.write_event(Event::Empty(margins))?;

    // pageSetup
    if let Some(ref ps) = worksheet.page_setup {
        write_page_setup(&mut writer, ps)?;
    }

    // legacyDrawing anchors the VML part Excel needs to display comments
    if has_comments {
        let mut legacy = BytesStart::new("legacyDrawing");
        legacy.push_attribute(("r:id", "rIdVml"));
        writer.write_event(quick_xml::events::Event::Empty(legacy))?;
    }

    // tableParts (near the end of CT_Worksheet)
    if !table_rel_ids.is_empty() {
        let mut table_parts = BytesStart::new("tableParts");
        table_parts.push_attribute(("count", table_rel_ids.len().to_string().as_str()));
        writer.write_event(quick_xml::events::Event::Start(table_parts))?;
        for rel_id in table_rel_ids {
            let mut part = BytesStart::new("tablePart");
            part.push_attribute(("r:id", rel_id.as_str()));
            writer.write_event(quick_xml::events::Event::Empty(part))?;
        }
        writer.write_event(quick_xml::events::Event::End(BytesEnd::new("tableParts")))?;
    }

    writer.write_event(quick_xml::events::Event::End(BytesEnd::new("worksheet")))?;

    let result = writer.into_inner().into_inner();
    zip.write_all(&result)?;
    Ok(())
}

pub fn write_comments_xml<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &FileOptions<'static, ExtendedFileOptions>,
    worksheet: &Worksheet,
    sheet_id: u32,
) -> Result<bool> {
    // Collect comments
    let comment_cells: Vec<((u32, u32), String)> = worksheet.cells
        .iter()
        .filter_map(|(key, cell_data)| {
            cell_data.comment.as_ref().map(|comment| {
                let (row, col) = decode_cell_key(*key);
                ((row, col), comment.clone())
            })
        })
        .collect();
    
    if comment_cells.is_empty() {
        return Ok(false); // No comments to write
    }
    
    let path = format!("xl/comments/comment{}.xml", sheet_id);
    zip.start_file(&path, options.clone())?;
    
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut comments_start = BytesStart::new("comments");
    comments_start.push_attribute(("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    writer.write_event(quick_xml::events::Event::Start(comments_start))?;
    
    // authors
    writer.write_event(quick_xml::events::Event::Start(BytesStart::new("authors")))?;
    writer.write_event(quick_xml::events::Event::Start(BytesStart::new("author")))?;
    writer.write_event(quick_xml::events::Event::Text(BytesText::new("RustyPyXL")))?;
    writer.write_event(quick_xml::events::Event::End(BytesEnd::new("author")))?;
    writer.write_event(quick_xml::events::Event::End(BytesEnd::new("authors")))?;
    
    // commentList
    let mut comment_list = BytesStart::new("commentList");
    comment_list.push_attribute(("count", comment_cells.len().to_string().as_str()));
    writer.write_event(quick_xml::events::Event::Start(comment_list))?;
    
    for ((row, col), comment_text) in comment_cells {
        let coord = format!("{}{}", column_to_letter(col), row);
        let mut comment = BytesStart::new("comment");
        comment.push_attribute(("ref", coord.as_str()));
        comment.push_attribute(("authorId", "0"));
        comment.push_attribute(("shapeId", "0"));
        writer.write_event(quick_xml::events::Event::Start(comment))?;
        
        // text
        writer.write_event(quick_xml::events::Event::Start(BytesStart::new("text")))?;
        writer.write_event(quick_xml::events::Event::Start(BytesStart::new("t")))?;
        writer.write_event(quick_xml::events::Event::Text(BytesText::new(&comment_text)))?;
        writer.write_event(quick_xml::events::Event::End(BytesEnd::new("t")))?;
        writer.write_event(quick_xml::events::Event::End(BytesEnd::new("text")))?;
        
        writer.write_event(quick_xml::events::Event::End(BytesEnd::new("comment")))?;
    }
    
    writer.write_event(quick_xml::events::Event::End(BytesEnd::new("commentList")))?;
    writer.write_event(quick_xml::events::Event::End(BytesEnd::new("comments")))?;

    let result = writer.into_inner().into_inner();
    zip.write_all(&result)?;
    Ok(true) // Comments were written
}

/// Write autoFilter element.
fn write_auto_filter<W: std::io::Write>(
    writer: &mut Writer<W>,
    auto_filter: &crate::autofilter::AutoFilter,
) -> Result<()> {
    let mut af = BytesStart::new("autoFilter");
    af.push_attribute(("ref", auto_filter.range.as_str()));

    if auto_filter.columns.is_empty() {
        writer.write_event(Event::Empty(af))?;
    } else {
        writer.write_event(Event::Start(af))?;

        for col_filter in &auto_filter.columns {
            let mut filter_col = BytesStart::new("filterColumn");
            filter_col.push_attribute(("colId", col_filter.column_id.to_string().as_str()));
            if !col_filter.show_button {
                filter_col.push_attribute(("hiddenButton", "1"));
            }
            writer.write_event(Event::Start(filter_col))?;

            match &col_filter.filter {
                FilterType::Values(values) => {
                    writer.write_event(Event::Start(BytesStart::new("filters")))?;
                    for value in values {
                        let mut filter = BytesStart::new("filter");
                        filter.push_attribute(("val", value.as_str()));
                        writer.write_event(Event::Empty(filter))?;
                    }
                    writer.write_event(Event::End(BytesEnd::new("filters")))?;
                }
                FilterType::Custom(custom) => {
                    let mut custom_filters = BytesStart::new("customFilters");
                    if !custom.and {
                        custom_filters.push_attribute(("and", "0"));
                    }
                    writer.write_event(Event::Start(custom_filters))?;

                    let mut cf1 = BytesStart::new("customFilter");
                    cf1.push_attribute(("operator", custom.operator1.xml_value()));
                    cf1.push_attribute(("val", custom.value1.as_str()));
                    writer.write_event(Event::Empty(cf1))?;

                    if let (Some(op2), Some(val2)) = (&custom.operator2, &custom.value2) {
                        let mut cf2 = BytesStart::new("customFilter");
                        cf2.push_attribute(("operator", op2.xml_value()));
                        cf2.push_attribute(("val", val2.as_str()));
                        writer.write_event(Event::Empty(cf2))?;
                    }

                    writer.write_event(Event::End(BytesEnd::new("customFilters")))?;
                }
                FilterType::DynamicFilter(df) => {
                    let mut dyn_filter = BytesStart::new("dynamicFilter");
                    dyn_filter.push_attribute(("type", df.xml_type()));
                    writer.write_event(Event::Empty(dyn_filter))?;
                }
                FilterType::Top10Filter(top10) => {
                    let mut t10 = BytesStart::new("top10");
                    t10.push_attribute(("top", if top10.top { "1" } else { "0" }));
                    t10.push_attribute(("percent", if top10.percent { "1" } else { "0" }));
                    t10.push_attribute(("val", top10.value.to_string().as_str()));
                    writer.write_event(Event::Empty(t10))?;
                }
                FilterType::ColorFilter(cf) => {
                    let mut color_filter = BytesStart::new("colorFilter");
                    color_filter.push_attribute(("cellColor", if cf.cell_color { "1" } else { "0" }));
                    // Color would be specified via dxfId in real implementation
                    writer.write_event(Event::Empty(color_filter))?;
                }
            }

            writer.write_event(Event::End(BytesEnd::new("filterColumn")))?;
        }

        // Sort state
        if let Some(sort_col) = auto_filter.sort_column {
            let mut sort_state = BytesStart::new("sortState");
            sort_state.push_attribute(("ref", auto_filter.range.as_str()));
            writer.write_event(Event::Start(sort_state))?;

            let mut sort_cond = BytesStart::new("sortCondition");
            if auto_filter.sort_descending {
                sort_cond.push_attribute(("descending", "1"));
            }
            sort_cond.push_attribute(("ref", format!("{}:{}",
                column_to_letter(sort_col + 1),
                column_to_letter(sort_col + 1)).as_str()));
            writer.write_event(Event::Empty(sort_cond))?;

            writer.write_event(Event::End(BytesEnd::new("sortState")))?;
        }

        writer.write_event(Event::End(BytesEnd::new("autoFilter")))?;
    }

    Ok(())
}

/// Write a conditional color element.
fn write_conditional_color<W: std::io::Write>(
    writer: &mut Writer<W>,
    color: &crate::conditional::ConditionalColor,
) -> Result<()> {
    let mut color_elem = BytesStart::new("color");
    if let Some(ref rgb) = color.rgb {
        color_elem.push_attribute(("rgb", rgb.as_str()));
    }
    if let Some(theme) = color.theme {
        color_elem.push_attribute(("theme", theme.to_string().as_str()));
    }
    if let Some(tint) = color.tint {
        color_elem.push_attribute(("tint", tint.to_string().as_str()));
    }
    writer.write_event(Event::Empty(color_elem))?;
    Ok(())
}

/// Write conditionalFormatting element.
/// Excel evaluates text/blank/error/time-period rules through a hidden
/// formula anchored at the top-left cell of the rule's range; without it
/// the rule never matches. Returns None for rule types that don't need one
/// or when the caller supplied formula1 explicitly.
fn implied_rule_formula(
    rule: &crate::conditional::ConditionalRule,
    anchor: &str,
) -> Option<String> {
    if rule.formula1.is_some() {
        return None;
    }
    let text = rule.text.as_deref().unwrap_or("");
    // Embedded quotes in the matched text are doubled in Excel formulas
    let quoted = format!("\"{}\"", text.replace('"', "\"\""));
    match rule.rule_type {
        ConditionalFormatType::ContainsText => Some(format!(
            "NOT(ISERROR(SEARCH({},{})))",
            quoted, anchor
        )),
        ConditionalFormatType::NotContainsText => {
            Some(format!("ISERROR(SEARCH({},{}))", quoted, anchor))
        }
        ConditionalFormatType::BeginsWith => Some(format!(
            "LEFT({},LEN({}))={}",
            anchor, quoted, quoted
        )),
        ConditionalFormatType::EndsWith => Some(format!(
            "RIGHT({},LEN({}))={}",
            anchor, quoted, quoted
        )),
        ConditionalFormatType::ContainsBlanks => {
            Some(format!("LEN(TRIM({}))=0", anchor))
        }
        ConditionalFormatType::NotContainsBlanks => {
            Some(format!("LEN(TRIM({}))>0", anchor))
        }
        ConditionalFormatType::ContainsErrors => Some(format!("ISERROR({})", anchor)),
        ConditionalFormatType::NotContainsErrors => {
            Some(format!("NOT(ISERROR({}))", anchor))
        }
        ConditionalFormatType::TimePeriod => {
            let period = rule.time_period.as_deref()?;
            Some(match period {
                "today" => format!("FLOOR({},1)=TODAY()", anchor),
                "yesterday" => format!("FLOOR({},1)=TODAY()-1", anchor),
                "tomorrow" => format!("FLOOR({},1)=TODAY()+1", anchor),
                "last7Days" => format!(
                    "AND(TODAY()-FLOOR({},1)<=6,FLOOR({},1)<=TODAY())",
                    anchor, anchor
                ),
                "thisWeek" => format!(
                    "AND(TODAY()-ROUNDDOWN({},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({},0)-TODAY()<=7-WEEKDAY(TODAY()))",
                    anchor, anchor
                ),
                "lastWeek" => format!(
                    "AND(TODAY()-ROUNDDOWN({},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN({},0)<(WEEKDAY(TODAY())+7))",
                    anchor, anchor
                ),
                "nextWeek" => format!(
                    "AND(ROUNDDOWN({},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({},0)-TODAY()<(15-WEEKDAY(TODAY())))",
                    anchor, anchor
                ),
                "thisMonth" => format!(
                    "AND(MONTH({})=MONTH(TODAY()),YEAR({})=YEAR(TODAY()))",
                    anchor, anchor
                ),
                "lastMonth" => format!(
                    "AND(MONTH({})=MONTH(EDATE(TODAY(),0-1)),YEAR({})=YEAR(EDATE(TODAY(),0-1)))",
                    anchor, anchor
                ),
                "nextMonth" => format!(
                    "AND(MONTH({})=MONTH(EDATE(TODAY(),0+1)),YEAR({})=YEAR(EDATE(TODAY(),0+1)))",
                    anchor, anchor
                ),
                _ => return None,
            })
        }
        _ => None,
    }
}

/// Default icon-set thresholds: N icons split the range into N equal
/// percent bands (Excel's own defaults), e.g. 3 icons -> 0/33/67.
fn default_icon_thresholds(icon_count: u32) -> Vec<(String, String)> {
    (0..icon_count)
        .map(|i| {
            (
                "percent".to_string(),
                ((100 * i) / icon_count).to_string(),
            )
        })
        .collect()
}

fn write_conditional_formatting<W: std::io::Write>(
    writer: &mut Writer<W>,
    cf: &crate::conditional::ConditionalFormatting,
    dxfs: &[ConditionalFormat],
) -> Result<()> {
    let mut cond_fmt = BytesStart::new("conditionalFormatting");
    cond_fmt.push_attribute(("sqref", cf.range.as_str()));
    writer.write_event(Event::Start(cond_fmt))?;

    // Top-left cell of the range, used to anchor implied formulas
    let anchor: String = cf
        .range
        .split([':', ' '])
        .next()
        .unwrap_or("A1")
        .replace('$', "");

    for rule in &cf.rules {
        let mut cf_rule = BytesStart::new("cfRule");
        cf_rule.push_attribute(("type", rule.rule_type.xml_value()));

        // dxfId ties the rule to its differential format in styles.xml;
        // without it the rule matches but applies no formatting
        if let Some(ref fmt) = rule.format {
            if let Some(dxf_id) = dxfs.iter().position(|d| d == fmt) {
                cf_rule.push_attribute(("dxfId", dxf_id.to_string().as_str()));
            }
        }

        cf_rule.push_attribute(("priority", rule.priority.to_string().as_str()));

        // Operator for cellIs rules
        if rule.rule_type == ConditionalFormatType::CellIs {
            if let Some(ref op) = rule.operator {
                cf_rule.push_attribute(("operator", op.xml_value()));
            }
        }

        // Top10 attributes
        if rule.rule_type == ConditionalFormatType::Top10 {
            if let Some(rank) = rule.rank {
                cf_rule.push_attribute(("rank", rank.to_string().as_str()));
            }
            if rule.percent {
                cf_rule.push_attribute(("percent", "1"));
            }
            if rule.bottom {
                cf_rule.push_attribute(("bottom", "1"));
            }
        }

        // AboveAverage attributes
        if rule.rule_type == ConditionalFormatType::AboveAverage {
            if !rule.above_average {
                cf_rule.push_attribute(("aboveAverage", "0"));
            }
            if rule.equal_average {
                cf_rule.push_attribute(("equalAverage", "1"));
            }
            if let Some(std_dev) = rule.std_dev {
                cf_rule.push_attribute(("stdDev", std_dev.to_string().as_str()));
            }
        }

        // Required timePeriod attribute for date rules
        if rule.rule_type == ConditionalFormatType::TimePeriod {
            if let Some(ref period) = rule.time_period {
                cf_rule.push_attribute(("timePeriod", period.as_str()));
            }
        }

        // Text value for text rules
        if let Some(ref text) = rule.text {
            cf_rule.push_attribute(("text", text.as_str()));
        }

        if rule.stop_if_true {
            cf_rule.push_attribute(("stopIfTrue", "1"));
        }

        writer.write_event(Event::Start(cf_rule))?;

        // Write formula if present, or the formula Excel requires for
        // text/blank/error/time-period rules
        let implied = implied_rule_formula(rule, &anchor);
        if let Some(formula) = rule.formula1.as_deref().or(implied.as_deref()) {
            writer.write_event(Event::Start(BytesStart::new("formula")))?;
            writer.write_event(Event::Text(BytesText::new(formula)))?;
            writer.write_event(Event::End(BytesEnd::new("formula")))?;
        }
        if let Some(ref formula) = rule.formula2 {
            writer.write_event(Event::Start(BytesStart::new("formula")))?;
            writer.write_event(Event::Text(BytesText::new(formula)))?;
            writer.write_event(Event::End(BytesEnd::new("formula")))?;
        }

        // ColorScale
        if let Some(ref cs) = rule.color_scale {
            writer.write_event(Event::Start(BytesStart::new("colorScale")))?;

            // cfvo elements
            let mut cfvo1 = BytesStart::new("cfvo");
            cfvo1.push_attribute(("type", cs.min_type.as_str()));
            if let Some(ref val) = cs.min_value {
                cfvo1.push_attribute(("val", val.as_str()));
            }
            writer.write_event(Event::Empty(cfvo1))?;

            if let (Some(ref mid_type), Some(_)) = (&cs.mid_type, &cs.mid_color) {
                let mut cfvo2 = BytesStart::new("cfvo");
                cfvo2.push_attribute(("type", mid_type.as_str()));
                if let Some(ref val) = cs.mid_value {
                    cfvo2.push_attribute(("val", val.as_str()));
                }
                writer.write_event(Event::Empty(cfvo2))?;
            }

            let mut cfvo3 = BytesStart::new("cfvo");
            cfvo3.push_attribute(("type", cs.max_type.as_str()));
            if let Some(ref val) = cs.max_value {
                cfvo3.push_attribute(("val", val.as_str()));
            }
            writer.write_event(Event::Empty(cfvo3))?;

            // color elements
            write_conditional_color(&mut *writer, &cs.min_color)?;

            if let Some(ref mid_color) = cs.mid_color {
                write_conditional_color(&mut *writer, mid_color)?;
            }

            write_conditional_color(&mut *writer, &cs.max_color)?;

            writer.write_event(Event::End(BytesEnd::new("colorScale")))?;
        }

        // DataBar
        if let Some(ref db) = rule.data_bar {
            let mut data_bar = BytesStart::new("dataBar");
            if !db.show_value {
                data_bar.push_attribute(("showValue", "0"));
            }
            writer.write_event(Event::Start(data_bar))?;

            let mut cfvo1 = BytesStart::new("cfvo");
            cfvo1.push_attribute(("type", db.min_type.as_str()));
            if let Some(ref val) = db.min_value {
                cfvo1.push_attribute(("val", val.as_str()));
            }
            writer.write_event(Event::Empty(cfvo1))?;

            let mut cfvo2 = BytesStart::new("cfvo");
            cfvo2.push_attribute(("type", db.max_type.as_str()));
            if let Some(ref val) = db.max_value {
                cfvo2.push_attribute(("val", val.as_str()));
            }
            writer.write_event(Event::Empty(cfvo2))?;

            write_conditional_color(&mut *writer, &db.fill_color)?;

            writer.write_event(Event::End(BytesEnd::new("dataBar")))?;
        }

        // IconSet
        if let Some(ref is) = rule.icon_set {
            let mut icon_set = BytesStart::new("iconSet");
            icon_set.push_attribute(("iconSet", is.style.xml_type()));
            if !is.show_value {
                icon_set.push_attribute(("showValue", "0"));
            }
            if is.reverse {
                icon_set.push_attribute(("reverse", "1"));
            }
            writer.write_event(Event::Start(icon_set))?;

            // CT_IconSet requires one cfvo per icon (>= 2); an empty
            // thresholds list previously produced files Excel repairs away
            let icon_count: u32 = is
                .style
                .xml_type()
                .chars()
                .next()
                .and_then(|c| c.to_digit(10))
                .unwrap_or(3);
            let defaults;
            let thresholds: &[(String, String)] = if is.thresholds.is_empty() {
                defaults = default_icon_thresholds(icon_count);
                &defaults
            } else {
                &is.thresholds
            };

            for (threshold_type, threshold_val) in thresholds {
                let mut cfvo = BytesStart::new("cfvo");
                cfvo.push_attribute(("type", threshold_type.as_str()));
                if !threshold_val.is_empty() {
                    cfvo.push_attribute(("val", threshold_val.as_str()));
                }
                writer.write_event(Event::Empty(cfvo))?;
            }

            writer.write_event(Event::End(BytesEnd::new("iconSet")))?;
        }

        writer.write_event(Event::End(BytesEnd::new("cfRule")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("conditionalFormatting")))?;

    Ok(())
}

/// Write pageSetup element.
fn write_page_setup<W: std::io::Write>(
    writer: &mut Writer<W>,
    ps: &crate::pagesetup::PageSetup,
) -> Result<()> {
    let mut page_setup = BytesStart::new("pageSetup");
    page_setup.push_attribute(("paperSize", ps.paper_size.code().to_string().as_str()));

    if ps.orientation == Orientation::Landscape {
        page_setup.push_attribute(("orientation", "landscape"));
    } else {
        page_setup.push_attribute(("orientation", "portrait"));
    }

    if ps.scale != 100 {
        page_setup.push_attribute(("scale", ps.scale.to_string().as_str()));
    }

    if let Some(fit_w) = ps.fit_to_width {
        page_setup.push_attribute(("fitToWidth", fit_w.to_string().as_str()));
    }
    if let Some(fit_h) = ps.fit_to_height {
        page_setup.push_attribute(("fitToHeight", fit_h.to_string().as_str()));
    }

    if let Some(first_page) = ps.first_page_number {
        page_setup.push_attribute(("firstPageNumber", first_page.to_string().as_str()));
        page_setup.push_attribute(("useFirstPageNumber", "1"));
    }

    if ps.black_and_white {
        page_setup.push_attribute(("blackAndWhite", "1"));
    }
    if ps.draft {
        page_setup.push_attribute(("draft", "1"));
    }

    if let Some(hdpi) = ps.horizontal_dpi {
        page_setup.push_attribute(("horizontalDpi", hdpi.to_string().as_str()));
    }
    if let Some(vdpi) = ps.vertical_dpi {
        page_setup.push_attribute(("verticalDpi", vdpi.to_string().as_str()));
    }

    if ps.copies > 1 {
        page_setup.push_attribute(("copies", ps.copies.to_string().as_str()));
    }

    writer.write_event(Event::Empty(page_setup))?;

    // headerFooter
    let hf = &ps.header_footer;
    if hf.odd_header.is_some() || hf.odd_footer.is_some() {
        let mut header_footer = BytesStart::new("headerFooter");
        if hf.different_odd_even {
            header_footer.push_attribute(("differentOddEven", "1"));
        }
        if hf.different_first {
            header_footer.push_attribute(("differentFirst", "1"));
        }
        writer.write_event(Event::Start(header_footer))?;

        if let Some(ref h) = hf.odd_header {
            writer.write_event(Event::Start(BytesStart::new("oddHeader")))?;
            writer.write_event(Event::Text(BytesText::new(&h.to_string())))?;
            writer.write_event(Event::End(BytesEnd::new("oddHeader")))?;
        }
        if let Some(ref f) = hf.odd_footer {
            writer.write_event(Event::Start(BytesStart::new("oddFooter")))?;
            writer.write_event(Event::Text(BytesText::new(&f.to_string())))?;
            writer.write_event(Event::End(BytesEnd::new("oddFooter")))?;
        }

        writer.write_event(Event::End(BytesEnd::new("headerFooter")))?;
    }

    Ok(())
}

/// Write the printOptions element (per CT_Worksheet schema order it must
/// precede pageMargins, so it cannot live inside write_page_setup).
fn write_print_options<W: std::io::Write>(
    writer: &mut Writer<W>,
    ps: &crate::pagesetup::PageSetup,
) -> Result<()> {
    if ps.print_gridlines || ps.print_headings || ps.center_horizontally || ps.center_vertically {
        let mut print_options = BytesStart::new("printOptions");
        if ps.print_gridlines {
            print_options.push_attribute(("gridLines", "1"));
        }
        if ps.print_headings {
            print_options.push_attribute(("headings", "1"));
        }
        if ps.center_horizontally {
            print_options.push_attribute(("horizontalCentered", "1"));
        }
        if ps.center_vertically {
            print_options.push_attribute(("verticalCentered", "1"));
        }
        writer.write_event(Event::Empty(print_options))?;
    }

    Ok(())
}

/// Write the legacy VML drawing part that anchors comment boxes.
/// Excel ignores comments entirely without one Note shape per comment.
pub fn write_vml_drawing<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &FileOptions<'static, ExtendedFileOptions>,
    worksheet: &Worksheet,
    sheet_id: u32,
) -> Result<()> {
    let path = format!("xl/drawings/vmlDrawing{}.vml", sheet_id);
    zip.start_file(&path, options.clone())?;

    let mut comment_cells: Vec<(u32, u32)> = worksheet
        .cells
        .iter()
        .filter(|(_, cd)| cd.comment.is_some())
        .map(|(key, _)| decode_cell_key(*key))
        .collect();
    comment_cells.sort_unstable();

    let mut xml = String::with_capacity(1024 + comment_cells.len() * 768);
    xml.push_str(
        r#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>"#,
    );

    for (i, (row, col)) in comment_cells.iter().enumerate() {
        // VML anchors are 0-based; place the box one column to the right
        let r0 = row.saturating_sub(1);
        let c0 = col.saturating_sub(1);
        xml.push_str(&format!(
            r##"
<v:shape id="_x0000_s{id}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:{z};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">
<v:fill color2="#ffffe1"/><v:shadow color="black" obscured="t"/><v:path o:connecttype="none"/>
<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>
<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>
<x:Anchor>{a1}, 15, {a2}, 2, {a3}, 31, {a4}, 9</x:Anchor>
<x:AutoFill>False</x:AutoFill><x:Row>{r}</x:Row><x:Column>{c}</x:Column></x:ClientData>
</v:shape>"##,
            id = 1025 + i,
            z = i + 1,
            a1 = c0 + 1,
            a2 = r0,
            a3 = c0 + 3,
            a4 = r0 + 4,
            r = r0,
            c = c0,
        ));
    }

    xml.push_str("\n</xml>");
    zip.write_all(xml.as_bytes())?;
    Ok(())
}

/// Write a table XML file.
pub fn write_table_xml<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &FileOptions<'static, ExtendedFileOptions>,
    table: &crate::table::Table,
    table_id: u32,
) -> Result<()> {
    let path = format!("xl/tables/table{}.xml", table_id);
    zip.start_file(&path, options.clone())?;

    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut table_start = BytesStart::new("table");
    table_start.push_attribute(("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    // Use the workbook-assigned id, not table.id, to guarantee uniqueness
    table_start.push_attribute(("id", table_id.to_string().as_str()));
    table_start.push_attribute(("name", table.name.as_str()));
    table_start.push_attribute(("displayName", table.display_name.as_str()));
    table_start.push_attribute(("ref", table.range.as_str()));

    if !table.header_row {
        table_start.push_attribute(("headerRowCount", "0"));
    }
    if table.totals_row {
        table_start.push_attribute(("totalsRowCount", "1"));
    }

    writer.write_event(Event::Start(table_start))?;

    // autoFilter: the filter range must exclude the totals row
    if table.auto_filter {
        let filter_ref = if table.totals_row {
            match crate::utils::parse_range(&table.range) {
                Ok(((r1, c1), (r2, c2))) if r2 > r1 => format!(
                    "{}{}:{}{}",
                    column_to_letter(c1),
                    r1,
                    column_to_letter(c2),
                    r2 - 1
                ),
                _ => table.range.clone(),
            }
        } else {
            table.range.clone()
        };
        let mut af = BytesStart::new("autoFilter");
        af.push_attribute(("ref", filter_ref.as_str()));
        writer.write_event(Event::Empty(af))?;
    }

    // tableColumns
    let mut table_columns = BytesStart::new("tableColumns");
    table_columns.push_attribute(("count", table.columns.len().to_string().as_str()));
    writer.write_event(Event::Start(table_columns))?;

    for col in &table.columns {
        let mut tc = BytesStart::new("tableColumn");
        tc.push_attribute(("id", col.id.to_string().as_str()));
        tc.push_attribute(("name", col.name.as_str()));

        if let Some(xml_name) = col.totals_row_function.xml_name() {
            tc.push_attribute(("totalsRowFunction", xml_name));
        }
        if let Some(ref label) = col.totals_row_label {
            tc.push_attribute(("totalsRowLabel", label.as_str()));
        }

        if col.calculated_column_formula.is_some() {
            writer.write_event(Event::Start(tc))?;
            if let Some(ref formula) = col.calculated_column_formula {
                let calc = BytesStart::new("calculatedColumnFormula");
                writer.write_event(Event::Start(calc))?;
                writer.write_event(Event::Text(BytesText::new(formula)))?;
                writer.write_event(Event::End(BytesEnd::new("calculatedColumnFormula")))?;
            }
            writer.write_event(Event::End(BytesEnd::new("tableColumn")))?;
        } else {
            writer.write_event(Event::Empty(tc))?;
        }
    }

    writer.write_event(Event::End(BytesEnd::new("tableColumns")))?;

    // tableStyleInfo
    let mut style_info = BytesStart::new("tableStyleInfo");
    style_info.push_attribute(("name", table.style.style_name().as_str()));
    style_info.push_attribute(("showFirstColumn", if table.show_first_column { "1" } else { "0" }));
    style_info.push_attribute(("showLastColumn", if table.show_last_column { "1" } else { "0" }));
    style_info.push_attribute(("showRowStripes", if table.show_row_stripes { "1" } else { "0" }));
    style_info.push_attribute(("showColumnStripes", if table.show_column_stripes { "1" } else { "0" }));
    writer.write_event(Event::Empty(style_info))?;

    writer.write_event(Event::End(BytesEnd::new("table")))?;

    let result = writer.into_inner().into_inner();
    zip.write_all(&result)?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::style::{Fill, BorderStyle};

    #[test]
    fn test_escape_xml_strips_illegal_control_chars() {
        assert_eq!(escape_xml("a\x01b\x08c\x0bd\x1fe"), "abcde");
        assert_eq!(escape_xml("tab\tnl\ncr\r"), "tab\tnl\ncr\r");
        assert_eq!(escape_xml("a<b\x02&c"), "a&lt;b&amp;c");
    }

    #[test]
    fn test_strip_illegal_xml_chars() {
        assert_eq!(strip_illegal_xml_chars("a\x00b\x1fc"), "abc");
        assert_eq!(strip_illegal_xml_chars("plain <kept> &"), "plain <kept> &");
    }

    #[test]
    fn test_legacy_password_hash_known_vectors() {
        // Reference values from openpyxl's hash_password / MS-XLS Method1.
        assert_eq!(legacy_password_hash("test"), 0xCBEB);
        assert_eq!(legacy_password_hash(""), 0xCE4B);
        assert_eq!(legacy_password_hash("password"), 0x83AF);
    }

    #[test]
    fn test_non_finite_numbers_become_error_cells() {
        let map = HashMap::new();
        for v in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            let mut buf = String::new();
            let cell = CellData {
                value: CellValue::Number(v),
                ..Default::default()
            };
            write_cell_direct(&mut buf, "A1", &cell, cell.style_index, &map);
            assert_eq!(buf, r#"<c r="A1" t="e"><v>#NUM!</v></c>"#);

            let mut buf2 = String::new();
            format_cell_value(&mut buf2, "A1", &CellValue::Number(v));
            assert_eq!(buf2, r#"<c r="A1" t="e"><v>#NUM!</v></c>"#);
        }
    }

    #[test]
    fn test_date_value_is_escaped() {
        let mut buf = String::new();
        format_cell_value(&mut buf, "A1", &CellValue::Date("<bad>&".to_string()));
        assert_eq!(buf, r#"<c r="A1" t="d"><v>&lt;bad&gt;&amp;</v></c>"#);
    }

    #[test]
    fn test_write_color_attr_theme() {
        let mut xml = String::new();
        write_color_attr(&mut xml, "fgColor", "theme:0");
        assert_eq!(xml, r#"<fgColor theme="0"/>"#);
    }

    #[test]
    fn test_write_color_attr_theme_with_index() {
        let mut xml = String::new();
        write_color_attr(&mut xml, "color", "theme:4");
        assert_eq!(xml, r#"<color theme="4"/>"#);
    }

    #[test]
    fn test_write_color_attr_rgb_hex() {
        let mut xml = String::new();
        write_color_attr(&mut xml, "fgColor", "FF0000");
        assert_eq!(xml, r#"<fgColor rgb="FFFF0000"/>"#);
    }

    #[test]
    fn test_write_color_attr_rgb_with_hash() {
        let mut xml = String::new();
        write_color_attr(&mut xml, "bgColor", "#00FF00");
        assert_eq!(xml, r#"<bgColor rgb="FF00FF00"/>"#);
    }

    #[test]
    fn test_write_color_attr_argb_8char() {
        // 8-char aRGB values from XML roundtrip should not get double-prefixed
        let mut xml = String::new();
        write_color_attr(&mut xml, "fgColor", "#0000FF00");
        assert_eq!(xml, r#"<fgColor rgb="0000FF00"/>"#);
    }

    #[test]
    fn test_write_color_attr_argb_8char_no_hash() {
        let mut xml = String::new();
        write_color_attr(&mut xml, "fgColor", "FF00FF00");
        assert_eq!(xml, r#"<fgColor rgb="FF00FF00"/>"#);
    }

    #[test]
    fn test_write_fill_xml_theme_fg_color() {
        let fill = Fill {
            pattern_type: Some("solid".to_string()),
            fg_color: Some("theme:0".to_string()),
            bg_color: None,
        };
        let mut xml = String::new();
        write_fill_xml(&mut xml, &fill);
        assert!(xml.contains(r#"<fgColor theme="0"/>"#));
        assert!(!xml.contains("FFtheme"));
    }

    #[test]
    fn test_write_fill_xml_theme_bg_color() {
        let fill = Fill {
            pattern_type: Some("solid".to_string()),
            fg_color: None,
            bg_color: Some("theme:2".to_string()),
        };
        let mut xml = String::new();
        write_fill_xml(&mut xml, &fill);
        assert!(xml.contains(r#"<bgColor theme="2"/>"#));
        assert!(!xml.contains("FFtheme"));
    }

    #[test]
    fn test_write_fill_xml_rgb_color() {
        let fill = Fill {
            pattern_type: Some("solid".to_string()),
            fg_color: Some("FFFF00".to_string()),
            bg_color: None,
        };
        let mut xml = String::new();
        write_fill_xml(&mut xml, &fill);
        assert!(xml.contains(r#"<fgColor rgb="FFFFFF00"/>"#));
    }

    #[test]
    fn test_write_border_side_theme_color() {
        let side = Some(BorderStyle {
            style: "thin".to_string(),
            color: Some("theme:1".to_string()),
        });
        let mut xml = String::new();
        write_border_side(&mut xml, "left", &side);
        assert!(xml.contains(r#"<color theme="1"/>"#));
        assert!(!xml.contains("FFtheme"));
    }

    #[test]
    fn test_write_border_side_rgb_color() {
        let side = Some(BorderStyle {
            style: "thick".to_string(),
            color: Some("FF0000".to_string()),
        });
        let mut xml = String::new();
        write_border_side(&mut xml, "left", &side);
        assert!(xml.contains(r#"<color rgb="FFFF0000"/>"#));
    }
}

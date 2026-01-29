use crate::cell::InternedString;
use crate::worksheet::{decode_cell_key, Worksheet, CellData};
use crate::cell::CellValue;
use crate::utils::column_to_letter;
use crate::error::Result;
use crate::autofilter::FilterType;
use crate::conditional::{ConditionalFormatType, ConditionalOperator};
use crate::pagesetup::Orientation;
use zip::write::{FileOptions, ExtendedFileOptions};
use zip::ZipWriter;
use quick_xml::Writer;
use quick_xml::events::{BytesStart, BytesEnd, BytesText, Event};
use std::io::{Write, Cursor, Seek};
use std::collections::HashMap;
use rayon::prelude::*;

/// Escape XML special characters in text content.
#[inline]
fn escape_xml(s: &str) -> std::borrow::Cow<'_, str> {
    if s.bytes().any(|b| matches!(b, b'<' | b'>' | b'&' | b'"' | b'\'')) {
        let mut escaped = String::with_capacity(s.len() + 8);
        for c in s.chars() {
            match c {
                '<' => escaped.push_str("&lt;"),
                '>' => escaped.push_str("&gt;"),
                '&' => escaped.push_str("&amp;"),
                '"' => escaped.push_str("&quot;"),
                '\'' => escaped.push_str("&apos;"),
                _ => escaped.push(c),
            }
        }
        std::borrow::Cow::Owned(escaped)
    } else {
        std::borrow::Cow::Borrowed(s)
    }
}

/// Write cell data directly to a string buffer (fast path, no quick_xml overhead).
/// Uses itoa/ryu for fast number formatting.
#[inline]
fn write_cell_direct(
    buf: &mut String,
    coord: &str,
    cell_data: &CellData,
    shared_string_map: &HashMap<InternedString, usize>,
) {
    match &cell_data.value {
        CellValue::String(s) => {
            if let Some(&idx) = shared_string_map.get(s) {
                // Shared string reference - use itoa for fast integer formatting
                buf.push_str("<c r=\"");
                buf.push_str(coord);
                buf.push_str("\" t=\"s\"><v>");
                buf.push_str(itoa::Buffer::new().format(idx));
                buf.push_str("</v></c>");
            } else {
                // Inline string
                let escaped = escape_xml(s.as_ref());
                buf.push_str("<c r=\"");
                buf.push_str(coord);
                buf.push_str("\" t=\"inlineStr\"><is><t>");
                buf.push_str(&escaped);
                buf.push_str("</t></is></c>");
            }
        }
        CellValue::Number(n) => {
            // Use ryu for fast float formatting
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
            buf.push_str("<c r=\"");
            buf.push_str(coord);
            buf.push_str("\" t=\"d\"><v>");
            buf.push_str(d);
            buf.push_str("</v></c>");
        }
        CellValue::Empty => {
            buf.push_str("<c r=\"");
            buf.push_str(coord);
            buf.push_str("\"/>");
        }
    }
}

pub fn write_content_types<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &FileOptions<'static, ExtendedFileOptions>,
    sheet_count: usize,
    has_shared_strings: bool,
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
    sheet_names: &[String],
    named_ranges: &[(String, String)],
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
    view.push_attribute(("activeTab", "0"));
    writer.write_event(quick_xml::events::Event::Empty(view))?;
    writer.write_event(quick_xml::events::Event::End(BytesEnd::new("bookViews")))?;
    
    // sheets
    writer.write_event(quick_xml::events::Event::Start(BytesStart::new("sheets")))?;
    for (idx, name) in sheet_names.iter().enumerate() {
        let sheet_id = (idx + 1) as u32;
        let r_id = format!("rId{}", idx + 1);
        let mut sheet = BytesStart::new("sheet");
        sheet.push_attribute(("name", name.as_str()));
        sheet.push_attribute(("sheetId", sheet_id.to_string().as_str()));
        sheet.push_attribute(("state", "visible"));
        sheet.push_attribute(("r:id", r_id.as_str()));
        writer.write_event(quick_xml::events::Event::Empty(sheet))?;
    }
    writer.write_event(quick_xml::events::Event::End(BytesEnd::new("sheets")))?;
    
    // definedNames (named ranges)
    if !named_ranges.is_empty() {
        writer.write_event(quick_xml::events::Event::Start(BytesStart::new("definedNames")))?;
        for (name, range) in named_ranges {
            let mut defined_name = BytesStart::new("definedName");
            defined_name.push_attribute(("name", name.as_str()));
            writer.write_event(quick_xml::events::Event::Start(defined_name))?;
            writer.write_event(quick_xml::events::Event::Text(BytesText::new(range)))?;
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
        writer.write_event(quick_xml::events::Event::Text(BytesText::new(s.as_ref())))?;
        writer.write_event(quick_xml::events::Event::End(BytesEnd::new("t")))?;
        writer.write_event(quick_xml::events::Event::End(BytesEnd::new("si")))?;
    }
    
    writer.write_event(quick_xml::events::Event::End(BytesEnd::new("sst")))?;
    
    let result = writer.into_inner().into_inner();
    zip.write_all(&result)?;
    Ok(())
}

pub fn write_styles_xml<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &FileOptions<'static, ExtendedFileOptions>,
) -> Result<()> {
    zip.start_file("xl/styles.xml", options.clone())?;
    
    // Basic minimal styles.xml
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="0"/>
<fonts count="1">
<font>
<sz val="11"/>
<color theme="1"/>
<name val="Calibri"/>
<family val="2"/>
<scheme val="minor"/>
</font>
</fonts>
<fills count="2">
<fill><patternFill/></fill>
<fill><patternFill patternType="gray125"/></fill>
</fills>
<borders count="1">
<border><left/><right/><top/><bottom/><diagonal/></border>
</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellXfs>
<cellStyles count="1">
<cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
</styleSheet>"#;
    
    zip.write_all(styles_xml.as_bytes())?;
    Ok(())
}

pub fn write_worksheet_xml<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &FileOptions<'static, ExtendedFileOptions>,
    worksheet: &Worksheet,
    sheet_id: u32,
    shared_string_map: &HashMap<InternedString, usize>,
    _has_comments: bool,
) -> Result<()> {
    let path = format!("xl/worksheets/sheet{}.xml", sheet_id);
    zip.start_file(&path, options.clone())?;

    // Pre-allocate buffer based on estimated size (rough estimate: 100 bytes per cell)
    let estimated_size = worksheet.cells.len() * 100;
    let mut writer = Writer::new(Cursor::new(Vec::with_capacity(estimated_size)));
    let mut worksheet_start = BytesStart::new("worksheet");
    worksheet_start.push_attribute(("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    writer.write_event(quick_xml::events::Event::Start(worksheet_start))?;
    
    // sheetPr
    writer.write_event(quick_xml::events::Event::Start(BytesStart::new("sheetPr")))?;
    let mut outline = BytesStart::new("outlinePr");
    outline.push_attribute(("summaryBelow", "1"));
    outline.push_attribute(("summaryRight", "1"));
    writer.write_event(quick_xml::events::Event::Empty(outline))?;
    writer.write_event(quick_xml::events::Event::Empty(BytesStart::new("pageSetUpPr")))?;
    writer.write_event(quick_xml::events::Event::End(BytesEnd::new("sheetPr")))?;
    
    // sheetProtection (if enabled)
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
            if let Some(ref pwd) = protection.password {
                // Excel stores password as a hash, but for now we'll just store it
                // In a real implementation, you'd hash it properly
                sheet_protection.push_attribute(("password", pwd.as_str()));
            }
            writer.write_event(quick_xml::events::Event::Empty(sheet_protection))?;
        }
    }
    
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
    let mut view = BytesStart::new("sheetView");
    view.push_attribute(("workbookViewId", "0"));
    writer.write_event(quick_xml::events::Event::Empty(view))?;
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
    let mut rows: HashMap<u32, Vec<((u32, u32), &CellData)>> = HashMap::with_capacity(estimated_rows);
    for (key, cell_data) in &worksheet.cells {
        let (row, col) = decode_cell_key(*key);
        rows.entry(row).or_insert_with(Vec::new).push(((row, col), cell_data));
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
                    let mut sorted_cells: Vec<_> = cells.iter().cloned().collect();
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
                        write_cell_direct(&mut buf, &coord, cell_data, shared_string_map);
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
                write_cell_direct(&mut buf, &coord, cell_data, shared_string_map);
            }

            buf.push_str("</row>");
        }
        buf
    };

    // Write the cell buffer to the XML writer
    writer.get_mut().write_all(cell_buf.as_bytes())?;

    writer.write_event(Event::End(BytesEnd::new("sheetData")))?;

    // sheetProtection (moved to correct position per OOXML spec)
    // Note: We handle protection earlier but per spec it should be after sheetData

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
            write_conditional_formatting(&mut writer, cf)?;
        }
    }

    // hyperlinks
    let hyperlink_cells: Vec<((u32, u32), String)> = worksheet.cells
        .iter()
        .filter_map(|(key, cell_data)| {
            cell_data.hyperlink.as_ref().map(|url| {
                let (row, col) = decode_cell_key(*key);
                ((row, col), url.clone())
            })
        })
        .collect();
    
    if !hyperlink_cells.is_empty() {
        let mut hyperlinks = BytesStart::new("hyperlinks");
        hyperlinks.push_attribute(("count", hyperlink_cells.len().to_string().as_str()));
        writer.write_event(quick_xml::events::Event::Start(hyperlinks))?;
        
        for ((row, col), url) in hyperlink_cells {
            let coord = format!("{}{}", column_to_letter(col), row);
            let mut hyperlink = BytesStart::new("hyperlink");
            hyperlink.push_attribute(("ref", coord.as_str()));
            
            // Check if it's an external URL or internal reference
            if url.starts_with("http://") || url.starts_with("https://") || url.starts_with("mailto:") {
                // External URL - would need relationship ID in full implementation
                // For now, store as location attribute
                hyperlink.push_attribute(("location", url.as_str()));
            } else if url.starts_with('#') {
                // Internal reference
                hyperlink.push_attribute(("location", url.as_str()));
            } else {
                // Assume it's a URL
                hyperlink.push_attribute(("location", url.as_str()));
            }
            
            writer.write_event(quick_xml::events::Event::Empty(hyperlink))?;
        }
        
        writer.write_event(quick_xml::events::Event::End(BytesEnd::new("hyperlinks")))?;
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

    // dataValidations
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
            dv.push_attribute(("sqref", coord.as_str()));
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
fn write_conditional_formatting<W: std::io::Write>(
    writer: &mut Writer<W>,
    cf: &crate::conditional::ConditionalFormatting,
) -> Result<()> {
    let mut cond_fmt = BytesStart::new("conditionalFormatting");
    cond_fmt.push_attribute(("sqref", cf.range.as_str()));
    writer.write_event(Event::Start(cond_fmt))?;

    for rule in &cf.rules {
        let mut cf_rule = BytesStart::new("cfRule");

        // Type
        let type_str = match &rule.rule_type {
            ConditionalFormatType::CellIs => "cellIs",
            ConditionalFormatType::Expression => "expression",
            ConditionalFormatType::ColorScale => "colorScale",
            ConditionalFormatType::DataBar => "dataBar",
            ConditionalFormatType::IconSet => "iconSet",
            ConditionalFormatType::Top10 => "top10",
            ConditionalFormatType::AboveAverage => "aboveAverage",
            ConditionalFormatType::DuplicateValues => "duplicateValues",
            ConditionalFormatType::UniqueValues => "uniqueValues",
            ConditionalFormatType::ContainsText => "containsText",
            ConditionalFormatType::NotContainsText => "notContainsText",
            ConditionalFormatType::BeginsWith => "beginsWith",
            ConditionalFormatType::EndsWith => "endsWith",
            ConditionalFormatType::ContainsBlanks => "containsBlanks",
            ConditionalFormatType::NotContainsBlanks => "notContainsBlanks",
            ConditionalFormatType::ContainsErrors => "containsErrors",
            ConditionalFormatType::NotContainsErrors => "notContainsErrors",
            ConditionalFormatType::TimePeriod => "timePeriod",
        };
        cf_rule.push_attribute(("type", type_str));
        cf_rule.push_attribute(("priority", rule.priority.to_string().as_str()));

        // Operator for cellIs rules
        if rule.rule_type == ConditionalFormatType::CellIs {
            if let Some(ref op) = rule.operator {
                let op_str = match op {
                    ConditionalOperator::Equal => "equal",
                    ConditionalOperator::NotEqual => "notEqual",
                    ConditionalOperator::GreaterThan => "greaterThan",
                    ConditionalOperator::GreaterThanOrEqual => "greaterThanOrEqual",
                    ConditionalOperator::LessThan => "lessThan",
                    ConditionalOperator::LessThanOrEqual => "lessThanOrEqual",
                    ConditionalOperator::Between => "between",
                    ConditionalOperator::NotBetween => "notBetween",
                };
                cf_rule.push_attribute(("operator", op_str));
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
        }

        // Text value for text rules
        if let Some(ref text) = rule.text {
            cf_rule.push_attribute(("text", text.as_str()));
        }

        if rule.stop_if_true {
            cf_rule.push_attribute(("stopIfTrue", "1"));
        }

        writer.write_event(Event::Start(cf_rule))?;

        // Write formula if present
        if let Some(ref formula) = rule.formula1 {
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

            for (threshold_type, threshold_val) in &is.thresholds {
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

    // printOptions
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
    table_start.push_attribute(("id", table.id.to_string().as_str()));
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

    // autoFilter
    if table.auto_filter {
        let mut af = BytesStart::new("autoFilter");
        af.push_attribute(("ref", table.range.as_str()));
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

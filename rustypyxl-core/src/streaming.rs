//! Streaming write support for memory-efficient Excel file creation.
//!
//! This module provides a write-only workbook that streams rows directly to disk
//! without holding them in memory, similar to openpyxl's write_only mode.

use crate::cell::CellValue;
use crate::error::{Result, RustypyxlError};
use crate::utils::column_to_letter;
use crate::writer::{escape_xml, format_cell_value};

use std::fs::File;
use std::io::{BufWriter, Write};
use zip::write::{ExtendedFileOptions, FileOptions};
use zip::{CompressionMethod, ZipWriter};

/// A streaming sheet that writes rows directly to the ZIP file.
pub struct StreamingSheet {
    name: String,
    current_row: u32,
    max_col: u32,
    finalized: bool,
}

/// A write-only workbook that streams data directly to disk.
///
/// This is much more memory efficient than the standard Workbook for large files,
/// as rows are written immediately and not held in memory.
///
/// # Example
/// ```no_run
/// use rustypyxl_core::streaming::StreamingWorkbook;
/// use rustypyxl_core::CellValue;
/// use std::sync::Arc;
///
/// let mut wb = StreamingWorkbook::new("output.xlsx").unwrap();
/// let mut sheet = wb.create_sheet("Data").unwrap();
///
/// // Write rows - they go directly to disk
/// wb.append_row(&mut sheet, vec![
///     CellValue::String(Arc::from("Name")),
///     CellValue::String(Arc::from("Age")),
/// ]).unwrap();
///
/// for i in 0..1000 {
///     wb.append_row(&mut sheet, vec![
///         CellValue::String(Arc::from(format!("Person {}", i))),
///         CellValue::Number(i as f64),
///     ]).unwrap();
/// }
///
/// wb.close(sheet).unwrap();
/// ```
pub struct StreamingWorkbook {
    zip: ZipWriter<BufWriter<File>>,
    options: FileOptions<'static, ExtendedFileOptions>,
    sheets: Vec<String>,
    current_sheet_idx: Option<usize>,
    sheet_xml_started: bool,
}

impl StreamingWorkbook {
    /// Create a new streaming workbook that writes to the given path.
    pub fn new(path: &str) -> Result<Self> {
        let file = File::create(path)?;
        let writer = BufWriter::with_capacity(1024 * 1024, file); // 1MB buffer
        let zip = ZipWriter::new(writer);

        let options = FileOptions::default()
            .compression_method(CompressionMethod::Deflated)
            .compression_level(Some(1)); // Fast compression

        Ok(StreamingWorkbook {
            zip,
            options,
            sheets: Vec::new(),
            current_sheet_idx: None,
            sheet_xml_started: false,
        })
    }

    /// Create a new sheet. Returns a StreamingSheet handle for writing rows.
    pub fn create_sheet(&mut self, name: &str) -> Result<StreamingSheet> {
        if self.current_sheet_idx.is_some() {
            return Err(RustypyxlError::custom(
                "Must close current sheet before creating a new one"
            ));
        }

        self.sheets.push(name.to_string());
        let idx = self.sheets.len() - 1;
        self.current_sheet_idx = Some(idx);

        // Start the sheet XML file
        let path = format!("xl/worksheets/sheet{}.xml", idx + 1);
        self.zip.start_file(&path, self.options.clone())?;

        // Write sheet header (we'll write sheetData rows as they come)
        self.zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
"#)?;
        self.sheet_xml_started = true;

        Ok(StreamingSheet {
            name: name.to_string(),
            current_row: 0,
            max_col: 0,
            finalized: false,
        })
    }

    /// Append a row to the current sheet.
    pub fn append_row(&mut self, sheet: &mut StreamingSheet, values: Vec<CellValue>) -> Result<()> {
        if self.current_sheet_idx.is_none() {
            return Err(RustypyxlError::custom("No sheet is open"));
        }

        sheet.current_row += 1;
        let row_num = sheet.current_row;

        if values.is_empty() {
            return Ok(());
        }

        // Track max column
        if values.len() as u32 > sheet.max_col {
            sheet.max_col = values.len() as u32;
        }

        // Build row XML
        let mut row_xml = format!("<row r=\"{}\">", row_num);

        for (col_idx, value) in values.iter().enumerate() {
            let col = (col_idx + 1) as u32;
            let coord = format!("{}{}", column_to_letter(col), row_num);
            format_cell_value(&mut row_xml, &coord, value);
        }

        row_xml.push_str("</row>\n");
        self.zip.write_all(row_xml.as_bytes())?;

        Ok(())
    }

    /// Finalize the current sheet.
    fn finalize_sheet(&mut self, sheet: &StreamingSheet) -> Result<()> {
        if !self.sheet_xml_started {
            return Ok(());
        }

        // Close sheetData and worksheet
        self.zip.write_all(b"</sheetData>\n")?;

        // Write page margins
        self.zip.write_all(br#"<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
</worksheet>"#)?;

        self.sheet_xml_started = false;
        self.current_sheet_idx = None;

        Ok(())
    }

    /// Close the workbook and finalize the ZIP file.
    pub fn close(mut self, mut sheet: StreamingSheet) -> Result<()> {
        // Finalize current sheet if open
        self.finalize_sheet(&sheet)?;
        sheet.finalized = true;

        // Write [Content_Types].xml
        self.write_content_types()?;

        // Write _rels/.rels
        self.write_rels()?;

        // Write docProps
        self.write_doc_props()?;

        // Write xl/workbook.xml
        self.write_workbook_xml()?;

        // Write xl/_rels/workbook.xml.rels
        self.write_workbook_rels()?;

        // Write xl/styles.xml
        self.write_styles_xml()?;

        // Finalize ZIP
        self.zip.finish()?;

        Ok(())
    }

    fn write_content_types(&mut self) -> Result<()> {
        self.zip.start_file("[Content_Types].xml", self.options.clone())?;

        let mut content = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
"#);

        for i in 0..self.sheets.len() {
            content.push_str(&format!(
                "<Override PartName=\"/xl/worksheets/sheet{}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n",
                i + 1
            ));
        }

        content.push_str("</Types>");
        self.zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_rels(&mut self) -> Result<()> {
        self.zip.start_file("_rels/.rels", self.options.clone())?;
        self.zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"#)?;
        Ok(())
    }

    fn write_doc_props(&mut self) -> Result<()> {
        self.zip.start_file("docProps/core.xml", self.options.clone())?;
        self.zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">
<dc:creator>rustypyxl</dc:creator>
</cp:coreProperties>"#)?;

        self.zip.start_file("docProps/app.xml", self.options.clone())?;
        self.zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
<Application>rustypyxl</Application>
</Properties>"#)?;
        Ok(())
    }

    fn write_workbook_xml(&mut self) -> Result<()> {
        self.zip.start_file("xl/workbook.xml", self.options.clone())?;

        let mut content = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
"#);

        for (i, name) in self.sheets.iter().enumerate() {
            let escaped_name = escape_xml(name);
            content.push_str(&format!(
                "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"rId{}\"/>\n",
                escaped_name, i + 1, i + 1
            ));
        }

        content.push_str("</sheets>\n</workbook>");
        self.zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_workbook_rels(&mut self) -> Result<()> {
        self.zip.start_file("xl/_rels/workbook.xml.rels", self.options.clone())?;

        let mut content = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
"#);

        for i in 0..self.sheets.len() {
            content.push_str(&format!(
                "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{}.xml\"/>\n",
                i + 1, i + 1
            ));
        }

        content.push_str(&format!(
            "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\n",
            self.sheets.len() + 1
        ));

        content.push_str("</Relationships>");
        self.zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_styles_xml(&mut self) -> Result<()> {
        self.zip.start_file("xl/styles.xml", self.options.clone())?;
        self.zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"#)?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Arc;
    use tempfile::NamedTempFile;

    #[test]
    fn test_streaming_write() {
        let temp = NamedTempFile::new().unwrap();
        let path = temp.path().to_str().unwrap();

        let mut wb = StreamingWorkbook::new(path).unwrap();
        let mut sheet = wb.create_sheet("Test").unwrap();

        // Write header
        wb.append_row(&mut sheet, vec![
            CellValue::String(Arc::from("Name")),
            CellValue::String(Arc::from("Value")),
        ]).unwrap();

        // Write data rows
        for i in 0..100 {
            wb.append_row(&mut sheet, vec![
                CellValue::String(Arc::from(format!("Item {}", i))),
                CellValue::Number(i as f64),
            ]).unwrap();
        }

        wb.close(sheet).unwrap();

        // Verify file exists and can be read
        let loaded = crate::Workbook::load(path).unwrap();
        let ws = loaded.get_sheet_by_name("Test").unwrap();
        assert_eq!(ws.get_cell_value(1, 1), Some(&CellValue::String(Arc::from("Name"))));
        assert_eq!(ws.get_cell_value(101, 2), Some(&CellValue::Number(99.0)));
    }
}

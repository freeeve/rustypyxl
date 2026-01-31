//! Workbook representation and file I/O operations.

#[cfg(feature = "fast-hash")]
use hashbrown::HashMap;
#[cfg(not(feature = "fast-hash"))]
use std::collections::HashMap;
use std::fs::File;
use std::io::{BufRead, BufReader, Cursor, Read, Seek};
use std::sync::Arc;
use zip::ZipArchive;
use quick_xml::events::Event;
use quick_xml::Reader;
use rayon::prelude::*;

use crate::cell::CellValue;
use crate::error::{Result, RustypyxlError};
use crate::style::{Alignment, Border, BorderStyle, CellStyle, CellXf, Fill, Font, Protection, StyleRegistry};
use crate::utils::{parse_coordinate, parse_coordinate_bytes, parse_u32_bytes, parse_f64_bytes};
use crate::worksheet::{cell_key, CellData, DataValidation, Worksheet, WorksheetProtection};
use crate::writer;

/// A named range definition.
#[derive(Clone, Debug)]
pub struct NamedRange {
    /// Name of the range.
    pub name: String,
    /// Range reference (e.g., "'Sheet1'!A1:B2").
    pub range: String,
}

/// Compression level for saving workbooks.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum CompressionLevel {
    /// No compression - fastest saves, largest files
    None,
    /// Fast compression (deflate level 1) - good balance
    Fast,
    /// Default compression (deflate level 6) - smaller files, slower
    Default,
    /// Best compression (deflate level 9) - smallest files, slowest
    Best,
}

impl std::default::Default for CompressionLevel {
    fn default() -> Self {
        CompressionLevel::None  // Default to fastest for benchmarking
    }
}

/// An Excel workbook containing worksheets.
pub struct Workbook {
    /// List of worksheets.
    pub worksheets: Vec<Worksheet>,
    /// Sheet names (parallel to worksheets).
    pub sheet_names: Vec<String>,
    /// Named ranges defined in the workbook.
    pub named_ranges: Vec<NamedRange>,
    /// Compression level for saving.
    pub compression: CompressionLevel,
    /// Style registry for fonts, fills, borders, number formats, and cell formats.
    pub styles: StyleRegistry,
}

impl Workbook {
    /// Create a new empty workbook.
    pub fn new() -> Self {
        Workbook {
            worksheets: Vec::new(),
            sheet_names: Vec::new(),
            named_ranges: Vec::new(),
            compression: CompressionLevel::default(),
            styles: StyleRegistry::new(),
        }
    }

    /// Set compression level for saving.
    pub fn set_compression(&mut self, level: CompressionLevel) {
        self.compression = level;
    }

    /// Load a workbook from a file path.
    pub fn load(path: &str) -> Result<Self> {
        let file = File::open(path).map_err(|e| {
            RustypyxlError::Io(std::io::Error::new(
                std::io::ErrorKind::NotFound,
                format!("Failed to open file '{}': {}", path, e),
            ))
        })?;

        let mut archive = ZipArchive::new(BufReader::new(file))?;

        let mut workbook = Workbook::new();
        workbook.compression = CompressionLevel::None; // Default to fast for loaded files
        workbook.parse_workbook(&mut archive)?;

        Ok(workbook)
    }

    /// Load a workbook from bytes (e.g., from memory or network).
    pub fn load_from_bytes(data: &[u8]) -> Result<Self> {
        let cursor = Cursor::new(data);
        let mut archive = ZipArchive::new(cursor)?;

        let mut workbook = Workbook::new();
        workbook.compression = CompressionLevel::None;
        workbook.parse_workbook(&mut archive)?;

        Ok(workbook)
    }

    /// Get the active (first) worksheet.
    pub fn active(&self) -> Result<&Worksheet> {
        self.worksheets.first().ok_or(RustypyxlError::NoWorksheets)
    }

    /// Get a mutable reference to the active worksheet.
    pub fn active_mut(&mut self) -> Result<&mut Worksheet> {
        self.worksheets.first_mut().ok_or(RustypyxlError::NoWorksheets)
    }

    /// Get all worksheets.
    pub fn worksheets(&self) -> &[Worksheet] {
        &self.worksheets
    }

    /// Get all sheet names.
    pub fn sheet_names(&self) -> &[String] {
        &self.sheet_names
    }

    /// Get a worksheet by name.
    pub fn get_sheet_by_name(&self, name: &str) -> Result<&Worksheet> {
        for (idx, sheet_name) in self.sheet_names.iter().enumerate() {
            if sheet_name == name {
                return Ok(&self.worksheets[idx]);
            }
        }
        Err(RustypyxlError::WorksheetNotFound(name.to_string()))
    }

    /// Get a mutable worksheet by name.
    pub fn get_sheet_by_name_mut(&mut self, name: &str) -> Result<&mut Worksheet> {
        for (idx, sheet_name) in self.sheet_names.iter().enumerate() {
            if sheet_name == name {
                return Ok(&mut self.worksheets[idx]);
            }
        }
        Err(RustypyxlError::WorksheetNotFound(name.to_string()))
    }

    /// Get a worksheet by index.
    pub fn get_sheet_by_index(&self, index: usize) -> Result<&Worksheet> {
        self.worksheets.get(index).ok_or_else(|| {
            RustypyxlError::WorksheetNotFound(format!("index {}", index))
        })
    }

    /// Get a mutable worksheet by index.
    pub fn get_sheet_by_index_mut(&mut self, index: usize) -> Result<&mut Worksheet> {
        self.worksheets.get_mut(index).ok_or_else(|| {
            RustypyxlError::WorksheetNotFound(format!("index {}", index))
        })
    }

    /// Create a new worksheet.
    pub fn create_sheet(&mut self, title: Option<String>) -> Result<&mut Worksheet> {
        let sheet_title = title.unwrap_or_else(|| format!("Sheet{}", self.worksheets.len() + 1));

        if self.sheet_names.contains(&sheet_title) {
            return Err(RustypyxlError::WorksheetAlreadyExists(sheet_title));
        }

        let worksheet = Worksheet::new(sheet_title.clone());
        self.worksheets.push(worksheet);
        self.sheet_names.push(sheet_title);

        Ok(self.worksheets.last_mut().unwrap())
    }

    /// Remove a worksheet by name.
    pub fn remove_sheet(&mut self, sheet_name: &str) -> Result<()> {
        for (idx, name) in self.sheet_names.iter().enumerate() {
            if name == sheet_name {
                self.worksheets.remove(idx);
                self.sheet_names.remove(idx);
                return Ok(());
            }
        }
        Err(RustypyxlError::WorksheetNotFound(sheet_name.to_string()))
    }

    /// Set a cell value in the active worksheet.
    pub fn set_cell_value(&mut self, row: u32, column: u32, value: CellValue) -> Result<()> {
        let ws = self.active_mut()?;
        ws.set_cell_value(row, column, value);
        Ok(())
    }

    /// Set a cell value in a specific worksheet.
    pub fn set_cell_value_in_sheet(
        &mut self,
        sheet_name: &str,
        row: u32,
        column: u32,
        value: CellValue,
    ) -> Result<()> {
        let ws = self.get_sheet_by_name_mut(sheet_name)?;
        ws.set_cell_value(row, column, value);
        Ok(())
    }

    /// Set cell style in the active worksheet.
    pub fn set_cell_style(&mut self, row: u32, column: u32, style: CellStyle) -> Result<()> {
        let ws = self.active_mut()?;
        ws.set_cell_style(row, column, style);
        Ok(())
    }

    /// Set cell font in the active worksheet.
    pub fn set_cell_font(&mut self, row: u32, column: u32, font: Font) -> Result<()> {
        let ws = self.active_mut()?;
        let cell = ws.cells.entry(cell_key(row, column)).or_insert_with(CellData::new);
        let mut new_style = cell.style.as_ref()
            .map(|s| (**s).clone())
            .unwrap_or_else(CellStyle::new);
        new_style.font = Some(font);
        cell.style = Some(Arc::new(new_style));
        Ok(())
    }

    /// Set cell alignment in the active worksheet.
    pub fn set_cell_alignment(&mut self, row: u32, column: u32, alignment: Alignment) -> Result<()> {
        let ws = self.active_mut()?;
        let cell = ws.cells.entry(cell_key(row, column)).or_insert_with(CellData::new);
        let mut new_style = cell.style.as_ref()
            .map(|s| (**s).clone())
            .unwrap_or_else(CellStyle::new);
        new_style.alignment = Some(alignment);
        cell.style = Some(Arc::new(new_style));
        Ok(())
    }

    /// Set cell number format in the active worksheet.
    pub fn set_cell_number_format(&mut self, row: u32, column: u32, format: String) -> Result<()> {
        let ws = self.active_mut()?;
        ws.set_cell_number_format(row, column, format);
        Ok(())
    }

    /// Set a cell formula in the active worksheet.
    pub fn set_cell_formula(&mut self, row: u32, column: u32, formula: String) -> Result<()> {
        let ws = self.active_mut()?;
        ws.set_cell_formula(row, column, formula);
        Ok(())
    }

    /// Set a cell hyperlink in the active worksheet.
    pub fn set_cell_hyperlink(&mut self, row: u32, column: u32, url: String) -> Result<()> {
        let ws = self.active_mut()?;
        ws.set_cell_hyperlink(row, column, url);
        Ok(())
    }

    /// Set a cell comment in the active worksheet.
    pub fn set_cell_comment(&mut self, row: u32, column: u32, comment: String) -> Result<()> {
        let ws = self.active_mut()?;
        ws.set_cell_comment(row, column, comment);
        Ok(())
    }

    /// Enable protection on the active worksheet.
    pub fn enable_protection(&mut self, password: Option<String>) -> Result<()> {
        let ws = self.active_mut()?;
        ws.enable_protection(password);
        Ok(())
    }

    /// Disable protection on the active worksheet.
    pub fn disable_protection(&mut self) -> Result<()> {
        let ws = self.active_mut()?;
        ws.disable_protection();
        Ok(())
    }

    /// Check if active worksheet is protected.
    pub fn is_protected(&self) -> Result<bool> {
        Ok(self.active()?.is_protected())
    }

    /// Add data validation to a cell in the active worksheet.
    pub fn add_data_validation(
        &mut self,
        row: u32,
        column: u32,
        validation_type: String,
        formula1: Option<String>,
        formula2: Option<String>,
    ) -> Result<()> {
        let ws = self.active_mut()?;
        let validation = DataValidation {
            validation_type,
            formula1,
            formula2,
            allow_blank: true,
            show_error: true,
            error_title: None,
            error_message: None,
            show_input: true,
            prompt_title: None,
            prompt_message: None,
        };
        ws.add_data_validation(row, column, validation);
        Ok(())
    }

    /// Create a named range.
    pub fn create_named_range(&mut self, name: String, range: String) -> Result<()> {
        if self.named_ranges.iter().any(|nr| nr.name == name) {
            return Err(RustypyxlError::NamedRangeAlreadyExists(name));
        }
        self.named_ranges.push(NamedRange { name, range });
        Ok(())
    }

    /// Get a named range by name.
    pub fn get_named_range(&self, name: &str) -> Option<&str> {
        self.named_ranges
            .iter()
            .find(|nr| nr.name == name)
            .map(|nr| nr.range.as_str())
    }

    /// Get all named ranges.
    pub fn get_named_ranges(&self) -> Vec<(&str, &str)> {
        self.named_ranges
            .iter()
            .map(|nr| (nr.name.as_str(), nr.range.as_str()))
            .collect()
    }

    /// Save the workbook to a file.
    pub fn save(&self, path: &str) -> Result<()> {
        let file = File::create(path)?;
        self.save_to_writer(file)
    }

    /// Save the workbook to an in-memory byte vector.
    pub fn save_to_bytes(&self) -> Result<Vec<u8>> {
        let buffer = Cursor::new(Vec::new());
        let mut zip = self.create_zip_writer(buffer)?;
        self.write_workbook_contents(&mut zip)?;
        let cursor = zip.finish()?;
        Ok(cursor.into_inner())
    }

    /// Save the workbook to any writer that implements Write + Seek.
    pub fn save_to_writer<W: std::io::Write + Seek>(&self, writer: W) -> Result<()> {
        let mut zip = self.create_zip_writer(writer)?;
        self.write_workbook_contents(&mut zip)?;
        zip.finish()?;
        Ok(())
    }

    /// Create a ZipWriter with the configured compression options.
    fn create_zip_writer<W: std::io::Write + Seek>(&self, writer: W) -> Result<zip::ZipWriter<W>> {
        Ok(zip::ZipWriter::new(writer))
    }

    /// Get the file options based on compression settings.
    fn get_file_options(&self) -> zip::write::FileOptions<'static, zip::write::ExtendedFileOptions> {
        use zip::write::FileOptions;
        use zip::CompressionMethod;

        match self.compression {
            CompressionLevel::None => FileOptions::default()
                .large_file(false)
                .compression_method(CompressionMethod::Stored),
            CompressionLevel::Fast => FileOptions::default()
                .large_file(false)
                .compression_method(CompressionMethod::Deflated)
                .compression_level(Some(1)),
            CompressionLevel::Default => FileOptions::default()
                .large_file(false)
                .compression_method(CompressionMethod::Deflated)
                .compression_level(Some(6)),
            CompressionLevel::Best => FileOptions::default()
                .large_file(false)
                .compression_method(CompressionMethod::Deflated)
                .compression_level(Some(9)),
        }
    }

    /// Write all workbook contents to a ZipWriter.
    fn write_workbook_contents<W: std::io::Write + Seek>(&self, zip: &mut zip::ZipWriter<W>) -> Result<()> {
        use std::io::Write;
        use zip::write::FileOptions;

        let options = self.get_file_options();

        // Collect shared strings first to know if we have any
        let (shared_strings_vec, shared_strings_map) = writer::collect_shared_strings(&self.worksheets);
        let has_shared_strings = !shared_strings_vec.is_empty();

        // Write [Content_Types].xml
        writer::write_content_types(zip, &options, self.worksheets.len(), has_shared_strings)?;

        // Write _rels/.rels
        writer::write_rels(zip, &options)?;

        // Write docProps files
        writer::write_doc_props(zip, &options)?;

        // Write xl/workbook.xml
        let named_ranges: Vec<(String, String)> = self
            .named_ranges
            .iter()
            .map(|nr| (nr.name.clone(), nr.range.clone()))
            .collect();
        writer::write_workbook_xml(zip, &options, &self.sheet_names, &named_ranges)?;

        // Write xl/_rels/workbook.xml.rels
        writer::write_workbook_rels(zip, &options, self.worksheets.len(), has_shared_strings)?;

        // Write shared strings if we have any
        if has_shared_strings {
            writer::write_shared_strings(zip, &options, &shared_strings_vec)?;
        }

        // Write styles.xml
        writer::write_styles_xml(zip, &options, &self.styles)?;

        // Write each worksheet and comments
        for (idx, worksheet) in self.worksheets.iter().enumerate() {
            let sheet_id = (idx + 1) as u32;

            // Check if worksheet has comments
            let has_comments = worksheet.cells.values().any(|cd| cd.comment.is_some());

            writer::write_worksheet_xml(
                zip,
                &options,
                worksheet,
                sheet_id,
                &shared_strings_map,
                has_comments,
            )?;

            // Write comments if any exist
            if has_comments {
                writer::write_comments_xml(zip, &options, worksheet, sheet_id)?;

                // Write worksheet relationships for comments
                let rels_path = format!("xl/worksheets/_rels/sheet{}.xml.rels", sheet_id);
                let rels_options: zip::write::FileOptions<'static, zip::write::ExtendedFileOptions> =
                    FileOptions::default().compression_method(zip::CompressionMethod::Deflated);
                zip.start_file(&rels_path, rels_options)?;
                let rels_content = format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment{}.xml" Id="comments" />
</Relationships>"#,
                    sheet_id
                );
                zip.write_all(rels_content.as_bytes())?;
            }
        }

        Ok(())
    }

    /// Parse workbook from ZIP archive with parallel worksheet parsing.
    fn parse_workbook<R: Read + Seek>(&mut self, archive: &mut ZipArchive<R>) -> Result<()> {
        // Phase 1: Load all file contents into memory (sequential ZIP extraction)
        let workbook_xml = Self::read_zip_file_to_vec(archive, "xl/workbook.xml")?;
        let workbook_rels_xml = Self::read_zip_file_to_vec(archive, "xl/_rels/workbook.xml.rels").ok();
        let shared_strings_xml = Self::read_zip_file_to_vec(archive, "xl/sharedStrings.xml").ok();
        let styles_xml = Self::read_zip_file_to_vec(archive, "xl/styles.xml").ok();

        // Parse workbook.xml to get sheet names, IDs, and relationship IDs
        let (sheet_info, named_ranges) =
            Self::parse_workbook_xml(Cursor::new(&workbook_xml))?;
        self.named_ranges = named_ranges;

        // Parse workbook.xml.rels to get the mapping from rId to actual file paths
        let rels_map: HashMap<String, String> = if let Some(rels_xml) = workbook_rels_xml {
            Self::parse_workbook_rels(Cursor::new(&rels_xml))?
        } else {
            HashMap::new()
        };

        // Load all worksheet and comments XML into memory
        let mut sheet_data: Vec<(String, u32, Vec<u8>, Option<Vec<u8>>)> = Vec::with_capacity(sheet_info.len());
        for (sheet_name, sheet_id, sheet_rid) in &sheet_info {
            // Look up the actual sheet path from the relationships, or fall back to sheetId-based path
            let sheet_path = if let Some(target) = rels_map.get(sheet_rid) {
                // Target is relative to xl/, e.g., "worksheets/sheet1.xml"
                if target.starts_with('/') {
                    // Absolute path within the package (rare)
                    target[1..].to_string()
                } else {
                    format!("xl/{}", target)
                }
            } else {
                // Fallback to legacy behavior if rels file is missing or incomplete
                format!("xl/worksheets/sheet{}.xml", sheet_id)
            };
            let sheet_xml = Self::read_zip_file_to_vec(archive, &sheet_path)?;

            let comments_path = format!("xl/comments/comment{}.xml", sheet_id);
            let comments_xml = Self::read_zip_file_to_vec(archive, &comments_path).ok();

            sheet_data.push((sheet_name.clone(), *sheet_id, sheet_xml, comments_xml));
        }

        // Phase 2: Parse shared data (must be done before worksheets)
        let shared_strings = if let Some(xml) = shared_strings_xml {
            Self::parse_shared_strings_xml(Cursor::new(&xml))?
        } else {
            Vec::new()
        };

        let (styles, style_registry) = if let Some(xml) = styles_xml {
            Self::parse_styles_xml(&xml)?
        } else {
            (HashMap::new(), StyleRegistry::new())
        };

        // Phase 3: Parse worksheets in parallel using Rayon
        let shared_strings_ref = &shared_strings;
        let styles_ref = &styles;

        let worksheets: Vec<Result<(String, Worksheet)>> = if sheet_data.len() > 1 {
            // Parallel parsing for multiple sheets
            sheet_data
                .par_iter()
                .map(|(sheet_name, _sheet_id, sheet_xml, comments_xml)| {
                    let mut worksheet = Worksheet::new(sheet_name.clone());
                    Self::parse_worksheet_xml(
                        Cursor::new(sheet_xml),
                        shared_strings_ref,
                        styles_ref,
                        &mut worksheet,
                    )?;

                    if let Some(comments) = comments_xml {
                        Self::parse_comments_xml(Cursor::new(comments), &mut worksheet)?;
                    }

                    Ok((sheet_name.clone(), worksheet))
                })
                .collect()
        } else {
            // Sequential for single sheet (avoid Rayon overhead)
            sheet_data
                .iter()
                .map(|(sheet_name, _sheet_id, sheet_xml, comments_xml)| {
                    let mut worksheet = Worksheet::new(sheet_name.clone());
                    Self::parse_worksheet_xml(
                        Cursor::new(sheet_xml),
                        shared_strings_ref,
                        styles_ref,
                        &mut worksheet,
                    )?;

                    if let Some(comments) = comments_xml {
                        Self::parse_comments_xml(Cursor::new(comments), &mut worksheet)?;
                    }

                    Ok((sheet_name.clone(), worksheet))
                })
                .collect()
        };

        // Collect results in order
        for result in worksheets {
            let (sheet_name, worksheet) = result?;
            self.worksheets.push(worksheet);
            self.sheet_names.push(sheet_name);
        }

        // Store the style registry
        self.styles = style_registry;

        Ok(())
    }

    /// Read a file from the ZIP archive into a Vec<u8>.
    fn read_zip_file_to_vec<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
        path: &str,
    ) -> Result<Vec<u8>> {
        let mut file = archive.by_name(path).map_err(|e| {
            RustypyxlError::InvalidFormat(format!("Failed to find {} in archive: {}", path, e))
        })?;
        let mut buf = Vec::with_capacity(file.size() as usize);
        file.read_to_end(&mut buf)?;
        Ok(buf)
    }

    /// Parses workbook.xml and returns sheet info (name, sheetId, rId) and named ranges.
    fn parse_workbook_xml<R: BufRead>(
        reader: R,
    ) -> Result<(Vec<(String, u32, String)>, Vec<NamedRange>)> {
        let mut reader = Reader::from_reader(reader);
        reader.config_mut().trim_text(true);

        let mut sheets = Vec::new();
        let mut named_ranges = Vec::new();
        let mut buf = Vec::new();
        let mut current_sheet_name: Option<String> = None;
        let mut current_sheet_id: Option<u32> = None;
        let mut current_sheet_rid: Option<String> = None;
        let mut in_defined_names = false;
        let mut current_name: Option<String> = None;
        let mut current_range: Option<String> = None;
        let mut in_defined_name = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) => {
                    let name = e.name();
                    let local = e.local_name();
                    let name = name.as_ref();
                    let local = local.as_ref();

                    // Handle self-closing sheet tags
                    if name == b"sheet" || local == b"sheet" {
                        let mut sheet_name: Option<String> = None;
                        let mut sheet_id: Option<u32> = None;
                        let mut sheet_rid: Option<String> = None;

                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key;
                                let attr_local = attr.key.local_name();
                                let attr_key = attr_key.as_ref();
                                let attr_local = attr_local.as_ref();

                                if attr_key == b"name" || attr_local == b"name" {
                                    sheet_name =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                } else if attr_key == b"sheetId" || attr_local == b"sheetId" {
                                    let id_str = String::from_utf8_lossy(&attr.value);
                                    sheet_id = id_str.parse().ok();
                                } else if attr_local == b"id" {
                                    // r:id attribute (namespace-qualified)
                                    sheet_rid =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }

                        if let (Some(name), Some(id), Some(rid)) = (sheet_name, sheet_id, sheet_rid) {
                            sheets.push((name, id, rid));
                        }
                    }
                }
                Ok(Event::Start(e)) => {
                    let name = e.name();
                    let local = e.local_name();
                    let name = name.as_ref();
                    let local = local.as_ref();
                    let is_sheet = name == b"sheet" || local == b"sheet";
                    let is_defined_names = name == b"definedNames" || local == b"definedNames";
                    let is_defined_name = name == b"definedName" || local == b"definedName";

                    if is_defined_names {
                        in_defined_names = true;
                    } else if is_defined_name && in_defined_names {
                        in_defined_name = true;
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key;
                                let attr_local = attr.key.local_name();
                                let attr_key = attr_key.as_ref();
                                let attr_local = attr_local.as_ref();
                                if attr_key == b"name" || attr_local == b"name" {
                                    current_name =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                    } else if is_sheet {
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key;
                                let attr_local = attr.key.local_name();
                                let attr_key = attr_key.as_ref();
                                let attr_local = attr_local.as_ref();

                                if attr_key == b"name" || attr_local == b"name" {
                                    current_sheet_name =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                } else if attr_key == b"sheetId" || attr_local == b"sheetId" {
                                    let id_str = String::from_utf8_lossy(&attr.value);
                                    current_sheet_id = id_str.parse().ok();
                                } else if attr_local == b"id" {
                                    // r:id attribute (namespace-qualified)
                                    current_sheet_rid =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                    }
                }
                Ok(Event::Text(e)) => {
                    if in_defined_name && in_defined_names {
                        let text = e.unescape().unwrap_or_default();
                        current_range = Some(text.to_string());
                    }
                }
                Ok(Event::End(e)) => {
                    let name = e.name();
                    let local = e.local_name();
                    let name = name.as_ref();
                    let local = local.as_ref();
                    let is_sheet = name == b"sheet" || local == b"sheet";
                    let is_defined_names = name == b"definedNames" || local == b"definedNames";
                    let is_defined_name = name == b"definedName" || local == b"definedName";

                    if is_defined_name && in_defined_name {
                        if let (Some(name), Some(range)) = (current_name.take(), current_range.take())
                        {
                            named_ranges.push(NamedRange { name, range });
                        }
                        in_defined_name = false;
                    } else if is_defined_names {
                        in_defined_names = false;
                    } else if is_sheet {
                        if let (Some(name), Some(id), Some(rid)) =
                            (current_sheet_name.take(), current_sheet_id.take(), current_sheet_rid.take())
                        {
                            sheets.push((name, id, rid));
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(RustypyxlError::ParseError(format!(
                        "XML parsing error: {}",
                        e
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok((sheets, named_ranges))
    }

    /// Parses workbook.xml.rels and returns a mapping of relationship IDs to target paths.
    fn parse_workbook_rels<R: BufRead>(reader: R) -> Result<HashMap<String, String>> {
        let mut reader = Reader::from_reader(reader);
        reader.config_mut().trim_text(true);

        let mut rels = HashMap::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                    let name = e.name();
                    let local = e.local_name();
                    let name = name.as_ref();
                    let local = local.as_ref();

                    if name == b"Relationship" || local == b"Relationship" {
                        let mut rel_id: Option<String> = None;
                        let mut target: Option<String> = None;

                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                if attr_key == b"Id" {
                                    rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                                } else if attr_key == b"Target" {
                                    target = Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }

                        if let (Some(id), Some(tgt)) = (rel_id, target) {
                            rels.insert(id, tgt);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(RustypyxlError::ParseError(format!(
                        "XML parsing error in workbook.xml.rels: {}",
                        e
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(rels)
    }

    fn parse_shared_strings_xml<R: BufRead>(reader: R) -> Result<Vec<crate::cell::InternedString>> {
        let mut reader = Reader::from_reader(reader);
        // Don't trim text - we need to preserve whitespace in string values
        reader.config_mut().trim_text(false);

        let mut strings = Vec::new();
        let mut buf = Vec::new();
        let mut current_string = String::new();
        let mut in_t = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.name().as_ref() == b"t" {
                        in_t = true;
                    }
                }
                Ok(Event::Text(e)) => {
                    if in_t {
                        current_string.push_str(&e.unescape().unwrap_or_default());
                    }
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"t" {
                        in_t = false;
                    } else if e.name().as_ref() == b"si" {
                        strings.push(std::sync::Arc::from(current_string.as_str()));
                        current_string.clear();
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(RustypyxlError::ParseError(format!(
                        "XML parsing error: {}",
                        e
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(strings)
    }

    /// Get a string attribute value from an XML element.
    fn get_attr_str(e: &quick_xml::events::BytesStart, key: &[u8]) -> Option<String> {
        for attr in e.attributes().flatten() {
            if attr.key.as_ref() == key {
                return Some(String::from_utf8_lossy(&attr.value).to_string());
            }
        }
        None
    }

    /// Get an optional u32 attribute value from an XML element.
    fn get_attr_u32(e: &quick_xml::events::BytesStart, key: &[u8]) -> Option<u32> {
        Self::get_attr_str(e, key).and_then(|s| s.parse().ok())
    }

    /// Get an optional f64 attribute value from an XML element.
    fn get_attr_f64(e: &quick_xml::events::BytesStart, key: &[u8]) -> Option<f64> {
        Self::get_attr_str(e, key).and_then(|s| s.parse().ok())
    }

    /// Check if an attribute equals "1" or "true".
    fn get_attr_bool(e: &quick_xml::events::BytesStart, key: &[u8]) -> bool {
        Self::get_attr_str(e, key)
            .map(|s| s == "1" || s == "true")
            .unwrap_or(false)
    }

    /// Parse font properties from an XML element (handles both Start and Empty events).
    fn parse_font_element(e: &quick_xml::events::BytesStart, font: &mut Font) {
        let name = e.name();
        let name = name.as_ref();
        match name {
            b"b" => font.bold = true,
            b"i" => font.italic = true,
            b"u" => font.underline = true,
            b"strike" => font.strike = true,
            b"sz" => font.size = Self::get_attr_f64(e, b"val"),
            b"name" => font.name = Self::get_attr_str(e, b"val"),
            b"vertAlign" => font.vert_align = Self::get_attr_str(e, b"val"),
            b"color" => {
                if let Some(rgb) = Self::get_attr_str(e, b"rgb") {
                    font.color = Some(format!("#{}", rgb));
                } else if let Some(theme) = Self::get_attr_str(e, b"theme") {
                    font.color = Some(format!("theme:{}", theme));
                }
            }
            _ => {}
        }
    }

    /// Parse fill properties from an XML element.
    fn parse_fill_element(e: &quick_xml::events::BytesStart, fill: &mut Fill) {
        let name = e.name();
        let name = name.as_ref();
        match name {
            b"patternFill" => {
                fill.pattern_type = Self::get_attr_str(e, b"patternType");
            }
            b"fgColor" => {
                if let Some(rgb) = Self::get_attr_str(e, b"rgb") {
                    fill.fg_color = Some(format!("#{}", rgb));
                } else if let Some(theme) = Self::get_attr_str(e, b"theme") {
                    fill.fg_color = Some(format!("theme:{}", theme));
                }
            }
            b"bgColor" => {
                if let Some(rgb) = Self::get_attr_str(e, b"rgb") {
                    fill.bg_color = Some(format!("#{}", rgb));
                } else if let Some(theme) = Self::get_attr_str(e, b"theme") {
                    fill.bg_color = Some(format!("theme:{}", theme));
                }
            }
            _ => {}
        }
    }

    /// Parse border side properties and return (style, color).
    fn parse_border_side_attrs(e: &quick_xml::events::BytesStart) -> (Option<String>, Option<String>) {
        let style = Self::get_attr_str(e, b"style");
        let color = None; // Color comes from nested element
        (style, color)
    }

    /// Parse a color element and return the color string.
    fn parse_color_element(e: &quick_xml::events::BytesStart) -> Option<String> {
        if let Some(rgb) = Self::get_attr_str(e, b"rgb") {
            Some(format!("#{}", rgb))
        } else if let Some(theme) = Self::get_attr_str(e, b"theme") {
            Some(format!("theme:{}", theme))
        } else {
            None
        }
    }

    fn parse_styles_xml(xml: &[u8]) -> Result<(HashMap<u32, Arc<CellStyle>>, StyleRegistry)> {
        let mut reader = Reader::from_reader(Cursor::new(xml));
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut fonts: Vec<Font> = Vec::new();
        let mut fills: Vec<Fill> = Vec::new();
        let mut borders: Vec<Border> = Vec::new();
        let mut number_formats: HashMap<u32, String> = HashMap::new();
        let mut cell_styles: HashMap<u32, Arc<CellStyle>> = HashMap::new();

        let mut in_font = false;
        let mut in_fill = false;
        let mut in_border = false;
        let mut _in_num_fmt = false;
        let mut in_border_side: Option<&'static str> = None; // "left", "right", "top", "bottom", "diagonal"

        let mut current_font = Font::default();
        let mut current_fill = Fill::default();
        let mut current_border = Border::default();
        let mut current_border_style: Option<String> = None;
        let mut current_border_color: Option<String> = None;
        let mut current_num_fmt_id: Option<u32> = None;
        let mut current_num_fmt_code: Option<String> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) => {
                    let name = e.name();
                    let name = name.as_ref();

                    // Handle font properties
                    if in_font {
                        Self::parse_font_element(&e, &mut current_font);
                    }

                    // Handle fill properties
                    if in_fill {
                        Self::parse_fill_element(&e, &mut current_fill);
                    }
                    // Handle self-closing border side elements (e.g., <left style="thin"/>)
                    if in_border && (name == b"left" || name == b"right" || name == b"top"
                                     || name == b"bottom" || name == b"diagonal") {
                        let mut style: Option<String> = None;
                        let color: Option<String> = None;
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                if attr.key.as_ref() == b"style" {
                                    style = Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                        if let Some(s) = style {
                            let border_style = BorderStyle { style: s, color };
                            match name {
                                b"left" => current_border.left = Some(border_style),
                                b"right" => current_border.right = Some(border_style),
                                b"top" => current_border.top = Some(border_style),
                                b"bottom" => current_border.bottom = Some(border_style),
                                b"diagonal" => current_border.diagonal = Some(border_style),
                                _ => {}
                            }
                        }
                    }
                    // Handle color inside border side (self-closing)
                    if in_border && in_border_side.is_some() && name == b"color" {
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                if attr.key.as_ref() == b"rgb" {
                                    current_border_color = Some(format!(
                                        "#{}",
                                        String::from_utf8_lossy(&attr.value)
                                    ));
                                }
                            }
                        }
                    }
                    // Handle numFmt as empty element (self-closing)
                    if name == b"numFmt" {
                        let mut fmt_id: Option<u32> = None;
                        let mut fmt_code: Option<String> = None;
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                if attr_key == b"numFmtId" {
                                    if let Ok(id) = String::from_utf8_lossy(&attr.value).parse::<u32>() {
                                        fmt_id = Some(id);
                                    }
                                } else if attr_key == b"formatCode" {
                                    fmt_code = Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                        if let (Some(id), Some(code)) = (fmt_id, fmt_code) {
                            number_formats.insert(id, code);
                        }
                    }
                }
                Ok(Event::Start(e)) => {
                    let name = e.name();
                    let name = name.as_ref();

                    if name == b"font" {
                        in_font = true;
                        current_font = Font::default();
                    } else if name == b"fill" {
                        in_fill = true;
                        current_fill = Fill::default();
                    } else if name == b"border" {
                        in_border = true;
                        current_border = Border::default();
                    } else if name == b"numFmt" {
                        _in_num_fmt = true;
                        current_num_fmt_id = None;
                        current_num_fmt_code = None;
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                if attr_key == b"numFmtId" {
                                    if let Ok(id) =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>()
                                    {
                                        current_num_fmt_id = Some(id);
                                    }
                                } else if attr_key == b"formatCode" {
                                    current_num_fmt_code =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                    } else if in_font {
                        Self::parse_font_element(&e, &mut current_font);
                    } else if in_fill {
                        Self::parse_fill_element(&e, &mut current_fill);
                    } else if in_border {
                        let prop_name = e.name();
                        let prop_name = prop_name.as_ref();
                        // Handle border side start elements
                        if prop_name == b"left" || prop_name == b"right" || prop_name == b"top"
                           || prop_name == b"bottom" || prop_name == b"diagonal" {
                            in_border_side = Some(match prop_name {
                                b"left" => "left",
                                b"right" => "right",
                                b"top" => "top",
                                b"bottom" => "bottom",
                                b"diagonal" => "diagonal",
                                _ => "left",
                            });
                            current_border_style = None;
                            current_border_color = None;
                            // Get style attribute
                            for attr in e.attributes() {
                                if let Ok(attr) = attr {
                                    if attr.key.as_ref() == b"style" {
                                        current_border_style = Some(
                                            String::from_utf8_lossy(&attr.value).to_string()
                                        );
                                    }
                                }
                            }
                        } else if prop_name == b"color" && in_border_side.is_some() {
                            // Get color for current border side
                            for attr in e.attributes() {
                                if let Ok(attr) = attr {
                                    if attr.key.as_ref() == b"rgb" {
                                        current_border_color = Some(format!(
                                            "#{}",
                                            String::from_utf8_lossy(&attr.value)
                                        ));
                                    }
                                }
                            }
                        }
                    }
                }
                Ok(Event::End(e)) => {
                    let name = e.name();
                    let name = name.as_ref();

                    if name == b"font" {
                        fonts.push(current_font.clone());
                        in_font = false;
                    } else if name == b"fill" {
                        fills.push(current_fill.clone());
                        in_fill = false;
                    } else if name == b"border" {
                        borders.push(current_border.clone());
                        in_border = false;
                    } else if in_border && (name == b"left" || name == b"right" || name == b"top"
                                            || name == b"bottom" || name == b"diagonal") {
                        // Finalize border side
                        if let Some(style) = current_border_style.take() {
                            let border_style = BorderStyle {
                                style,
                                color: current_border_color.take(),
                            };
                            match name {
                                b"left" => current_border.left = Some(border_style),
                                b"right" => current_border.right = Some(border_style),
                                b"top" => current_border.top = Some(border_style),
                                b"bottom" => current_border.bottom = Some(border_style),
                                b"diagonal" => current_border.diagonal = Some(border_style),
                                _ => {}
                            }
                        }
                        in_border_side = None;
                    } else if name == b"numFmt" {
                        if let (Some(id), Some(code)) =
                            (current_num_fmt_id, current_num_fmt_code.take())
                        {
                            number_formats.insert(id, code);
                        }
                        _in_num_fmt = false;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(RustypyxlError::ParseError(format!(
                        "XML parsing error: {}",
                        e
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        // Re-parse to build cellXfs mapping
        let mut reader2 = Reader::from_reader(Cursor::new(xml));
        reader2.config_mut().trim_text(true);
        let mut buf2 = Vec::new();
        let mut xf_index = 0u32;
        let mut current_xf = CellStyle::default();
        let mut in_cell_xfs = false;
        let mut in_xf = false;
        let mut has_alignment = false;
        let mut current_align = Alignment::default();
        let mut has_protection = false;
        let mut current_protection = Protection::default();

        loop {
            match reader2.read_event_into(&mut buf2) {
                Ok(Event::Start(e)) => {
                    let name = e.name();
                    let name = name.as_ref();
                    if name == b"cellXfs" {
                        in_cell_xfs = true;
                        xf_index = 0;
                    } else if name == b"xf" && in_cell_xfs {
                        in_xf = true;
                        current_xf = CellStyle::default();
                        current_align = Alignment::default();

                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                if attr_key == b"fontId" {
                                    if let Ok(id) =
                                        String::from_utf8_lossy(&attr.value).parse::<usize>()
                                    {
                                        if id < fonts.len() {
                                            current_xf.font = Some(fonts[id].clone());
                                        }
                                    }
                                } else if attr_key == b"fillId" {
                                    if let Ok(id) =
                                        String::from_utf8_lossy(&attr.value).parse::<usize>()
                                    {
                                        if id < fills.len() {
                                            current_xf.fill = Some(fills[id].clone());
                                        }
                                    }
                                } else if attr_key == b"borderId" {
                                    if let Ok(id) =
                                        String::from_utf8_lossy(&attr.value).parse::<usize>()
                                    {
                                        if id < borders.len() {
                                            current_xf.border = Some(borders[id].clone());
                                        }
                                    }
                                } else if attr_key == b"numFmtId" {
                                    if let Ok(id) =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>()
                                    {
                                        if let Some(format) = number_formats.get(&id) {
                                            current_xf.number_format = Some(format.clone());
                                        } else {
                                            let builtin_format = match id {
                                                0 => Some("General".to_string()),
                                                1 => Some("0".to_string()),
                                                2 => Some("0.00".to_string()),
                                                3 => Some("#,##0".to_string()),
                                                4 => Some("#,##0.00".to_string()),
                                                9 => Some("0%".to_string()),
                                                10 => Some("0.00%".to_string()),
                                                11 => Some("0.00E+00".to_string()),
                                                14 => Some("mm/dd/yyyy".to_string()),
                                                22 => Some("m/d/yy h:mm".to_string()),
                                                _ => None,
                                            };
                                            if let Some(format) = builtin_format {
                                                current_xf.number_format = Some(format);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    } else if name == b"alignment" && in_xf {
                        has_alignment = true;
                        current_align = Alignment::default();
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                if attr_key == b"horizontal" {
                                    current_align.horizontal =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                } else if attr_key == b"vertical" {
                                    current_align.vertical =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                } else if attr_key == b"wrapText" {
                                    current_align.wrap_text =
                                        String::from_utf8_lossy(&attr.value) == "1";
                                } else if attr_key == b"textRotation" {
                                    if let Ok(rotation) = String::from_utf8_lossy(&attr.value).parse::<i32>() {
                                        current_align.text_rotation = Some(rotation);
                                    }
                                } else if attr_key == b"shrinkToFit" {
                                    current_align.shrink_to_fit =
                                        String::from_utf8_lossy(&attr.value) == "1";
                                } else if attr_key == b"indent" {
                                    if let Ok(indent) = String::from_utf8_lossy(&attr.value).parse::<u32>() {
                                        current_align.indent = Some(indent);
                                    }
                                }
                            }
                        }
                    } else if name == b"protection" && in_xf {
                        has_protection = true;
                        current_protection = Protection::default();
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                if attr_key == b"locked" {
                                    current_protection.locked =
                                        String::from_utf8_lossy(&attr.value) == "1";
                                } else if attr_key == b"hidden" {
                                    current_protection.hidden =
                                        String::from_utf8_lossy(&attr.value) == "1";
                                }
                            }
                        }
                    }
                }
                Ok(Event::End(e)) => {
                    let name = e.name();
                    let name = name.as_ref();
                    if name == b"xf" && in_xf && in_cell_xfs {
                        current_xf.alignment = if has_alignment {
                            Some(current_align.clone())
                        } else {
                            None
                        };
                        current_xf.protection = if has_protection {
                            Some(current_protection.clone())
                        } else {
                            None
                        };
                        cell_styles.insert(xf_index, Arc::new(current_xf.clone()));
                        xf_index += 1;
                        in_xf = false;
                        has_alignment = false;
                        has_protection = false;
                        current_align = Alignment::default();
                        current_protection = Protection::default();
                    } else if name == b"cellXfs" {
                        in_cell_xfs = false;
                    }
                }
                Ok(Event::Empty(e)) => {
                    let name = e.name();
                    let name = name.as_ref();
                    if name == b"alignment" && in_xf && in_cell_xfs {
                        has_alignment = true;
                        current_align = Alignment::default();
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                if attr_key == b"horizontal" {
                                    current_align.horizontal =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                } else if attr_key == b"vertical" {
                                    current_align.vertical =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                } else if attr_key == b"wrapText" {
                                    current_align.wrap_text =
                                        String::from_utf8_lossy(&attr.value) == "1";
                                } else if attr_key == b"textRotation" {
                                    if let Ok(rotation) = String::from_utf8_lossy(&attr.value).parse::<i32>() {
                                        current_align.text_rotation = Some(rotation);
                                    }
                                } else if attr_key == b"shrinkToFit" {
                                    current_align.shrink_to_fit =
                                        String::from_utf8_lossy(&attr.value) == "1";
                                } else if attr_key == b"indent" {
                                    if let Ok(indent) = String::from_utf8_lossy(&attr.value).parse::<u32>() {
                                        current_align.indent = Some(indent);
                                    }
                                }
                            }
                        }
                    } else if name == b"protection" && in_xf && in_cell_xfs {
                        has_protection = true;
                        current_protection = Protection::default();
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                if attr_key == b"locked" {
                                    current_protection.locked =
                                        String::from_utf8_lossy(&attr.value) == "1";
                                } else if attr_key == b"hidden" {
                                    current_protection.hidden =
                                        String::from_utf8_lossy(&attr.value) == "1";
                                }
                            }
                        }
                    } else if name == b"xf" && in_cell_xfs {
                        let mut xf = CellStyle::default();
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                if attr_key == b"fontId" {
                                    if let Ok(id) =
                                        String::from_utf8_lossy(&attr.value).parse::<usize>()
                                    {
                                        if id < fonts.len() {
                                            xf.font = Some(fonts[id].clone());
                                        }
                                    }
                                } else if attr_key == b"fillId" {
                                    if let Ok(id) =
                                        String::from_utf8_lossy(&attr.value).parse::<usize>()
                                    {
                                        if id < fills.len() {
                                            xf.fill = Some(fills[id].clone());
                                        }
                                    }
                                } else if attr_key == b"borderId" {
                                    if let Ok(id) =
                                        String::from_utf8_lossy(&attr.value).parse::<usize>()
                                    {
                                        if id < borders.len() {
                                            xf.border = Some(borders[id].clone());
                                        }
                                    }
                                } else if attr_key == b"numFmtId" {
                                    if let Ok(id) =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>()
                                    {
                                        if let Some(format) = number_formats.get(&id) {
                                            xf.number_format = Some(format.clone());
                                        } else {
                                            let builtin = match id {
                                                0 => Some("General".to_string()),
                                                1 => Some("0".to_string()),
                                                2 => Some("0.00".to_string()),
                                                3 => Some("#,##0".to_string()),
                                                4 => Some("#,##0.00".to_string()),
                                                9 => Some("0%".to_string()),
                                                10 => Some("0.00%".to_string()),
                                                11 => Some("0.00E+00".to_string()),
                                                14 => Some("mm/dd/yyyy".to_string()),
                                                22 => Some("m/d/yy h:mm".to_string()),
                                                _ => None,
                                            };
                                            if let Some(fmt) = builtin {
                                                xf.number_format = Some(fmt);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        cell_styles.insert(xf_index, Arc::new(xf));
                        xf_index += 1;
                    }
                }
                Ok(Event::Eof) => break,
                _ => {}
            }
            buf2.clear();
        }

        // Build StyleRegistry from parsed data
        let mut registry = StyleRegistry::default();

        // Add fonts (ensure at least one default)
        if fonts.is_empty() {
            registry.fonts.push(Font {
                name: Some("Calibri".to_string()),
                size: Some(11.0),
                ..Default::default()
            });
        } else {
            registry.fonts = fonts;
        }

        // Add fills (ensure at least two defaults: none and gray125)
        if fills.is_empty() {
            registry.fills.push(Fill::default());
            registry.fills.push(Fill {
                pattern_type: Some("gray125".to_string()),
                ..Default::default()
            });
        } else {
            registry.fills = fills;
        }

        // Add borders (ensure at least one default)
        if borders.is_empty() {
            registry.borders.push(Border::default());
        } else {
            registry.borders = borders;
        }

        // Add custom number formats
        for (id, code) in number_formats {
            if id >= 164 {
                registry.num_fmts.push((id as usize, code));
            }
        }

        // Build cellXfs from the cell_styles
        // Iterate in order since cell_styles HashMap keys are indices
        let max_xf = cell_styles.keys().copied().max().unwrap_or(0);
        for i in 0..=max_xf {
            if let Some(style) = cell_styles.get(&i) {
                let xf = CellXf {
                    font_id: style.font.as_ref()
                        .and_then(|f| registry.fonts.iter().position(|rf| rf == f))
                        .unwrap_or(0),
                    fill_id: style.fill.as_ref()
                        .and_then(|f| registry.fills.iter().position(|rf| rf == f))
                        .unwrap_or(0),
                    border_id: style.border.as_ref()
                        .and_then(|b| registry.borders.iter().position(|rb| rb == b))
                        .unwrap_or(0),
                    num_fmt_id: style.number_format.as_ref()
                        .and_then(|nf| StyleRegistry::builtin_num_fmt_id(nf))
                        .or_else(|| {
                            style.number_format.as_ref().and_then(|nf| {
                                registry.num_fmts.iter().find(|(_, code)| code == nf).map(|(id, _)| *id)
                            })
                        })
                        .unwrap_or(0),
                    alignment: style.alignment.clone(),
                    protection: style.protection.clone(),
                    apply_font: style.font.is_some(),
                    apply_fill: style.fill.is_some(),
                    apply_border: style.border.is_some(),
                    apply_number_format: style.number_format.is_some(),
                    apply_alignment: style.alignment.is_some(),
                    apply_protection: style.protection.is_some(),
                };
                registry.cell_xfs.push(xf);
            } else {
                // Fill gaps with default xf
                registry.cell_xfs.push(CellXf::default());
            }
        }

        // Ensure at least one cellXf
        if registry.cell_xfs.is_empty() {
            registry.cell_xfs.push(CellXf::default());
        }

        Ok((cell_styles, registry))
    }

    fn estimate_dimension_cells(ref_str: &str) -> Option<usize> {
        let ref_str = ref_str.trim();
        if ref_str.is_empty() {
            return None;
        }

        let (start, end) = if let Some(colon_pos) = ref_str.find(':') {
            let start = parse_coordinate(&ref_str[..colon_pos]).ok()?;
            let end = parse_coordinate(&ref_str[colon_pos + 1..]).ok()?;
            (start, end)
        } else {
            let coord = parse_coordinate(ref_str).ok()?;
            (coord, coord)
        };

        if end.0 < start.0 || end.1 < start.1 {
            return None;
        }

        let rows = (end.0 - start.0 + 1) as u64;
        let cols = (end.1 - start.1 + 1) as u64;
        let cells = rows.saturating_mul(cols);
        let max_reserve = 5_000_000u64;

        if cells == 0 || cells > max_reserve {
            return None;
        }

        Some(cells as usize)
    }

    fn parse_worksheet_xml<R: BufRead>(
        reader: R,
        shared_strings: &[crate::cell::InternedString],
        styles: &HashMap<u32, Arc<CellStyle>>,
        worksheet: &mut Worksheet,
    ) -> Result<()> {
        let mut reader = Reader::from_reader(reader);
        // Don't trim text - we need to preserve whitespace in cell values
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut current_row: Option<u32> = None;
        let mut current_col: Option<u32> = None;
        enum TempValue {
            SharedIdx(usize),
            Bool(bool),
            Number(f64),
            Date(String),
            String(String),
        }

        let mut current_value: Option<TempValue> = None;
        // Cell type as single byte: b's'=shared, b'b'=bool, b'd'=date, b'i'=inline, 0=number
        let mut current_type: u8 = 0;
        let mut current_style_id: Option<u32> = None;
        let mut current_formula: Option<String> = None;
        let mut current_number_format: Option<String> = None;
        let mut in_cell = false;
        let mut in_v = false;
        let mut in_t = false;
        let mut in_f = false;
        let mut _in_hyperlinks = false;
        let mut current_merge_ref: Option<String> = None;
        let mut hyperlinks: HashMap<(u32, u32), String> = HashMap::new();
        let mut protection: Option<WorksheetProtection> = None;
        let mut reserved_cells = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) => {
                    let name = e.name();
                    let name = name.as_ref();
                    if name == b"sheetProtection" {
                        let mut prot = WorksheetProtection::default();
                        prot.sheet = true;
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                let attr_value = String::from_utf8_lossy(&attr.value);
                                let value_bool = attr_value == "1";
                                match attr_key {
                                    b"password" => prot.password = Some(attr_value.to_string()),
                                    b"selectLockedCells" => prot.select_locked_cells = value_bool,
                                    b"selectUnlockedCells" => {
                                        prot.select_unlocked_cells = value_bool
                                    }
                                    b"formatCells" => prot.format_cells = value_bool,
                                    b"formatColumns" => prot.format_columns = value_bool,
                                    b"formatRows" => prot.format_rows = value_bool,
                                    b"insertColumns" => prot.insert_columns = value_bool,
                                    b"insertRows" => prot.insert_rows = value_bool,
                                    b"insertHyperlinks" => prot.insert_hyperlinks = value_bool,
                                    b"deleteColumns" => prot.delete_columns = value_bool,
                                    b"deleteRows" => prot.delete_rows = value_bool,
                                    b"sort" => prot.sort = value_bool,
                                    b"autoFilter" => prot.auto_filter = value_bool,
                                    b"pivotTables" => prot.pivot_tables = value_bool,
                                    b"objects" => prot.objects = value_bool,
                                    b"scenarios" => prot.scenarios = value_bool,
                                    _ => {}
                                }
                            }
                        }
                        protection = Some(prot);
                    } else if name == b"dimension" && !reserved_cells {
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key;
                                let attr_key = attr_key.as_ref();
                                if attr_key == b"ref" {
                                    let ref_str = String::from_utf8_lossy(&attr.value);
                                    if let Some(cap) =
                                        Self::estimate_dimension_cells(ref_str.as_ref())
                                    {
                                        worksheet.cells.reserve(cap);
                                        reserved_cells = true;
                                    }
                                }
                            }
                        }
                    } else if name == b"mergeCell" {
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                if attr_key == b"ref" {
                                    let ref_str = String::from_utf8_lossy(&attr.value);
                                    if let Some(dash_pos) = ref_str.find(':') {
                                        let start = ref_str[..dash_pos].to_string();
                                        let end = ref_str[dash_pos + 1..].to_string();
                                        worksheet.add_merged_cell(start, end);
                                    }
                                }
                            }
                        }
                    } else if name == b"hyperlink" {
                        let mut hyperlink_ref: Option<String> = None;
                        let mut hyperlink_url: Option<String> = None;
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                if attr_key == b"ref" {
                                    hyperlink_ref =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                } else if attr_key == b"location" {
                                    hyperlink_url =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                        if let Some(ref_coord) = hyperlink_ref {
                            if let Ok((row, col)) = parse_coordinate(&ref_coord) {
                                if let Some(url) = hyperlink_url {
                                    hyperlinks.insert((row, col), url);
                                } else {
                                    hyperlinks.insert((row, col), format!("#{}", ref_coord));
                                }
                            }
                        }
                    } else if name == b"col" {
                        let mut col_min: Option<u32> = None;
                        let mut col_max: Option<u32> = None;
                        let mut width: Option<f64> = None;
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                if attr_key == b"min" {
                                    if let Ok(num) =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>()
                                    {
                                        col_min = Some(num);
                                    }
                                } else if attr_key == b"max" {
                                    if let Ok(num) =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>()
                                    {
                                        col_max = Some(num);
                                    }
                                } else if attr_key == b"width" {
                                    if let Ok(w) =
                                        String::from_utf8_lossy(&attr.value).parse::<f64>()
                                    {
                                        width = Some(w);
                                    }
                                }
                            }
                        }
                        if let Some(w) = width {
                            let start = col_min.unwrap_or(1);
                            let end = col_max.unwrap_or(start);
                            for col in start..=end {
                                worksheet.set_column_width(col, w);
                            }
                        }
                    } else if name == b"c" {
                        // Handle self-closing cell elements like <c r="A1" t="inlineStr" />
                        // These are typically empty cells but with a specific type (e.g., empty string)
                        let mut cell_row: Option<u32> = None;
                        let mut cell_col: Option<u32> = None;
                        let mut cell_type: u8 = 0;
                        let mut style_id: Option<u32> = None;

                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                if attr_key == b"r" {
                                    if let Some((row, col)) = parse_coordinate_bytes(&attr.value) {
                                        cell_row = Some(row);
                                        cell_col = Some(col);
                                    }
                                } else if attr_key == b"t" {
                                    cell_type = attr.value.first().copied().unwrap_or(0);
                                } else if attr_key == b"s" {
                                    style_id = parse_u32_bytes(&attr.value);
                                }
                            }
                        }

                        if let (Some(row), Some(col)) = (cell_row, cell_col) {
                            // If it's marked as a string type (inline or shared), treat as empty string
                            // Otherwise it's truly empty
                            let cell_value = if cell_type == b'i' || cell_type == b's' {
                                CellValue::String(std::sync::Arc::from(""))
                            } else {
                                CellValue::Empty
                            };

                            let style = style_id.and_then(|id| styles.get(&id).cloned());
                            let num_format = style.as_ref().and_then(|s| s.number_format.clone());
                            let data_type_str = match cell_type {
                                b's' => Some("s".to_string()),
                                b'b' => Some("b".to_string()),
                                b'd' => Some("d".to_string()),
                                b'i' => Some("str".to_string()),
                                _ => None,
                            };

                            let cell_data = CellData {
                                value: cell_value,
                                style,
                                style_index: style_id,
                                number_format: num_format,
                                data_type: data_type_str,
                                hyperlink: None,
                                comment: None,
                            };

                            worksheet.set_cell_data(row, col, cell_data);
                        }
                    }
                }
                Ok(Event::Start(e)) => {
                    let name = e.name();
                    let name = name.as_ref();

                    if name == b"dimension" && !reserved_cells {
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key;
                                let attr_key = attr_key.as_ref();
                                if attr_key == b"ref" {
                                    let ref_str = String::from_utf8_lossy(&attr.value);
                                    if let Some(cap) =
                                        Self::estimate_dimension_cells(ref_str.as_ref())
                                    {
                                        worksheet.cells.reserve(cap);
                                        reserved_cells = true;
                                    }
                                }
                            }
                        }
                    } else if name == b"row" {
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                if attr_key == b"r" {
                                    let r_str = String::from_utf8_lossy(&attr.value);
                                    current_row = r_str.parse().ok();
                                } else if attr_key == b"ht" {
                                    if let (Some(row), Ok(height)) = (
                                        current_row,
                                        String::from_utf8_lossy(&attr.value).parse::<f64>(),
                                    ) {
                                        worksheet.set_row_height(row, height);
                                    }
                                }
                            }
                        }
                    } else if name == b"c" {
                        in_cell = true;
                        current_value = None;
                        current_type = 0; // 0 = number (default)
                        current_style_id = None;
                        current_formula = None;
                        current_number_format = None;

                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                if attr_key == b"r" {
                                    // Use byte-based coordinate parsing (no String allocation)
                                    if let Some((row, col)) = parse_coordinate_bytes(&attr.value) {
                                        current_row = Some(row);
                                        current_col = Some(col);
                                    }
                                } else if attr_key == b"t" {
                                    // Store just the first byte of type (s, b, d, i, n)
                                    current_type = attr.value.first().copied().unwrap_or(0);
                                } else if attr_key == b"s" {
                                    // Parse style index directly from bytes
                                    current_style_id = parse_u32_bytes(&attr.value);
                                }
                            }
                        }
                    } else if name == b"v" {
                        in_v = true;
                    } else if name == b"t" {
                        in_t = true;
                    } else if name == b"f" {
                        in_f = true;
                    } else if name == b"mergeCell" {
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                if attr_key == b"ref" {
                                    current_merge_ref =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                    } else if name == b"col" {
                        let mut col_min: Option<u32> = None;
                        let mut col_max: Option<u32> = None;
                        let mut width: Option<f64> = None;
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                if attr_key == b"min" {
                                    if let Ok(num) =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>()
                                    {
                                        col_min = Some(num);
                                    }
                                } else if attr_key == b"max" {
                                    if let Ok(num) =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>()
                                    {
                                        col_max = Some(num);
                                    }
                                } else if attr_key == b"width" {
                                    if let Ok(w) =
                                        String::from_utf8_lossy(&attr.value).parse::<f64>()
                                    {
                                        width = Some(w);
                                    }
                                }
                            }
                        }
                        if let Some(w) = width {
                            let start = col_min.unwrap_or(1);
                            let end = col_max.unwrap_or(start);
                            for col in start..=end {
                                worksheet.set_column_width(col, w);
                            }
                        }
                    }
                }
                Ok(Event::Text(e)) => {
                    let text = e.unescape().unwrap_or_default();
                    if in_v && in_cell {
                        // Use byte-based type check (b's'=shared, b'b'=bool, b'd'=date)
                        current_value = match current_type {
                            b's' => {
                                // Shared string index - parse directly
                                match text.parse::<usize>() {
                                    Ok(idx) => Some(TempValue::SharedIdx(idx)),
                                    Err(_) => Some(TempValue::String(text.into_owned())),
                                }
                            }
                            b'b' => {
                                // Boolean - check first byte
                                let is_true = text.as_bytes().first() == Some(&b'1');
                                Some(TempValue::Bool(is_true))
                            }
                            b'd' => Some(TempValue::Date(text.into_owned())),
                            _ => {
                                // Number (default) - try fast f64 parsing
                                match parse_f64_bytes(text.as_bytes()) {
                                    Some(n) => Some(TempValue::Number(n)),
                                    None => Some(TempValue::String(text.into_owned())),
                                }
                            }
                        };
                    } else if in_t && in_cell {
                        current_value = Some(TempValue::String(text.into_owned()));
                    } else if in_f && in_cell {
                        current_formula = Some(text.to_string());
                    }
                }
                Ok(Event::End(e)) => {
                    let name = e.name();
                    let name = name.as_ref();

                    if name == b"hyperlinks" {
                        _in_hyperlinks = false;
                        for ((row, col), url) in &hyperlinks {
                            if let Some(cell_data) =
                                worksheet.cells.get_mut(&cell_key(*row, *col))
                            {
                                cell_data.hyperlink = Some(url.clone());
                            } else {
                                let cell_data = CellData {
                                    value: CellValue::Empty,
                                    style: None,
                                    style_index: None,
                                    number_format: None,
                                    data_type: None,
                                    hyperlink: Some(url.clone()),
                                    comment: None,
                                };
                                worksheet.set_cell_data(*row, *col, cell_data);
                            }
                        }
                    } else if name == b"c" {
                        if let (Some(row), Some(col)) = (current_row, current_col) {
                            let cell_value = if let Some(formula) = current_formula.take() {
                                CellValue::Formula(formula)
                            } else if let Some(value) = current_value.take() {
                                match value {
                                    TempValue::SharedIdx(idx) => {
                                        if idx < shared_strings.len() {
                                            CellValue::String(shared_strings[idx].clone())
                                        } else {
                                            CellValue::String(std::sync::Arc::from(idx.to_string()))
                                        }
                                    }
                                    TempValue::Bool(b) => CellValue::Boolean(b),
                                    TempValue::Number(n) => CellValue::Number(n),
                                    TempValue::Date(d) => CellValue::Date(d),
                                    TempValue::String(s) => CellValue::String(std::sync::Arc::from(s)),
                                }
                            } else {
                                // If it was marked as a string type but has no value,
                                // treat it as an empty string (openpyxl writes empty strings this way)
                                if current_type == b'i' || current_type == b's' {
                                    CellValue::String(std::sync::Arc::from(""))
                                } else {
                                    CellValue::Empty
                                }
                            };

                            let style = current_style_id.and_then(|id| styles.get(&id).cloned());

                            let num_format = current_number_format
                                .take()
                                .or_else(|| style.as_ref().and_then(|s| s.number_format.clone()));

                            // Convert u8 type back to Option<String> for CellData
                            // Only allocate if there's an explicit type
                            let data_type_str = match current_type {
                                b's' => Some("s".to_string()),
                                b'b' => Some("b".to_string()),
                                b'd' => Some("d".to_string()),
                                b'i' => Some("str".to_string()),
                                _ => None,
                            };

                            let cell_data = CellData {
                                value: cell_value,
                                style,
                                style_index: current_style_id,
                                number_format: num_format,
                                data_type: data_type_str,
                                hyperlink: None,
                                comment: None,
                            };

                            worksheet.set_cell_data(row, col, cell_data);
                        }
                        in_cell = false;
                        current_type = 0;
                        current_style_id = None;
                    } else if name == b"v" {
                        in_v = false;
                    } else if name == b"t" {
                        in_t = false;
                    } else if name == b"f" {
                        in_f = false;
                    } else if name == b"row" {
                        current_row = None;
                    } else if name == b"mergeCell" {
                        if let Some(ref_str) = current_merge_ref.take() {
                            if let Some(dash_pos) = ref_str.find(':') {
                                let start = ref_str[..dash_pos].to_string();
                                let end = ref_str[dash_pos + 1..].to_string();
                                worksheet.add_merged_cell(start, end);
                            }
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(RustypyxlError::ParseError(format!(
                        "XML parsing error: {}",
                        e
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        worksheet.protection = protection;

        Ok(())
    }

    fn parse_comments_xml<R: BufRead>(reader: R, worksheet: &mut Worksheet) -> Result<()> {
        let mut reader = Reader::from_reader(reader);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut current_cell_ref: Option<String> = None;
        let mut current_comment_text = String::new();
        let mut in_comment = false;
        let mut in_text = false;
        let mut in_t = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    let name = e.name();
                    let name = name.as_ref();
                    if name == b"comment" {
                        in_comment = true;
                        current_comment_text.clear();
                        for attr in e.attributes() {
                            if let Ok(attr) = attr {
                                let attr_key = attr.key.as_ref();
                                if attr_key == b"ref" {
                                    current_cell_ref =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                    } else if name == b"text" && in_comment {
                        in_text = true;
                    } else if name == b"t" && in_text {
                        in_t = true;
                    }
                }
                Ok(Event::Text(e)) => {
                    if in_t && in_text && in_comment {
                        let text = e.unescape().unwrap_or_default();
                        current_comment_text.push_str(&text);
                    }
                }
                Ok(Event::End(e)) => {
                    let name = e.name();
                    let name = name.as_ref();
                    if name == b"comment" {
                        if let Some(ref_coord) = current_cell_ref.take() {
                            if let Ok((row, col)) = parse_coordinate(&ref_coord) {
                                worksheet.set_cell_comment(row, col, current_comment_text.clone());
                            }
                        }
                        in_comment = false;
                        in_text = false;
                        in_t = false;
                        current_comment_text.clear();
                    } else if name == b"text" {
                        in_text = false;
                    } else if name == b"t" {
                        in_t = false;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(RustypyxlError::ParseError(format!(
                        "XML parsing error in comments: {}",
                        e
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(())
    }
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_workbook_new() {
        let wb = Workbook::new();
        assert!(wb.worksheets.is_empty());
        assert!(wb.sheet_names.is_empty());
    }

    #[test]
    fn test_create_sheet() {
        let mut wb = Workbook::new();
        let _ = wb.create_sheet(Some("Sheet1".to_string())).unwrap();
        assert_eq!(wb.sheet_names.len(), 1);
        assert_eq!(wb.sheet_names[0], "Sheet1");
    }

    #[test]
    fn test_create_sheet_duplicate() {
        let mut wb = Workbook::new();
        let _ = wb.create_sheet(Some("Sheet1".to_string())).unwrap();
        let result = wb.create_sheet(Some("Sheet1".to_string()));
        assert!(result.is_err());
    }

    #[test]
    fn test_get_sheet_by_name() {
        let mut wb = Workbook::new();
        let _ = wb.create_sheet(Some("MySheet".to_string())).unwrap();
        let ws = wb.get_sheet_by_name("MySheet").unwrap();
        assert_eq!(ws.title(), "MySheet");
    }

    #[test]
    fn test_remove_sheet() {
        let mut wb = Workbook::new();
        let _ = wb.create_sheet(Some("Sheet1".to_string())).unwrap();
        let _ = wb.create_sheet(Some("Sheet2".to_string())).unwrap();
        wb.remove_sheet("Sheet1").unwrap();
        assert_eq!(wb.sheet_names.len(), 1);
        assert_eq!(wb.sheet_names[0], "Sheet2");
    }

    #[test]
    fn test_named_ranges() {
        let mut wb = Workbook::new();
        wb.create_named_range("MyRange".to_string(), "'Sheet1'!A1:B10".to_string())
            .unwrap();
        assert_eq!(wb.get_named_range("MyRange"), Some("'Sheet1'!A1:B10"));
    }

    #[test]
    fn test_parse_workbook_rels() {
        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet5.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
</Relationships>"#;

        let rels = Workbook::parse_workbook_rels(Cursor::new(rels_xml)).unwrap();

        assert_eq!(rels.get("rId1"), Some(&"worksheets/sheet1.xml".to_string()));
        assert_eq!(rels.get("rId2"), Some(&"worksheets/sheet5.xml".to_string()));
        assert_eq!(rels.get("rId3"), Some(&"worksheets/sheet3.xml".to_string()));
    }

    #[test]
    fn test_parse_workbook_xml_with_rids() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Data" sheetId="8" r:id="rId1"/>
        <sheet name="Summary" sheetId="2" r:id="rId2"/>
    </sheets>
</workbook>"#;

        let (sheets, _) = Workbook::parse_workbook_xml(Cursor::new(workbook_xml)).unwrap();

        assert_eq!(sheets.len(), 2);
        assert_eq!(sheets[0], ("Data".to_string(), 8, "rId1".to_string()));
        assert_eq!(sheets[1], ("Summary".to_string(), 2, "rId2".to_string()));
    }

    #[test]
    fn test_save_to_bytes() {
        let mut wb = Workbook::new();
        let ws = wb.create_sheet(Some("Test".to_string())).unwrap();
        ws.set_cell_value(1, 1, CellValue::String(std::sync::Arc::from("Hello")));
        ws.set_cell_value(1, 2, CellValue::Number(42.0));
        ws.set_cell_value(2, 1, CellValue::Boolean(true));

        let bytes = wb.save_to_bytes().unwrap();

        // Verify it's a valid ZIP file (starts with PK)
        assert!(bytes.len() > 4);
        assert_eq!(&bytes[0..2], b"PK");
    }

    #[test]
    fn test_load_from_bytes() {
        // Create a workbook with data
        let mut wb = Workbook::new();
        let ws = wb.create_sheet(Some("TestSheet".to_string())).unwrap();
        ws.set_cell_value(1, 1, CellValue::String(std::sync::Arc::from("Hello World")));
        ws.set_cell_value(1, 2, CellValue::Number(123.45));

        // Save to bytes
        let bytes = wb.save_to_bytes().unwrap();

        // Load from bytes
        let wb2 = Workbook::load_from_bytes(&bytes).unwrap();

        // Verify the loaded workbook
        assert_eq!(wb2.sheet_names.len(), 1);
        assert_eq!(wb2.sheet_names[0], "TestSheet");

        let ws2 = wb2.get_sheet_by_name("TestSheet").unwrap();
        let cell1 = ws2.get_cell(1, 1).unwrap();
        let cell2 = ws2.get_cell(1, 2).unwrap();

        match &cell1.value {
            CellValue::String(s) => assert_eq!(s.as_ref(), "Hello World"),
            _ => panic!("Expected String value"),
        }

        match &cell2.value {
            CellValue::Number(n) => assert!((n - 123.45).abs() < 0.001),
            _ => panic!("Expected Number value"),
        }
    }

    #[test]
    fn test_bytes_roundtrip_with_multiple_sheets() {
        let mut wb = Workbook::new();

        // Create multiple sheets with data
        let ws1 = wb.create_sheet(Some("Sheet1".to_string())).unwrap();
        ws1.set_cell_value(1, 1, CellValue::String(std::sync::Arc::from("Sheet1 Data")));

        let ws2 = wb.create_sheet(Some("Sheet2".to_string())).unwrap();
        ws2.set_cell_value(1, 1, CellValue::String(std::sync::Arc::from("Sheet2 Data")));
        ws2.set_cell_value(2, 2, CellValue::Number(999.0));

        // Roundtrip through bytes
        let bytes = wb.save_to_bytes().unwrap();
        let wb2 = Workbook::load_from_bytes(&bytes).unwrap();

        // Verify
        assert_eq!(wb2.sheet_names.len(), 2);
        assert!(wb2.sheet_names.contains(&"Sheet1".to_string()));
        assert!(wb2.sheet_names.contains(&"Sheet2".to_string()));
    }
}

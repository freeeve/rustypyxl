//! Workbook representation and file I/O operations.

#[cfg(feature = "fast-hash")]
use hashbrown::HashMap;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use rayon::prelude::*;
#[cfg(not(feature = "fast-hash"))]
use std::collections::HashMap;
use std::fs::File;
use std::io::{BufRead, BufReader, Cursor, Read, Seek};
use std::sync::Arc;
use zip::ZipArchive;

use crate::autofilter::{
    AutoFilter, ColorFilter, CustomFilter, DynamicFilterType, FilterColumn, FilterOperator,
    FilterType, Top10Filter,
};
use crate::cell::CellValue;
use crate::conditional::{
    ColorScale, ConditionalColor, ConditionalFormat, ConditionalFormatType, ConditionalFormatting,
    ConditionalOperator, ConditionalRule, DataBar, IconSet, IconSetStyle,
};
use crate::error::{Result, RustypyxlError};
use crate::pagesetup::{Orientation, PageSetup, PaperSize};
use crate::style::{
    Alignment, Border, BorderStyle, CellStyle, CellXf, Fill, Font, Protection, StyleRegistry,
};
use crate::table::{Table, TableColumn, TableStyle, TotalsRowFunction};
use crate::utils::{parse_coordinate, parse_coordinate_bytes, parse_f64_bytes, parse_u32_bytes};
use crate::worksheet::{
    cell_key, CellData, DataValidation, SheetVisibility, Worksheet, WorksheetProtection,
};
use crate::writer;

/// A named range definition.
#[derive(Clone, Debug)]
pub struct NamedRange {
    /// Name of the range.
    pub name: String,
    /// Range reference (e.g., "'Sheet1'!A1:B2").
    pub range: String,
    /// Sheet index this name is scoped to (None = workbook-global).
    pub local_sheet_id: Option<u32>,
    /// Hidden from the Excel name manager UI.
    pub hidden: bool,
}

/// Compression level for saving workbooks.
#[derive(Clone, Copy, Debug, PartialEq, Default)]
pub enum CompressionLevel {
    /// No compression - fastest saves, largest files
    None,
    /// Fast compression (deflate level 1) - good balance
    Fast,
    /// Default compression (deflate level 6) - smaller files, slower
    #[default]
    Default,
    /// Best compression (deflate level 9) - smallest files, slowest
    Best,
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
    /// Index of the active (selected) sheet tab.
    pub active_sheet: usize,
    /// True when the file uses the 1904 date system (Excel for Mac's legacy
    /// epoch). Date serials are stored as written; this preserves the flag so
    /// consumers can interpret them against the right epoch.
    pub date1904: bool,
    /// Monotonic source for Worksheet::uid values; never reused so stale
    /// handles can't silently resolve to a different sheet.
    next_sheet_uid: u64,
}

/// (sheet name, sheet id, relationship id, visibility) parsed from workbook.xml.
type SheetInfo = (String, u32, String, SheetVisibility);

/// A single entry from a worksheet's .rels part.
#[derive(Clone, Debug)]
pub(crate) struct SheetRel {
    /// Relationship type URI (e.g. ".../hyperlink", ".../comments", ".../table").
    pub rel_type: String,
    /// Target, relative to the worksheet part unless `external`.
    pub target: String,
    /// TargetMode="External" (hyperlinks to URLs).
    pub external: bool,
}

/// Everything read from the archive for one sheet before parsing.
struct SheetParseInput {
    name: String,
    visibility: SheetVisibility,
    sheet_xml: Vec<u8>,
    comments_xml: Option<Vec<u8>>,
    rels: HashMap<String, SheetRel>,
    table_xmls: Vec<Vec<u8>>,
}

/// Resolve a relationship target relative to the part that declares it.
/// `base_part` is a full package path like "xl/worksheets/sheet1.xml".
fn resolve_rel_target(base_part: &str, target: &str) -> String {
    if let Some(stripped) = target.strip_prefix('/') {
        return stripped.to_string();
    }
    let base_dir = match base_part.rfind('/') {
        Some(idx) => &base_part[..idx],
        None => "",
    };
    let mut parts: Vec<&str> = base_dir.split('/').filter(|p| !p.is_empty()).collect();
    for seg in target.split('/') {
        match seg {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            other => parts.push(other),
        }
    }
    parts.join("/")
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
            active_sheet: 0,
            date1904: false,
            next_sheet_uid: 1,
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
        workbook.parse_workbook(&mut archive)?;

        Ok(workbook)
    }

    /// Load a workbook from bytes (e.g., from memory or network).
    pub fn load_from_bytes(data: &[u8]) -> Result<Self> {
        let cursor = Cursor::new(data);
        let mut archive = ZipArchive::new(cursor)?;

        let mut workbook = Workbook::new();
        workbook.parse_workbook(&mut archive)?;

        Ok(workbook)
    }

    /// Get the active (first) worksheet.
    pub fn active(&self) -> Result<&Worksheet> {
        self.worksheets.first().ok_or(RustypyxlError::NoWorksheets)
    }

    /// Get a mutable reference to the active worksheet.
    pub fn active_mut(&mut self) -> Result<&mut Worksheet> {
        self.worksheets
            .first_mut()
            .ok_or(RustypyxlError::NoWorksheets)
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
        self.worksheets
            .get(index)
            .ok_or_else(|| RustypyxlError::WorksheetNotFound(format!("index {}", index)))
    }

    /// Get a mutable worksheet by index.
    pub fn get_sheet_by_index_mut(&mut self, index: usize) -> Result<&mut Worksheet> {
        self.worksheets
            .get_mut(index)
            .ok_or_else(|| RustypyxlError::WorksheetNotFound(format!("index {}", index)))
    }

    /// Create a new worksheet.
    pub fn create_sheet(&mut self, title: Option<String>) -> Result<&mut Worksheet> {
        let sheet_title = title.unwrap_or_else(|| format!("Sheet{}", self.worksheets.len() + 1));

        if self.sheet_names.contains(&sheet_title) {
            return Err(RustypyxlError::WorksheetAlreadyExists(sheet_title));
        }

        let mut worksheet = Worksheet::new(sheet_title.clone());
        worksheet.uid = self.allocate_sheet_uid();
        self.worksheets.push(worksheet);
        self.sheet_names.push(sheet_title);

        Ok(self.worksheets.last_mut().unwrap())
    }

    /// Hand out the next stable sheet uid. Callers adding worksheets to
    /// `worksheets` directly (e.g. when cloning a sheet) must stamp the new
    /// sheet with this so handle resolution stays unambiguous.
    pub fn allocate_sheet_uid(&mut self) -> u64 {
        let uid = self.next_sheet_uid;
        self.next_sheet_uid += 1;
        uid
    }

    /// Find the current position of the sheet with the given stable uid.
    pub fn sheet_index_by_uid(&self, uid: u64) -> Option<usize> {
        if uid == 0 {
            return None;
        }
        self.worksheets.iter().position(|ws| ws.uid == uid)
    }

    /// Remove a worksheet by name.
    ///
    /// The active tab follows the sheet it pointed at: removing a sheet before
    /// it shifts the index down, and removing the active sheet itself leaves
    /// the index in place so it lands on the next sheet (clamped to the end),
    /// matching openpyxl.
    pub fn remove_sheet(&mut self, sheet_name: &str) -> Result<()> {
        for (idx, name) in self.sheet_names.iter().enumerate() {
            if name == sheet_name {
                self.worksheets.remove(idx);
                self.sheet_names.remove(idx);
                if idx < self.active_sheet {
                    self.active_sheet -= 1;
                }
                self.active_sheet = self
                    .active_sheet
                    .min(self.worksheets.len().saturating_sub(1));
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
        ws.set_cell_font(row, column, font);
        Ok(())
    }

    /// Set cell alignment in the active worksheet.
    pub fn set_cell_alignment(
        &mut self,
        row: u32,
        column: u32,
        alignment: Alignment,
    ) -> Result<()> {
        let ws = self.active_mut()?;
        ws.set_cell_alignment(row, column, alignment);
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
            ..Default::default()
        };
        ws.add_data_validation(row, column, validation);
        Ok(())
    }

    /// Create a named range.
    pub fn create_named_range(&mut self, name: String, range: String) -> Result<()> {
        if self.named_ranges.iter().any(|nr| nr.name == name) {
            return Err(RustypyxlError::NamedRangeAlreadyExists(name));
        }
        self.named_ranges.push(NamedRange {
            name,
            range,
            local_sheet_id: None,
            hidden: false,
        });
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
    fn get_file_options(
        &self,
    ) -> zip::write::FileOptions<'static, zip::write::ExtendedFileOptions> {
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
    fn write_workbook_contents<W: std::io::Write + Seek>(
        &self,
        zip: &mut zip::ZipWriter<W>,
    ) -> Result<()> {
        use std::io::Write;
        use zip::write::FileOptions;

        let options = self.get_file_options();

        // Collect shared strings first to know if we have any
        let (shared_strings_vec, shared_strings_map, shared_strings_refs) =
            writer::collect_shared_strings(&self.worksheets);
        let has_shared_strings = !shared_strings_vec.is_empty();

        // Pre-compute per-sheet metadata so [Content_Types].xml, the sheet
        // XML, and the sheet .rels parts all agree on ids and paths.
        let comment_sheet_ids: Vec<u32> = self
            .worksheets
            .iter()
            .enumerate()
            .filter(|(_, ws)| ws.cells.values().any(|cd| cd.comment.is_some()))
            .map(|(idx, _)| (idx + 1) as u32)
            .collect();

        // Assign each table a workbook-unique id; part path is xl/tables/table{id}.xml
        let mut table_assignments: Vec<Vec<u32>> = Vec::with_capacity(self.worksheets.len());
        let mut next_table_id: u32 = 1;
        for worksheet in &self.worksheets {
            let ids: Vec<u32> = worksheet
                .tables
                .iter()
                .map(|_| {
                    let id = next_table_id;
                    next_table_id += 1;
                    id
                })
                .collect();
            table_assignments.push(ids);
        }
        let table_count = (next_table_id - 1) as usize;

        // Write [Content_Types].xml
        writer::write_content_types(
            zip,
            &options,
            self.worksheets.len(),
            has_shared_strings,
            &comment_sheet_ids,
            table_count,
        )?;

        // Write _rels/.rels
        writer::write_rels(zip, &options)?;

        // Write docProps files
        writer::write_doc_props(zip, &options)?;

        // Write xl/workbook.xml
        let sheet_meta: Vec<(String, crate::worksheet::SheetVisibility)> = self
            .sheet_names
            .iter()
            .zip(&self.worksheets)
            .map(|(name, ws)| (name.clone(), ws.visibility))
            .collect();
        writer::write_workbook_xml(
            zip,
            &options,
            &sheet_meta,
            &self.named_ranges,
            self.active_sheet,
            self.date1904,
        )?;

        // Write xl/_rels/workbook.xml.rels
        writer::write_workbook_rels(zip, &options, self.worksheets.len(), has_shared_strings)?;

        // Write shared strings if we have any
        if has_shared_strings {
            writer::write_shared_strings(zip, &options, &shared_strings_vec, shared_strings_refs)?;
        }

        // Resolve styles set through the core API into registry xfs. Cells
        // styled via set_cell_style/set_cell_font/... carry the style on
        // CellData but have no xf index (style_index is None), so without
        // this pass the writer would emit them unstyled.
        let mut styles_for_save = self.styles.clone();
        let style_overrides: Vec<std::collections::HashMap<u64, u32>> = self
            .worksheets
            .iter()
            .map(|ws| {
                let mut overrides = std::collections::HashMap::new();
                for (key, cell) in &ws.cells {
                    if cell.style_index.is_none()
                        && (cell.style.is_some() || cell.number_format.is_some())
                    {
                        let mut style = cell
                            .style
                            .as_deref()
                            .cloned()
                            .unwrap_or_else(crate::style::CellStyle::new);
                        if style.number_format.is_none() {
                            style.number_format = cell.number_format.clone();
                        }
                        let idx = styles_for_save.get_or_add_cell_xf(&style);
                        overrides.insert(*key, idx as u32);
                    }
                }
                overrides
            })
            .collect();

        // Write styles.xml with the differential formats used by
        // conditional-formatting rules (referenced by dxfId)
        let dxfs = writer::collect_dxfs(&self.worksheets);
        writer::write_styles_xml(zip, &options, &styles_for_save, &dxfs)?;

        // Write each worksheet, its tables/comments, and its .rels part
        for (idx, worksheet) in self.worksheets.iter().enumerate() {
            let sheet_id = (idx + 1) as u32;
            let has_comments = comment_sheet_ids.contains(&sheet_id);
            let table_ids = &table_assignments[idx];
            let table_rel_ids: Vec<String> = table_ids
                .iter()
                .map(|id| format!("rIdTable{}", id))
                .collect();

            writer::write_worksheet_xml(
                zip,
                &options,
                worksheet,
                sheet_id,
                &shared_strings_map,
                &table_rel_ids,
                &dxfs,
                has_comments,
                &style_overrides[idx],
            )?;

            for (table, table_id) in worksheet.tables.iter().zip(table_ids) {
                writer::write_table_xml(zip, &options, table, *table_id)?;
            }

            if has_comments {
                writer::write_comments_xml(zip, &options, worksheet, sheet_id)?;
                writer::write_vml_drawing(zip, &options, worksheet, sheet_id)?;
            }

            // The sheet .rels part ties comments, external hyperlinks, and
            // tables to the relationship ids used in the worksheet XML.
            let external_links = writer::collect_external_hyperlinks(worksheet);
            if has_comments || !external_links.is_empty() || !table_ids.is_empty() {
                let rels_path = format!("xl/worksheets/_rels/sheet{}.xml.rels", sheet_id);
                let rels_options: zip::write::FileOptions<
                    'static,
                    zip::write::ExtendedFileOptions,
                > = FileOptions::default().compression_method(zip::CompressionMethod::Deflated);
                zip.start_file(&rels_path, rels_options)?;

                let mut rels_content = String::from(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
                );
                if has_comments {
                    rels_content.push_str(&format!(
                        "<Relationship Id=\"rIdComments\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../comments/comment{}.xml\"/>\n",
                        sheet_id
                    ));
                    rels_content.push_str(&format!(
                        "<Relationship Id=\"rIdVml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/vmlDrawing{}.vml\"/>\n",
                        sheet_id
                    ));
                }
                for (i, (_, url)) in external_links.iter().enumerate() {
                    rels_content.push_str(&format!(
                        "<Relationship Id=\"rIdHL{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{}\" TargetMode=\"External\"/>\n",
                        i + 1,
                        writer::escape_xml(url)
                    ));
                }
                for table_id in table_ids {
                    rels_content.push_str(&format!(
                        "<Relationship Id=\"rIdTable{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table{}.xml\"/>\n",
                        table_id, table_id
                    ));
                }
                rels_content.push_str("</Relationships>");
                zip.write_all(rels_content.as_bytes())?;
            }
        }

        Ok(())
    }

    /// Parse workbook from ZIP archive with parallel worksheet parsing.
    fn parse_workbook<R: Read + Seek>(&mut self, archive: &mut ZipArchive<R>) -> Result<()> {
        // Phase 1: Load all file contents into memory (sequential ZIP extraction)
        let workbook_xml = Self::read_zip_file_to_vec(archive, "xl/workbook.xml")?;
        let workbook_rels_xml =
            Self::read_zip_file_to_vec(archive, "xl/_rels/workbook.xml.rels").ok();
        let shared_strings_xml = Self::read_zip_file_to_vec(archive, "xl/sharedStrings.xml").ok();
        let styles_xml = Self::read_zip_file_to_vec(archive, "xl/styles.xml").ok();

        // Parse workbook.xml to get sheet names, IDs, relationship IDs,
        // visibility, and the active tab
        let (sheet_info, named_ranges, active_tab, date1904) =
            Self::parse_workbook_xml(Cursor::new(&workbook_xml))?;
        self.named_ranges = named_ranges;
        self.active_sheet = active_tab;
        self.date1904 = date1904;

        // Parse workbook.xml.rels to get the mapping from rId to actual file paths
        let rels_map: HashMap<String, String> = if let Some(rels_xml) = workbook_rels_xml {
            Self::parse_workbook_rels(Cursor::new(&rels_xml))?
        } else {
            HashMap::new()
        };

        // Load all worksheet XML, sheet rels, comments, and table parts into memory
        let mut sheet_data: Vec<SheetParseInput> = Vec::with_capacity(sheet_info.len());
        for (sheet_name, sheet_id, sheet_rid, visibility) in &sheet_info {
            // Look up the actual sheet path from the relationships, or fall back to sheetId-based path
            let sheet_path = if let Some(target) = rels_map.get(sheet_rid) {
                // Target is relative to xl/, e.g., "worksheets/sheet1.xml"
                if let Some(stripped) = target.strip_prefix('/') {
                    // Absolute path within the package (rare)
                    stripped.to_string()
                } else {
                    format!("xl/{}", target)
                }
            } else {
                // Fallback to legacy behavior if rels file is missing or incomplete
                format!("xl/worksheets/sheet{}.xml", sheet_id)
            };
            let sheet_xml = Self::read_zip_file_to_vec(archive, &sheet_path)?;

            // The sheet's .rels part lives at <dir>/_rels/<file>.rels
            let rels_path = match sheet_path.rfind('/') {
                Some(idx) => format!(
                    "{}/_rels/{}.rels",
                    &sheet_path[..idx],
                    &sheet_path[idx + 1..]
                ),
                None => format!("_rels/{}.rels", sheet_path),
            };
            let rels = match Self::read_zip_file_to_vec(archive, &rels_path) {
                Ok(xml) => Self::parse_sheet_rels(Cursor::new(&xml)).unwrap_or_default(),
                Err(_) => HashMap::new(),
            };

            // Comments: resolve via the sheet rels (real files use xl/comments1.xml),
            // falling back to the legacy path this library used to write.
            let comments_path = rels
                .values()
                .find(|r| r.rel_type.ends_with("/comments"))
                .map(|r| resolve_rel_target(&sheet_path, &r.target))
                .unwrap_or_else(|| format!("xl/comments/comment{}.xml", sheet_id));
            let comments_xml = Self::read_zip_file_to_vec(archive, &comments_path).ok();

            // Table parts referenced from this sheet
            let table_xmls: Vec<Vec<u8>> = rels
                .values()
                .filter(|r| r.rel_type.ends_with("/table"))
                .filter_map(|r| {
                    Self::read_zip_file_to_vec(archive, &resolve_rel_target(&sheet_path, &r.target))
                        .ok()
                })
                .collect();

            sheet_data.push(SheetParseInput {
                name: sheet_name.clone(),
                visibility: *visibility,
                sheet_xml,
                comments_xml,
                rels,
                table_xmls,
            });
        }

        // Phase 2: Parse shared data (must be done before worksheets)
        let shared_strings = if let Some(xml) = shared_strings_xml {
            Self::parse_shared_strings_xml(Cursor::new(&xml))?
        } else {
            Vec::new()
        };

        let (styles, mut style_registry) = if let Some(ref xml) = styles_xml {
            Self::parse_styles_xml(xml)?
        } else {
            (HashMap::new(), StyleRegistry::new())
        };
        if let Some(ref xml) = styles_xml {
            style_registry.dxfs = Self::parse_dxfs_xml(xml).unwrap_or_default();
        }

        // Phase 3: Parse worksheets in parallel using Rayon
        let shared_strings_ref = &shared_strings;
        let styles_ref = &styles;

        let dxfs_ref: &[ConditionalFormat] = &style_registry.dxfs;
        let parse_one = |input: &SheetParseInput| -> Result<(String, Worksheet)> {
            let mut worksheet = Worksheet::new(input.name.clone());
            worksheet.visibility = input.visibility;
            Self::parse_worksheet_xml(
                Cursor::new(&input.sheet_xml),
                shared_strings_ref,
                styles_ref,
                &input.rels,
                dxfs_ref,
                &mut worksheet,
                input.sheet_xml.len(),
            )?;

            if let Some(comments) = &input.comments_xml {
                Self::parse_comments_xml(Cursor::new(comments), &mut worksheet)?;
            }

            for table_xml in &input.table_xmls {
                if let Ok(table) = Self::parse_table_xml(Cursor::new(table_xml)) {
                    worksheet.tables.push(table);
                }
            }

            Ok((input.name.clone(), worksheet))
        };

        let worksheets: Vec<Result<(String, Worksheet)>> = if sheet_data.len() > 1 {
            // Parallel parsing for multiple sheets
            sheet_data.par_iter().map(parse_one).collect()
        } else {
            // Sequential for single sheet (avoid Rayon overhead)
            sheet_data.iter().map(parse_one).collect()
        };

        // Collect results in order, stamping each sheet with a stable uid
        for result in worksheets {
            let (sheet_name, mut worksheet) = result?;
            worksheet.uid = self.allocate_sheet_uid();
            self.worksheets.push(worksheet);
            self.sheet_names.push(sheet_name);
        }

        // Store the style registry
        self.styles = style_registry;

        Ok(())
    }

    /// Read a file from the ZIP archive into a Vec<u8>.
    /// The declared uncompressed size in the ZIP header is untrusted: it is
    /// rejected past a hard cap and only used for pre-allocation up to a small
    /// bound, so a crafted archive cannot trigger huge allocations up front.
    fn read_zip_file_to_vec<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
        path: &str,
    ) -> Result<Vec<u8>> {
        const MAX_PREALLOC: usize = 16 * 1024 * 1024;
        const MAX_PART_SIZE: u64 = 4 * 1024 * 1024 * 1024;

        let mut file = archive.by_name(path).map_err(|e| {
            RustypyxlError::InvalidFormat(format!("Failed to find {} in archive: {}", path, e))
        })?;
        let declared_size = file.size();
        if declared_size > MAX_PART_SIZE {
            return Err(RustypyxlError::InvalidFormat(format!(
                "Archive member {} declares an unreasonable uncompressed size of {} bytes",
                path, declared_size
            )));
        }
        let mut buf = Vec::with_capacity((declared_size as usize).min(MAX_PREALLOC));
        file.read_to_end(&mut buf)?;
        Ok(buf)
    }

    /// Parses workbook.xml and returns sheet info (name, sheetId, rId,
    /// visibility), named ranges, and the active tab index.
    fn parse_workbook_xml<R: BufRead>(
        reader: R,
    ) -> Result<(Vec<SheetInfo>, Vec<NamedRange>, usize, bool)> {
        let mut reader = Reader::from_reader(reader);
        reader.config_mut().trim_text(true);

        let mut sheets = Vec::new();
        let mut named_ranges = Vec::new();
        let mut active_tab: usize = 0;
        let mut date1904 = false;
        let mut buf = Vec::new();
        let mut current_sheet_name: Option<String> = None;
        let mut current_sheet_id: Option<u32> = None;
        let mut current_sheet_rid: Option<String> = None;
        let mut current_sheet_state = SheetVisibility::Visible;
        let mut in_defined_names = false;
        let mut current_name: Option<String> = None;
        let mut current_range: Option<String> = None;
        let mut current_local_sheet_id: Option<u32> = None;
        let mut current_hidden = false;
        let mut in_defined_name = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) => {
                    let name = e.name();
                    let local = e.local_name();
                    let name = name.as_ref();
                    let local = local.as_ref();

                    if local == b"workbookPr" {
                        date1904 = Self::parse_date1904(&e);
                    }

                    // Handle self-closing sheet tags
                    if name == b"sheet" || local == b"sheet" {
                        let mut sheet_name: Option<String> = None;
                        let mut sheet_id: Option<u32> = None;
                        let mut sheet_rid: Option<String> = None;
                        let mut sheet_state = SheetVisibility::Visible;

                        for attr in e.attributes().flatten() {
                            let attr_key = attr.key;
                            let attr_local = attr.key.local_name();
                            let attr_key = attr_key.as_ref();
                            let attr_local = attr_local.as_ref();

                            if attr_key == b"name" || attr_local == b"name" {
                                sheet_name = Some(String::from_utf8_lossy(&attr.value).to_string());
                            } else if attr_key == b"sheetId" || attr_local == b"sheetId" {
                                let id_str = String::from_utf8_lossy(&attr.value);
                                sheet_id = id_str.parse().ok();
                            } else if attr_key == b"state" || attr_local == b"state" {
                                sheet_state = SheetVisibility::from_attr(&String::from_utf8_lossy(
                                    &attr.value,
                                ));
                            } else if attr_local == b"id" {
                                // r:id attribute (namespace-qualified)
                                sheet_rid = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }

                        if let (Some(name), Some(id), Some(rid)) = (sheet_name, sheet_id, sheet_rid)
                        {
                            sheets.push((name, id, rid, sheet_state));
                        }
                    } else if name == b"workbookView" || local == b"workbookView" {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"activeTab" {
                                active_tab =
                                    String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                            }
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

                    if local == b"workbookPr" {
                        date1904 = Self::parse_date1904(&e);
                    }

                    if is_defined_names {
                        in_defined_names = true;
                    } else if is_defined_name && in_defined_names {
                        in_defined_name = true;
                        for attr in e.attributes().flatten() {
                            let attr_key = attr.key;
                            let attr_local = attr.key.local_name();
                            let attr_key = attr_key.as_ref();
                            let attr_local = attr_local.as_ref();
                            if attr_key == b"name" || attr_local == b"name" {
                                current_name =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            } else if attr_key == b"localSheetId" || attr_local == b"localSheetId" {
                                current_local_sheet_id =
                                    String::from_utf8_lossy(&attr.value).parse().ok();
                            } else if attr_key == b"hidden" || attr_local == b"hidden" {
                                let v = String::from_utf8_lossy(&attr.value);
                                current_hidden = v == "1" || v == "true";
                            }
                        }
                    } else if is_sheet {
                        for attr in e.attributes().flatten() {
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
                            } else if attr_key == b"state" || attr_local == b"state" {
                                current_sheet_state = SheetVisibility::from_attr(
                                    &String::from_utf8_lossy(&attr.value),
                                );
                            } else if attr_local == b"id" {
                                // r:id attribute (namespace-qualified)
                                current_sheet_rid =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
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
                        if let (Some(name), Some(range)) =
                            (current_name.take(), current_range.take())
                        {
                            named_ranges.push(NamedRange {
                                name,
                                range,
                                local_sheet_id: current_local_sheet_id.take(),
                                hidden: current_hidden,
                            });
                        }
                        current_local_sheet_id = None;
                        current_hidden = false;
                        in_defined_name = false;
                    } else if is_defined_names {
                        in_defined_names = false;
                    } else if is_sheet {
                        if let (Some(name), Some(id), Some(rid)) = (
                            current_sheet_name.take(),
                            current_sheet_id.take(),
                            current_sheet_rid.take(),
                        ) {
                            sheets.push((name, id, rid, current_sheet_state));
                        }
                        current_sheet_state = SheetVisibility::Visible;
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

        Ok((sheets, named_ranges, active_tab, date1904))
    }

    /// Reads the date1904 flag off `<workbookPr>`; Excel writes it as "1",
    /// other producers as "true".
    fn parse_date1904(e: &quick_xml::events::BytesStart) -> bool {
        e.attributes().flatten().any(|attr| {
            attr.key.local_name().as_ref() == b"date1904"
                && matches!(attr.value.as_ref(), b"1" | b"true")
        })
    }

    /// Parses a worksheet's .rels part into a map of relationship id -> SheetRel.
    fn parse_sheet_rels<R: BufRead>(reader: R) -> Result<HashMap<String, SheetRel>> {
        let mut reader = Reader::from_reader(reader);
        reader.config_mut().trim_text(true);

        let mut rels = HashMap::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"Relationship" {
                        let mut rel_id: Option<String> = None;
                        let mut rel_type: Option<String> = None;
                        let mut target: Option<String> = None;
                        let mut external = false;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"Id" => {
                                    rel_id = Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                b"Type" => {
                                    rel_type =
                                        Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                b"Target" => {
                                    // Targets may contain escaped entities (e.g. &amp; in URLs)
                                    target = attr.unescape_value().ok().map(|v| v.to_string())
                                }
                                b"TargetMode" => {
                                    external = attr.value.as_ref() == b"External";
                                }
                                _ => {}
                            }
                        }

                        if let (Some(id), Some(rel_type), Some(target)) = (rel_id, rel_type, target)
                        {
                            rels.insert(
                                id,
                                SheetRel {
                                    rel_type,
                                    target,
                                    external,
                                },
                            );
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(RustypyxlError::ParseError(format!(
                        "XML parsing error in sheet rels: {}",
                        e
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(rels)
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

                        for attr in e.attributes().flatten() {
                            let attr_key = attr.key.as_ref();
                            if attr_key == b"Id" {
                                rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            } else if attr_key == b"Target" {
                                target = Some(String::from_utf8_lossy(&attr.value).to_string());
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
                    // Local names so <x:t>/<x:si> parse like their unprefixed forms
                    if e.local_name().as_ref() == b"t" {
                        in_t = true;
                    }
                }
                Ok(Event::Text(e)) => {
                    if in_t {
                        current_string.push_str(&e.unescape().unwrap_or_default());
                    }
                }
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_t = false;
                    } else if e.local_name().as_ref() == b"si" {
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
    #[allow(dead_code)]
    fn get_attr_u32(e: &quick_xml::events::BytesStart, key: &[u8]) -> Option<u32> {
        Self::get_attr_str(e, key).and_then(|s| s.parse().ok())
    }

    /// Get an optional f64 attribute value from an XML element.
    fn get_attr_f64(e: &quick_xml::events::BytesStart, key: &[u8]) -> Option<f64> {
        Self::get_attr_str(e, key).and_then(|s| s.parse().ok())
    }

    /// Check if an attribute equals "1" or "true".
    #[allow(dead_code)]
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
            b"u" => {
                font.underline =
                    Some(Self::get_attr_str(e, b"val").unwrap_or_else(|| "single".to_string()))
            }
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
    #[allow(dead_code)]
    fn parse_border_side_attrs(
        e: &quick_xml::events::BytesStart,
    ) -> (Option<String>, Option<String>) {
        let style = Self::get_attr_str(e, b"style");
        let color = None; // Color comes from nested element
        (style, color)
    }

    /// Parse a color element and return the color string.
    #[allow(dead_code)]
    fn parse_color_element(e: &quick_xml::events::BytesStart) -> Option<String> {
        if let Some(rgb) = Self::get_attr_str(e, b"rgb") {
            Some(format!("#{}", rgb))
        } else {
            Self::get_attr_str(e, b"theme").map(|theme| format!("theme:{}", theme))
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
                    if in_border
                        && (name == b"left"
                            || name == b"right"
                            || name == b"top"
                            || name == b"bottom"
                            || name == b"diagonal")
                    {
                        let mut style: Option<String> = None;
                        let color: Option<String> = None;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"style" {
                                style = Some(String::from_utf8_lossy(&attr.value).to_string());
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
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"rgb" {
                                current_border_color =
                                    Some(format!("#{}", String::from_utf8_lossy(&attr.value)));
                            }
                        }
                    }
                    // Handle numFmt as empty element (self-closing)
                    if name == b"numFmt" {
                        let mut fmt_id: Option<u32> = None;
                        let mut fmt_code: Option<String> = None;
                        for attr in e.attributes().flatten() {
                            let attr_key = attr.key.as_ref();
                            if attr_key == b"numFmtId" {
                                if let Ok(id) = String::from_utf8_lossy(&attr.value).parse::<u32>()
                                {
                                    fmt_id = Some(id);
                                }
                            } else if attr_key == b"formatCode" {
                                fmt_code = Some(String::from_utf8_lossy(&attr.value).to_string());
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
                        for attr in e.attributes().flatten() {
                            let attr_key = attr.key.as_ref();
                            if attr_key == b"numFmtId" {
                                if let Ok(id) = String::from_utf8_lossy(&attr.value).parse::<u32>()
                                {
                                    current_num_fmt_id = Some(id);
                                }
                            } else if attr_key == b"formatCode" {
                                current_num_fmt_code =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
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
                        if prop_name == b"left"
                            || prop_name == b"right"
                            || prop_name == b"top"
                            || prop_name == b"bottom"
                            || prop_name == b"diagonal"
                        {
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
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"style" {
                                    current_border_style =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        } else if prop_name == b"color" && in_border_side.is_some() {
                            // Get color for current border side
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"rgb" {
                                    current_border_color =
                                        Some(format!("#{}", String::from_utf8_lossy(&attr.value)));
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
                    } else if in_border
                        && (name == b"left"
                            || name == b"right"
                            || name == b"top"
                            || name == b"bottom"
                            || name == b"diagonal")
                    {
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

                        for attr in e.attributes().flatten() {
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
                                if let Ok(id) = String::from_utf8_lossy(&attr.value).parse::<u32>()
                                {
                                    if let Some(format) = number_formats.get(&id) {
                                        current_xf.number_format = Some(Arc::from(format.as_str()));
                                    } else if let Some(code) =
                                        StyleRegistry::builtin_num_fmt_code(id)
                                    {
                                        current_xf.number_format = Some(Arc::from(code));
                                    }
                                }
                            }
                        }
                    } else if name == b"alignment" && in_xf {
                        has_alignment = true;
                        current_align = Alignment::default();
                        for attr in e.attributes().flatten() {
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
                                if let Ok(rotation) =
                                    String::from_utf8_lossy(&attr.value).parse::<i32>()
                                {
                                    current_align.text_rotation = Some(rotation);
                                }
                            } else if attr_key == b"shrinkToFit" {
                                current_align.shrink_to_fit =
                                    String::from_utf8_lossy(&attr.value) == "1";
                            } else if attr_key == b"indent" {
                                if let Ok(indent) =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>()
                                {
                                    current_align.indent = Some(indent);
                                }
                            }
                        }
                    } else if name == b"protection" && in_xf {
                        has_protection = true;
                        current_protection = Protection::default();
                        for attr in e.attributes().flatten() {
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
                        for attr in e.attributes().flatten() {
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
                                if let Ok(rotation) =
                                    String::from_utf8_lossy(&attr.value).parse::<i32>()
                                {
                                    current_align.text_rotation = Some(rotation);
                                }
                            } else if attr_key == b"shrinkToFit" {
                                current_align.shrink_to_fit =
                                    String::from_utf8_lossy(&attr.value) == "1";
                            } else if attr_key == b"indent" {
                                if let Ok(indent) =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>()
                                {
                                    current_align.indent = Some(indent);
                                }
                            }
                        }
                    } else if name == b"protection" && in_xf && in_cell_xfs {
                        has_protection = true;
                        current_protection = Protection::default();
                        for attr in e.attributes().flatten() {
                            let attr_key = attr.key.as_ref();
                            if attr_key == b"locked" {
                                current_protection.locked =
                                    String::from_utf8_lossy(&attr.value) == "1";
                            } else if attr_key == b"hidden" {
                                current_protection.hidden =
                                    String::from_utf8_lossy(&attr.value) == "1";
                            }
                        }
                    } else if name == b"xf" && in_cell_xfs {
                        let mut xf = CellStyle::default();
                        for attr in e.attributes().flatten() {
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
                                if let Ok(id) = String::from_utf8_lossy(&attr.value).parse::<u32>()
                                {
                                    if let Some(format) = number_formats.get(&id) {
                                        xf.number_format = Some(Arc::from(format.as_str()));
                                    } else if let Some(code) =
                                        StyleRegistry::builtin_num_fmt_code(id)
                                    {
                                        xf.number_format = Some(Arc::from(code));
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
                    font_id: style
                        .font
                        .as_ref()
                        .and_then(|f| registry.fonts.iter().position(|rf| rf == f))
                        .unwrap_or(0),
                    fill_id: style
                        .fill
                        .as_ref()
                        .and_then(|f| registry.fills.iter().position(|rf| rf == f))
                        .unwrap_or(0),
                    border_id: style
                        .border
                        .as_ref()
                        .and_then(|b| registry.borders.iter().position(|rb| rb == b))
                        .unwrap_or(0),
                    num_fmt_id: style
                        .number_format
                        .as_ref()
                        .and_then(|nf| StyleRegistry::builtin_num_fmt_id(nf))
                        .or_else(|| {
                            style.number_format.as_ref().and_then(|nf| {
                                registry
                                    .num_fmts
                                    .iter()
                                    .find(|(_, code)| code.as_str() == nf.as_ref())
                                    .map(|(id, _)| *id)
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

    /// The smallest a cell can be in the XML: `<c/>`. Used to bound the
    /// up-front reserve by what the sheet could actually contain.
    const MIN_CELL_XML_BYTES: usize = 4;

    /// How many cells to reserve for a sheet whose `<dimension>` claims `ref`.
    ///
    /// `<dimension>` is untrusted: a few-byte `<dimension ref="A1:E1000000"/>`
    /// would otherwise reserve five million entries -- hundreds of megabytes --
    /// before a single cell is read. Cap the estimate by the number of cells the
    /// sheet XML has room for, which cannot be inflated without actually
    /// shipping the bytes.
    fn dimension_reserve(ref_str: &str, sheet_xml_len: usize) -> Option<usize> {
        let estimate = Self::estimate_dimension_cells(ref_str)?;
        let possible = sheet_xml_len / Self::MIN_CELL_XML_BYTES;
        Some(estimate.min(possible)).filter(|cap| *cap > 0)
    }

    /// Apply a frozen `<pane>` element to the worksheet's freeze_panes.
    fn parse_pane_attrs(e: &BytesStart, worksheet: &mut Worksheet) {
        let mut top_left: Option<String> = None;
        let mut frozen = false;
        for attr in e.attributes().flatten() {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.as_ref() {
                b"topLeftCell" => top_left = Some(val.to_string()),
                b"state" => frozen = val == "frozen" || val == "frozenSplit",
                _ => {}
            }
        }
        if frozen {
            worksheet.freeze_panes = top_left;
        }
    }

    /// Apply the `ref` range of a worksheet-level `<autoFilter>` element. The
    /// filter criteria live in child elements; see parse_autofilter_children.
    fn parse_autofilter_attrs(e: &BytesStart, worksheet: &mut Worksheet) {
        for attr in e.attributes().flatten() {
            if attr.key.as_ref() == b"ref" {
                let range = String::from_utf8_lossy(&attr.value).to_string();
                worksheet.auto_filter = Some(AutoFilter::new(range));
            }
        }
    }

    /// Read a single attribute as an owned String.
    fn attr_value(e: &BytesStart, key: &[u8]) -> Option<String> {
        e.attributes()
            .flatten()
            .find(|attr| attr.key.local_name().as_ref() == key)
            .map(|attr| String::from_utf8_lossy(&attr.value).to_string())
    }

    /// True when an attribute is present and not explicitly disabled.
    fn attr_flag(e: &BytesStart, key: &[u8], default: bool) -> bool {
        match Self::attr_value(e, key) {
            Some(v) => v == "1" || v == "true",
            None => default,
        }
    }

    fn parse_filter_operator(value: &str) -> FilterOperator {
        match value {
            "notEqual" => FilterOperator::NotEqual,
            "greaterThan" => FilterOperator::GreaterThan,
            "greaterThanOrEqual" => FilterOperator::GreaterThanOrEqual,
            "lessThan" => FilterOperator::LessThan,
            "lessThanOrEqual" => FilterOperator::LessThanOrEqual,
            _ => FilterOperator::Equal,
        }
    }

    fn parse_dynamic_filter(value: &str) -> DynamicFilterType {
        match value {
            "yesterday" => DynamicFilterType::Yesterday,
            "tomorrow" => DynamicFilterType::Tomorrow,
            "thisWeek" => DynamicFilterType::ThisWeek,
            "nextWeek" => DynamicFilterType::NextWeek,
            "lastWeek" => DynamicFilterType::LastWeek,
            "thisMonth" => DynamicFilterType::ThisMonth,
            "nextMonth" => DynamicFilterType::NextMonth,
            "lastMonth" => DynamicFilterType::LastMonth,
            "thisQuarter" => DynamicFilterType::ThisQuarter,
            "nextQuarter" => DynamicFilterType::NextQuarter,
            "lastQuarter" => DynamicFilterType::LastQuarter,
            "thisYear" => DynamicFilterType::ThisYear,
            "nextYear" => DynamicFilterType::NextYear,
            "lastYear" => DynamicFilterType::LastYear,
            "yearToDate" => DynamicFilterType::YearToDate,
            "aboveAverage" => DynamicFilterType::AboveAverage,
            "belowAverage" => DynamicFilterType::BelowAverage,
            _ => DynamicFilterType::Today,
        }
    }

    /// Parse the children of a worksheet `<autoFilter>` -- the per-column filter
    /// criteria and the sort state -- consuming events through `</autoFilter>`.
    /// Without this a load->save cycle silently clears an active filter.
    fn parse_autofilter_children<R: BufRead>(
        reader: &mut Reader<R>,
        auto_filter: &mut AutoFilter,
    ) -> Result<()> {
        let mut buf = Vec::new();
        let mut column_id: u32 = 0;
        let mut show_button = true;
        let mut values: Vec<String> = Vec::new();
        let mut custom_and = true;
        let mut custom_conditions: Vec<(FilterOperator, String)> = Vec::new();
        let mut filter: Option<FilterType> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                    b"filterColumn" => {
                        column_id = Self::attr_value(&e, b"colId")
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(0);
                        show_button = !Self::attr_flag(&e, b"hiddenButton", false);
                        values.clear();
                        custom_conditions.clear();
                        filter = None;
                    }
                    b"filters" => values.clear(),
                    b"filter" => {
                        if let Some(val) = Self::attr_value(&e, b"val") {
                            values.push(val);
                        }
                    }
                    b"customFilters" => {
                        custom_and = Self::attr_flag(&e, b"and", true);
                        custom_conditions.clear();
                    }
                    b"customFilter" => {
                        let operator = Self::attr_value(&e, b"operator")
                            .map(|v| Self::parse_filter_operator(&v))
                            .unwrap_or(FilterOperator::Equal);
                        let value = Self::attr_value(&e, b"val").unwrap_or_default();
                        custom_conditions.push((operator, value));
                    }
                    b"dynamicFilter" => {
                        let kind = Self::attr_value(&e, b"type")
                            .map(|v| Self::parse_dynamic_filter(&v))
                            .unwrap_or(DynamicFilterType::Today);
                        filter = Some(FilterType::DynamicFilter(kind));
                    }
                    b"top10" => {
                        filter = Some(FilterType::Top10Filter(Top10Filter {
                            top: Self::attr_flag(&e, b"top", true),
                            value: Self::attr_value(&e, b"val")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(10.0),
                            percent: Self::attr_flag(&e, b"percent", false),
                        }));
                    }
                    b"colorFilter" => {
                        filter = Some(FilterType::ColorFilter(ColorFilter {
                            cell_color: Self::attr_flag(&e, b"cellColor", true),
                            color: Self::attr_value(&e, b"dxfId").unwrap_or_default(),
                        }));
                    }
                    b"sortCondition" => {
                        // The writer emits a single-column ref like "B:B"
                        if let Some(reference) = Self::attr_value(&e, b"ref") {
                            let letters: String = reference
                                .chars()
                                .take_while(|c| c.is_ascii_alphabetic())
                                .collect();
                            if let Ok(col) = crate::utils::letter_to_column(&letters) {
                                auto_filter
                                    .sort_by(col - 1, Self::attr_flag(&e, b"descending", false));
                            }
                        }
                    }
                    _ => {}
                },
                Ok(Event::End(e)) => match e.local_name().as_ref() {
                    b"filters" => {
                        if !values.is_empty() {
                            filter = Some(FilterType::Values(std::mem::take(&mut values)));
                        }
                    }
                    b"customFilters" => {
                        let mut conditions = custom_conditions.drain(..);
                        if let Some((operator1, value1)) = conditions.next() {
                            let second = conditions.next();
                            filter = Some(FilterType::Custom(CustomFilter {
                                operator1,
                                value1,
                                and: custom_and,
                                operator2: second.as_ref().map(|(op, _)| op.clone()),
                                value2: second.map(|(_, val)| val),
                            }));
                        }
                    }
                    b"filterColumn" => {
                        if let Some(filter) = filter.take() {
                            auto_filter.columns.push(FilterColumn {
                                column_id,
                                filter,
                                show_button,
                            });
                        }
                    }
                    b"autoFilter" => break,
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(RustypyxlError::ParseError(format!(
                        "XML parsing error in autoFilter: {}",
                        e
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(())
    }

    /// Apply `<pageMargins>` to the worksheet's page setup.
    fn parse_page_margins_attrs(e: &BytesStart, worksheet: &mut Worksheet) {
        let ps = worksheet.page_setup.get_or_insert_with(PageSetup::new);
        for attr in e.attributes().flatten() {
            if let Ok(v) = String::from_utf8_lossy(&attr.value).parse::<f64>() {
                match attr.key.as_ref() {
                    b"left" => ps.margins.left = v,
                    b"right" => ps.margins.right = v,
                    b"top" => ps.margins.top = v,
                    b"bottom" => ps.margins.bottom = v,
                    b"header" => ps.margins.header = v,
                    b"footer" => ps.margins.footer = v,
                    _ => {}
                }
            }
        }
    }

    /// Apply `<pageSetup>` to the worksheet's page setup.
    fn parse_page_setup_attrs(e: &BytesStart, worksheet: &mut Worksheet) {
        let ps = worksheet.page_setup.get_or_insert_with(PageSetup::new);
        for attr in e.attributes().flatten() {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.as_ref() {
                b"paperSize" => {
                    if let Ok(code) = val.parse() {
                        ps.paper_size = PaperSize::from_code(code);
                    }
                }
                b"orientation" => {
                    ps.orientation = if val == "landscape" {
                        Orientation::Landscape
                    } else {
                        Orientation::Portrait
                    };
                }
                b"scale" => {
                    if let Ok(v) = val.parse() {
                        ps.scale = v;
                    }
                }
                b"fitToWidth" => ps.fit_to_width = val.parse().ok(),
                b"fitToHeight" => ps.fit_to_height = val.parse().ok(),
                b"firstPageNumber" => ps.first_page_number = val.parse().ok(),
                b"blackAndWhite" => ps.black_and_white = val == "1" || val == "true",
                b"draft" => ps.draft = val == "1" || val == "true",
                b"horizontalDpi" => ps.horizontal_dpi = val.parse().ok(),
                b"verticalDpi" => ps.vertical_dpi = val.parse().ok(),
                b"copies" => {
                    if let Ok(v) = val.parse() {
                        ps.copies = v;
                    }
                }
                _ => {}
            }
        }
    }

    /// Apply `<printOptions>` to the worksheet's page setup.
    fn parse_print_options_attrs(e: &BytesStart, worksheet: &mut Worksheet) {
        let ps = worksheet.page_setup.get_or_insert_with(PageSetup::new);
        for attr in e.attributes().flatten() {
            let val = String::from_utf8_lossy(&attr.value);
            let on = val == "1" || val == "true";
            match attr.key.as_ref() {
                b"gridLines" => ps.print_gridlines = on,
                b"headings" => ps.print_headings = on,
                b"horizontalCentered" => ps.center_horizontally = on,
                b"verticalCentered" => ps.center_vertically = on,
                _ => {}
            }
        }
    }

    /// Build a DataValidation (and its sqref) from a `<dataValidation>` element.
    /// Boolean attributes default to false when absent, per the schema.
    fn parse_data_validation_attrs(e: &BytesStart) -> (DataValidation, Option<String>) {
        let mut dv = DataValidation {
            allow_blank: false,
            show_error: false,
            show_input: false,
            ..Default::default()
        };
        let mut sqref = None;
        for attr in e.attributes().flatten() {
            let val = String::from_utf8_lossy(&attr.value).to_string();
            let on = val == "1" || val == "true";
            match attr.key.as_ref() {
                b"type" => dv.validation_type = val,
                b"operator" => dv.operator = Some(val),
                b"errorStyle" => dv.error_style = Some(val),
                b"allowBlank" => dv.allow_blank = on,
                b"showErrorMessage" => dv.show_error = on,
                b"showInputMessage" => dv.show_input = on,
                b"errorTitle" => dv.error_title = Some(val),
                b"error" => dv.error_message = Some(val),
                b"promptTitle" => dv.prompt_title = Some(val),
                b"prompt" => dv.prompt_message = Some(val),
                b"sqref" => sqref = Some(val),
                _ => {}
            }
        }
        (dv, sqref)
    }

    /// Insert a parsed data validation, keyed by the first cell of its sqref.
    fn insert_data_validation(
        worksheet: &mut Worksheet,
        mut dv: DataValidation,
        sqref: Option<String>,
    ) {
        if let Some(sq) = sqref {
            let first = sq.split([' ', ':']).next().unwrap_or("").replace('$', "");
            if let Ok((row, col)) = parse_coordinate(&first) {
                dv.sqref = Some(sq);
                worksheet.data_validations.insert((row, col), dv);
            }
        }
    }

    /// Parse a `<color>`-shaped element's attributes into a ConditionalColor.
    fn parse_conditional_color(e: &BytesStart) -> ConditionalColor {
        let mut color = ConditionalColor {
            rgb: None,
            theme: None,
            tint: None,
        };
        for attr in e.attributes().flatten() {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.as_ref() {
                b"rgb" => color.rgb = Some(val.to_string()),
                b"theme" => color.theme = val.parse().ok(),
                b"tint" => color.tint = val.parse().ok(),
                _ => {}
            }
        }
        color
    }

    /// Parse the `<dxfs>` section of styles.xml into differential formats,
    /// indexed by dxfId.
    fn parse_dxfs_xml(data: &[u8]) -> Result<Vec<ConditionalFormat>> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);

        let mut dxfs = Vec::new();
        let mut buf = Vec::new();
        let mut in_dxfs = false;
        let mut current: Option<ConditionalFormat> = None;
        let mut in_font = false;
        let mut in_fill = false;
        let mut in_border = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let is_empty_tag = false; // handled uniformly below
                    let _ = is_empty_tag;
                    let name = e.local_name();
                    let name = name.as_ref();
                    match name {
                        b"dxfs" => in_dxfs = true,
                        b"dxf" if in_dxfs => current = Some(ConditionalFormat::default()),
                        b"font" if current.is_some() => in_font = true,
                        b"fill" if current.is_some() => in_fill = true,
                        b"border" if current.is_some() => in_border = true,
                        b"b" if in_font => {
                            if let Some(fmt) = current.as_mut() {
                                fmt.bold = Some(!Self::attr_is_off(&e));
                            }
                        }
                        b"i" if in_font => {
                            if let Some(fmt) = current.as_mut() {
                                fmt.italic = Some(!Self::attr_is_off(&e));
                            }
                        }
                        b"strike" if in_font => {
                            if let Some(fmt) = current.as_mut() {
                                fmt.strikethrough = Some(!Self::attr_is_off(&e));
                            }
                        }
                        b"u" if in_font => {
                            if let Some(fmt) = current.as_mut() {
                                fmt.underline = Some(!Self::attr_is_off(&e));
                            }
                        }
                        b"color" if in_font => {
                            if let Some(fmt) = current.as_mut() {
                                fmt.font_color = Some(Self::parse_conditional_color(&e));
                            }
                        }
                        b"color" if in_border => {
                            if let Some(fmt) = current.as_mut() {
                                if fmt.border_color.is_none() {
                                    fmt.border_color = Some(Self::parse_conditional_color(&e));
                                }
                            }
                        }
                        b"bgColor" if in_fill => {
                            if let Some(fmt) = current.as_mut() {
                                fmt.fill_color = Some(Self::parse_conditional_color(&e));
                            }
                        }
                        b"fgColor" if in_fill => {
                            if let Some(fmt) = current.as_mut() {
                                if fmt.fill_color.is_none() {
                                    fmt.fill_color = Some(Self::parse_conditional_color(&e));
                                }
                            }
                        }
                        b"numFmt" if current.is_some() => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"formatCode" {
                                    if let (Some(fmt), Ok(code)) =
                                        (current.as_mut(), attr.unescape_value())
                                    {
                                        fmt.number_format = Some(code.to_string());
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(e)) => match e.local_name().as_ref() {
                    b"dxfs" => break,
                    b"dxf" => {
                        if let Some(fmt) = current.take() {
                            dxfs.push(fmt);
                        }
                    }
                    b"font" => in_font = false,
                    b"fill" => in_fill = false,
                    b"border" => in_border = false,
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(RustypyxlError::ParseError(format!(
                        "XML parsing error in dxfs: {}",
                        e
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(dxfs)
    }

    /// True when a toggle element like `<b val="0"/>` or `<u val="none"/>`
    /// explicitly disables the property; a bare `<b/>` enables it.
    fn attr_is_off(e: &BytesStart) -> bool {
        for attr in e.attributes().flatten() {
            if attr.key.as_ref() == b"val" {
                return matches!(attr.value.as_ref(), b"0" | b"false" | b"none");
            }
        }
        false
    }

    /// Parse an xl/tables/tableN.xml part into a Table.
    fn parse_table_xml<R: BufRead>(reader: R) -> Result<Table> {
        let mut reader = Reader::from_reader(reader);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut table = Table::new(0, "", "");
        table.auto_filter = false;
        let mut header_row_count: u32 = 1;
        let mut totals_row_count: u32 = 0;
        // calculatedColumnFormula is a child element of tableColumn, not an
        // attribute of it, so its text arrives in a later event.
        let mut in_calc_formula = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.local_name().as_ref() == b"calculatedColumnFormula" => {
                    in_calc_formula = true;
                }
                Ok(Event::Text(e)) if in_calc_formula => {
                    let text = e.unescape().unwrap_or_default();
                    if let Some(col) = table.columns.last_mut() {
                        col.calculated_column_formula = Some(text.into_owned());
                    }
                }
                Ok(Event::End(e)) if e.local_name().as_ref() == b"calculatedColumnFormula" => {
                    in_calc_formula = false;
                }
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) => match e.local_name().as_ref() {
                    b"table" => {
                        for attr in e.attributes().flatten() {
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match attr.key.as_ref() {
                                b"id" => table.id = val.parse().unwrap_or(0),
                                b"name" => table.name = val,
                                b"displayName" => table.display_name = val,
                                b"ref" => table.range = val,
                                b"headerRowCount" => header_row_count = val.parse().unwrap_or(1),
                                b"totalsRowCount" => totals_row_count = val.parse().unwrap_or(0),
                                b"comment" => table.comment = Some(val),
                                _ => {}
                            }
                        }
                    }
                    b"autoFilter" => table.auto_filter = true,
                    b"tableColumn" => {
                        let mut id = 0u32;
                        let mut name = String::new();
                        let mut totals_label: Option<String> = None;
                        let mut totals_fn: Option<String> = None;
                        let mut formula: Option<String> = None;
                        for attr in e.attributes().flatten() {
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match attr.key.as_ref() {
                                b"id" => id = val.parse().unwrap_or(0),
                                b"name" => name = val,
                                b"totalsRowLabel" => totals_label = Some(val),
                                b"totalsRowFunction" => totals_fn = Some(val),
                                b"calculatedColumnFormula" => formula = Some(val),
                                _ => {}
                            }
                        }
                        let mut col = TableColumn::new(id, &name);
                        col.totals_row_label = totals_label;
                        col.calculated_column_formula = formula;
                        if let Some(f) = totals_fn {
                            col.totals_row_function = match f.as_str() {
                                "average" => TotalsRowFunction::Average,
                                "count" => TotalsRowFunction::Count,
                                "countNums" => TotalsRowFunction::CountNums,
                                "max" => TotalsRowFunction::Max,
                                "min" => TotalsRowFunction::Min,
                                "stdDev" => TotalsRowFunction::StdDev,
                                "sum" => TotalsRowFunction::Sum,
                                "var" => TotalsRowFunction::Var,
                                other => TotalsRowFunction::Custom(other.to_string()),
                            };
                        }
                        table.columns.push(col);
                    }
                    b"tableStyleInfo" => {
                        for attr in e.attributes().flatten() {
                            let val = String::from_utf8_lossy(&attr.value);
                            let on = val == "1" || val == "true";
                            match attr.key.as_ref() {
                                // Keep the exact style name for round-trip fidelity
                                b"name" => table.style = TableStyle::Custom(val.to_string()),
                                b"showFirstColumn" => table.show_first_column = on,
                                b"showLastColumn" => table.show_last_column = on,
                                b"showRowStripes" => table.show_row_stripes = on,
                                b"showColumnStripes" => table.show_column_stripes = on,
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(RustypyxlError::ParseError(format!(
                        "XML parsing error in table: {}",
                        e
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        if table.range.is_empty() {
            return Err(RustypyxlError::InvalidFormat(
                "table part missing ref attribute".to_string(),
            ));
        }
        table.header_row = header_row_count > 0;
        table.totals_row = totals_row_count > 0;
        Ok(table)
    }

    /// Read the `<row>` attributes in a single pass so the result does not
    /// depend on the order the attributes appear in.
    fn parse_row_attrs(e: &quick_xml::events::BytesStart) -> (Option<u32>, Option<f64>) {
        let mut index = None;
        let mut height = None;
        for attr in e.attributes().flatten() {
            match attr.key.as_ref() {
                b"r" => index = String::from_utf8_lossy(&attr.value).parse().ok(),
                b"ht" => height = String::from_utf8_lossy(&attr.value).parse().ok(),
                _ => {}
            }
        }
        (index, height)
    }

    /// Map the internal one-byte cell type to its OOXML `t` attribute. The
    /// codes are a fixed set, so this borrows rather than allocating a String
    /// for every typed cell on the sheet.
    fn data_type_code(cell_type: u8) -> Option<&'static str> {
        match cell_type {
            b's' => Some("s"),
            b'b' => Some("b"),
            b'd' => Some("d"),
            b'f' => Some("str"),
            b'e' => Some("e"),
            b'i' => Some("inlineStr"),
            _ => None,
        }
    }

    /// Read the `<c>` attributes. `r` is optional in OOXML, so the coordinate
    /// is returned as an Option and the caller supplies the implied position.
    fn parse_cell_attrs(
        e: &quick_xml::events::BytesStart,
    ) -> (Option<(u32, u32)>, u8, Option<u32>) {
        let mut coord = None;
        let mut cell_type = 0u8;
        let mut style_id = None;
        for attr in e.attributes().flatten() {
            match attr.key.as_ref() {
                b"r" => coord = parse_coordinate_bytes(&attr.value),
                // Map the full type to a one-byte code; matching on the first
                // byte alone would conflate t="s" (shared string) with t="str"
                // (formula string result).
                b"t" => {
                    cell_type = match attr.value.as_ref() {
                        b"s" => b's',
                        b"str" => b'f',
                        b"b" => b'b',
                        b"d" => b'd',
                        b"e" => b'e',
                        b"inlineStr" => b'i',
                        _ => 0,
                    }
                }
                b"s" => style_id = parse_u32_bytes(&attr.value),
                _ => {}
            }
        }
        (coord, cell_type, style_id)
    }

    fn parse_worksheet_xml<R: BufRead>(
        reader: R,
        shared_strings: &[crate::cell::InternedString],
        styles: &HashMap<u32, Arc<CellStyle>>,
        rels: &HashMap<String, SheetRel>,
        dxfs: &[ConditionalFormat],
        worksheet: &mut Worksheet,
        sheet_xml_len: usize,
    ) -> Result<()> {
        let mut reader = Reader::from_reader(reader);
        // Don't trim text - we need to preserve whitespace in cell values
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut current_row: Option<u32> = None;
        let mut current_col: Option<u32> = None;
        // `r` is optional on both <row> and <c>; when absent the position is
        // implied by the element's index within its parent, so track the next
        // implied row and column as a fallback.
        let mut next_row: u32 = 1;
        let mut next_col: u32 = 1;
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
        let mut current_number_format: Option<crate::cell::InternedString> = None;
        // Raw <v> text of a formula cell, kept verbatim so the cached result
        // round-trips as written rather than being reformatted as an f64.
        let mut current_v_raw: Option<String> = None;
        // True once an inline-string run has contributed text to this cell, so
        // that later <t> runs append instead of replacing.
        let mut inline_runs = false;
        let mut in_cell = false;
        let mut in_v = false;
        let mut in_t = false;
        let mut in_f = false;
        let mut _in_hyperlinks = false;
        let mut current_merge_ref: Option<String> = None;
        let mut hyperlinks: HashMap<(u32, u32), String> = HashMap::new();
        let mut protection: Option<WorksheetProtection> = None;
        let mut reserved_cells = false;
        let mut current_validation: Option<(DataValidation, Option<String>)> = None;
        let mut in_formula1 = false;
        let mut in_formula2 = false;
        // Conditional formatting state
        let mut current_cf: Option<ConditionalFormatting> = None;
        let mut current_cf_rule: Option<ConditionalRule> = None;
        let mut in_cf_formula = false;
        let mut cf_formula_count: u8 = 0;
        // 0 = none, 1 = colorScale, 2 = dataBar, 3 = iconSet
        let mut cf_container: u8 = 0;
        let mut cf_cfvos: Vec<(String, String)> = Vec::new();
        let mut cf_colors: Vec<ConditionalColor> = Vec::new();
        let mut cf_show_value = true;
        let mut cf_icon: Option<IconSet> = None;
        let mut in_odd_header = false;
        let mut in_odd_footer = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) => {
                    // Match on the local name so namespace-prefixed documents
                    // (<x:c>, <x:row>, ...) parse the same as unprefixed ones.
                    let name = e.local_name();
                    let name = name.as_ref();
                    if name == b"sheetProtection" {
                        let mut prot = WorksheetProtection {
                            sheet: true,
                            ..Default::default()
                        };
                        for attr in e.attributes().flatten() {
                            let attr_key = attr.key.as_ref();
                            let attr_value = String::from_utf8_lossy(&attr.value);
                            let value_bool = attr_value == "1";
                            match attr_key {
                                b"password" => prot.password_hash = Some(attr_value.to_string()),
                                b"selectLockedCells" => prot.select_locked_cells = value_bool,
                                b"selectUnlockedCells" => prot.select_unlocked_cells = value_bool,
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
                        protection = Some(prot);
                    } else if name == b"dimension" && !reserved_cells {
                        for attr in e.attributes().flatten() {
                            let attr_key = attr.key;
                            let attr_key = attr_key.as_ref();
                            if attr_key == b"ref" {
                                let ref_str = String::from_utf8_lossy(&attr.value);
                                if let Some(cap) =
                                    Self::dimension_reserve(ref_str.as_ref(), sheet_xml_len)
                                {
                                    worksheet.cells.reserve(cap);
                                    reserved_cells = true;
                                }
                            }
                        }
                    } else if name == b"mergeCell" {
                        for attr in e.attributes().flatten() {
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
                    } else if name == b"hyperlink" {
                        let mut hyperlink_ref: Option<String> = None;
                        let mut location: Option<String> = None;
                        let mut rel_id: Option<String> = None;
                        for attr in e.attributes().flatten() {
                            let attr_key = attr.key.as_ref();
                            if attr_key == b"ref" {
                                hyperlink_ref =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            } else if attr_key == b"location" {
                                location = Some(String::from_utf8_lossy(&attr.value).to_string());
                            } else if attr.key.local_name().as_ref() == b"id" {
                                // r:id pointing into the sheet rels (external URL)
                                rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                        if let Some(ref_coord) = hyperlink_ref {
                            if let Ok((row, col)) = parse_coordinate(&ref_coord) {
                                let url = rel_id
                                    .and_then(|id| rels.get(&id))
                                    .filter(|rel| rel.external)
                                    .map(|rel| rel.target.clone())
                                    .or_else(|| {
                                        location.map(|loc| {
                                            if loc.starts_with('#') {
                                                loc
                                            } else {
                                                format!("#{}", loc)
                                            }
                                        })
                                    });
                                if let Some(url) = url {
                                    hyperlinks.insert((row, col), url);
                                }
                            }
                        }
                    } else if name == b"cfvo" && cf_container != 0 {
                        let mut cfvo_type = String::new();
                        let mut cfvo_val = String::new();
                        for attr in e.attributes().flatten() {
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match attr.key.as_ref() {
                                b"type" => cfvo_type = val,
                                b"val" => cfvo_val = val,
                                _ => {}
                            }
                        }
                        cf_cfvos.push((cfvo_type, cfvo_val));
                    } else if name == b"color" && cf_container != 0 {
                        cf_colors.push(Self::parse_conditional_color(&e));
                    } else if name == b"pane" {
                        Self::parse_pane_attrs(&e, worksheet);
                    } else if name == b"autoFilter" {
                        Self::parse_autofilter_attrs(&e, worksheet);
                    } else if name == b"pageMargins" {
                        Self::parse_page_margins_attrs(&e, worksheet);
                    } else if name == b"pageSetup" {
                        Self::parse_page_setup_attrs(&e, worksheet);
                    } else if name == b"printOptions" {
                        Self::parse_print_options_attrs(&e, worksheet);
                    } else if name == b"dataValidation" {
                        // Self-closing form (no formula children)
                        let (dv, sqref) = Self::parse_data_validation_attrs(&e);
                        Self::insert_data_validation(worksheet, dv, sqref);
                    } else if name == b"col" {
                        let mut col_min: Option<u32> = None;
                        let mut col_max: Option<u32> = None;
                        let mut width: Option<f64> = None;
                        for attr in e.attributes().flatten() {
                            let attr_key = attr.key.as_ref();
                            if attr_key == b"min" {
                                if let Ok(num) = String::from_utf8_lossy(&attr.value).parse::<u32>()
                                {
                                    col_min = Some(num);
                                }
                            } else if attr_key == b"max" {
                                if let Ok(num) = String::from_utf8_lossy(&attr.value).parse::<u32>()
                                {
                                    col_max = Some(num);
                                }
                            } else if attr_key == b"width" {
                                if let Ok(w) = String::from_utf8_lossy(&attr.value).parse::<f64>() {
                                    width = Some(w);
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
                    } else if name == b"row" {
                        // A row with no cells still carries formatting, e.g.
                        // <row r="3" ht="20" customHeight="1"/>
                        let (index, height) = Self::parse_row_attrs(&e);
                        let row = index.unwrap_or(next_row);
                        next_row = row.saturating_add(1);
                        next_col = 1;
                        if let Some(height) = height {
                            worksheet.set_row_height(row, height);
                        }
                    } else if name == b"c" {
                        // Handle self-closing cell elements like <c r="A1" t="inlineStr" />
                        // These are typically empty cells but with a specific type (e.g., empty string)
                        let (coord, cell_type, style_id) = Self::parse_cell_attrs(&e);
                        let cell_row = coord.map(|(r, _)| r).or(current_row);
                        let cell_col = Some(coord.map_or(next_col, |(_, c)| c));
                        if let Some(col) = cell_col {
                            next_col = col.saturating_add(1);
                        }

                        if let (Some(row), Some(col)) = (cell_row, cell_col) {
                            // If it's marked as a string type (inline or shared), treat as empty string
                            // Otherwise it's truly empty
                            let cell_value = if matches!(cell_type, b'i' | b's' | b'f') {
                                CellValue::String(std::sync::Arc::from(""))
                            } else {
                                CellValue::Empty
                            };

                            let style = style_id.and_then(|id| styles.get(&id).cloned());
                            let num_format = style.as_ref().and_then(|s| s.number_format.clone());
                            let data_type_str = Self::data_type_code(cell_type);

                            let cell_data = CellData {
                                value: cell_value,
                                style,
                                style_index: style_id,
                                number_format: num_format,
                                data_type: data_type_str,
                                ..Default::default()
                            };

                            worksheet.set_cell_data(row, col, cell_data);
                        }
                    }
                }
                Ok(Event::Start(e)) => {
                    let name = e.local_name();
                    let name = name.as_ref();

                    if name == b"dimension" && !reserved_cells {
                        for attr in e.attributes().flatten() {
                            let attr_key = attr.key;
                            let attr_key = attr_key.as_ref();
                            if attr_key == b"ref" {
                                let ref_str = String::from_utf8_lossy(&attr.value);
                                if let Some(cap) =
                                    Self::dimension_reserve(ref_str.as_ref(), sheet_xml_len)
                                {
                                    worksheet.cells.reserve(cap);
                                    reserved_cells = true;
                                }
                            }
                        }
                    } else if name == b"row" {
                        let (index, height) = Self::parse_row_attrs(&e);
                        let row = index.unwrap_or(next_row);
                        current_row = Some(row);
                        next_row = row.saturating_add(1);
                        next_col = 1;
                        if let Some(height) = height {
                            worksheet.set_row_height(row, height);
                        }
                    } else if name == b"c" {
                        in_cell = true;
                        current_value = None;
                        current_formula = None;
                        current_number_format = None;
                        current_v_raw = None;
                        inline_runs = false;

                        // cell_type 0 = number, the default when t is absent
                        let (coord, cell_type, style_id) = Self::parse_cell_attrs(&e);
                        current_type = cell_type;
                        current_style_id = style_id;
                        if let Some((row, _)) = coord {
                            current_row = Some(row);
                        }
                        let col = coord.map_or(next_col, |(_, c)| c);
                        current_col = Some(col);
                        next_col = col.saturating_add(1);
                    } else if name == b"v" {
                        in_v = true;
                    } else if name == b"t" {
                        in_t = true;
                    } else if name == b"f" {
                        in_f = true;
                    } else if name == b"mergeCell" {
                        for attr in e.attributes().flatten() {
                            let attr_key = attr.key.as_ref();
                            if attr_key == b"ref" {
                                current_merge_ref =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    } else if name == b"conditionalFormatting" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"sqref" {
                                current_cf = Some(ConditionalFormatting::new(
                                    String::from_utf8_lossy(&attr.value).to_string(),
                                ));
                            }
                        }
                    } else if name == b"cfRule" && current_cf.is_some() {
                        let mut rule = ConditionalRule::default();
                        for attr in e.attributes().flatten() {
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            let on = val == "1" || val == "true";
                            match attr.key.as_ref() {
                                b"type" => {
                                    if let Some(t) = ConditionalFormatType::from_xml(&val) {
                                        rule.rule_type = t;
                                    }
                                }
                                b"dxfId" => {
                                    rule.format = val
                                        .parse::<usize>()
                                        .ok()
                                        .and_then(|id| dxfs.get(id))
                                        .cloned();
                                }
                                b"priority" => rule.priority = val.parse().unwrap_or(1),
                                b"operator" => rule.operator = ConditionalOperator::from_xml(&val),
                                b"stopIfTrue" => rule.stop_if_true = on,
                                b"text" => rule.text = Some(val),
                                b"rank" => rule.rank = val.parse().ok(),
                                b"percent" => rule.percent = on,
                                b"bottom" => rule.bottom = on,
                                b"aboveAverage" => rule.above_average = on,
                                b"equalAverage" => rule.equal_average = on,
                                b"stdDev" => rule.std_dev = val.parse().ok(),
                                b"timePeriod" => rule.time_period = Some(val),
                                _ => {}
                            }
                        }
                        cf_formula_count = 0;
                        current_cf_rule = Some(rule);
                    } else if name == b"formula" && current_cf_rule.is_some() {
                        in_cf_formula = true;
                    } else if name == b"headerFooter" {
                        let ps = worksheet.page_setup.get_or_insert_with(PageSetup::new);
                        for attr in e.attributes().flatten() {
                            let val = String::from_utf8_lossy(&attr.value);
                            let on = val == "1" || val == "true";
                            match attr.key.as_ref() {
                                b"differentOddEven" => ps.header_footer.different_odd_even = on,
                                b"differentFirst" => ps.header_footer.different_first = on,
                                _ => {}
                            }
                        }
                    } else if name == b"oddHeader" {
                        in_odd_header = true;
                    } else if name == b"oddFooter" {
                        in_odd_footer = true;
                    } else if name == b"colorScale" && current_cf_rule.is_some() {
                        cf_container = 1;
                        cf_cfvos.clear();
                        cf_colors.clear();
                    } else if name == b"dataBar" && current_cf_rule.is_some() {
                        cf_container = 2;
                        cf_cfvos.clear();
                        cf_colors.clear();
                        cf_show_value = true;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"showValue" {
                                cf_show_value = !matches!(attr.value.as_ref(), b"0" | b"false");
                            }
                        }
                    } else if name == b"iconSet" && current_cf_rule.is_some() {
                        cf_container = 3;
                        cf_cfvos.clear();
                        cf_colors.clear();
                        let mut icon = IconSet::new(IconSetStyle::ThreeTrafficLights);
                        for attr in e.attributes().flatten() {
                            let val = String::from_utf8_lossy(&attr.value);
                            match attr.key.as_ref() {
                                b"iconSet" => {
                                    if let Some(style) = IconSetStyle::from_xml(&val) {
                                        icon.style = style;
                                    }
                                }
                                b"showValue" => {
                                    icon.show_value = !matches!(val.as_ref(), "0" | "false")
                                }
                                b"reverse" => icon.reverse = val == "1" || val == "true",
                                _ => {}
                            }
                        }
                        cf_icon = Some(icon);
                    } else if name == b"autoFilter" {
                        // Start form: the criteria live in the child elements
                        Self::parse_autofilter_attrs(&e, worksheet);
                        if let Some(auto_filter) = worksheet.auto_filter.as_mut() {
                            Self::parse_autofilter_children(&mut reader, auto_filter)?;
                        }
                    } else if name == b"dataValidation" {
                        current_validation = Some(Self::parse_data_validation_attrs(&e));
                    } else if name == b"formula1" {
                        in_formula1 = current_validation.is_some();
                    } else if name == b"formula2" {
                        in_formula2 = current_validation.is_some();
                    } else if name == b"col" {
                        let mut col_min: Option<u32> = None;
                        let mut col_max: Option<u32> = None;
                        let mut width: Option<f64> = None;
                        for attr in e.attributes().flatten() {
                            let attr_key = attr.key.as_ref();
                            if attr_key == b"min" {
                                if let Ok(num) = String::from_utf8_lossy(&attr.value).parse::<u32>()
                                {
                                    col_min = Some(num);
                                }
                            } else if attr_key == b"max" {
                                if let Ok(num) = String::from_utf8_lossy(&attr.value).parse::<u32>()
                                {
                                    col_max = Some(num);
                                }
                            } else if attr_key == b"width" {
                                if let Ok(w) = String::from_utf8_lossy(&attr.value).parse::<f64>() {
                                    width = Some(w);
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
                        // <f> precedes <v> in the schema, so a formula seen by now
                        // means this <v> is a cached result. Keep it verbatim: a
                        // round-trip through f64 would rewrite "5" as "5.0". Shared
                        // indices still need resolving, so leave those to the parse.
                        if current_formula.is_some() && current_type != b's' {
                            current_v_raw = Some(text.to_string());
                        }
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
                            // Formula string results and error values are literal text
                            b'f' | b'e' => Some(TempValue::String(text.into_owned())),
                            _ => {
                                // Number (default) - try fast f64 parsing
                                match parse_f64_bytes(text.as_bytes()) {
                                    Some(n) => Some(TempValue::Number(n)),
                                    None => Some(TempValue::String(text.into_owned())),
                                }
                            }
                        };
                    } else if in_t && in_cell {
                        // Rich-text inline strings split their content across
                        // <is><r><t> runs; concatenate them rather than keeping
                        // only the last one.
                        match current_value.as_mut() {
                            Some(TempValue::String(s)) if inline_runs => s.push_str(&text),
                            _ => {
                                current_value = Some(TempValue::String(text.into_owned()));
                                inline_runs = true;
                            }
                        }
                    } else if in_f && in_cell {
                        current_formula = Some(text.to_string());
                    } else if in_formula1 {
                        if let Some((dv, _)) = current_validation.as_mut() {
                            dv.formula1 = Some(text.to_string());
                        }
                    } else if in_formula2 {
                        if let Some((dv, _)) = current_validation.as_mut() {
                            dv.formula2 = Some(text.to_string());
                        }
                    } else if in_cf_formula {
                        if let Some(rule) = current_cf_rule.as_mut() {
                            if cf_formula_count == 0 {
                                rule.formula1 = Some(text.to_string());
                            } else {
                                rule.formula2 = Some(text.to_string());
                            }
                        }
                    } else if in_odd_header || in_odd_footer {
                        let section = crate::pagesetup::HeaderFooterSection::parse_encoded(&text);
                        let ps = worksheet.page_setup.get_or_insert_with(PageSetup::new);
                        if in_odd_header {
                            ps.header_footer.odd_header = Some(section);
                        } else {
                            ps.header_footer.odd_footer = Some(section);
                        }
                    }
                }
                Ok(Event::End(e)) => {
                    let name = e.local_name();
                    let name = name.as_ref();

                    if name == b"formula" && in_cf_formula {
                        in_cf_formula = false;
                        cf_formula_count = cf_formula_count.saturating_add(1);
                    } else if name == b"colorScale" {
                        if let Some(rule) = current_cf_rule.as_mut() {
                            if cf_colors.len() >= 2 && cf_cfvos.len() >= 2 {
                                let three = cf_colors.len() >= 3 && cf_cfvos.len() >= 3;
                                let last = cf_cfvos.len() - 1;
                                let opt = |v: &str| {
                                    if v.is_empty() {
                                        None
                                    } else {
                                        Some(v.to_string())
                                    }
                                };
                                rule.color_scale = Some(ColorScale {
                                    min_color: cf_colors[0].clone(),
                                    mid_color: three.then(|| cf_colors[1].clone()),
                                    max_color: cf_colors[cf_colors.len() - 1].clone(),
                                    min_type: cf_cfvos[0].0.clone(),
                                    min_value: opt(&cf_cfvos[0].1),
                                    mid_type: three.then(|| cf_cfvos[1].0.clone()),
                                    mid_value: if three { opt(&cf_cfvos[1].1) } else { None },
                                    max_type: cf_cfvos[last].0.clone(),
                                    max_value: opt(&cf_cfvos[last].1),
                                });
                            }
                        }
                        cf_container = 0;
                    } else if name == b"dataBar" {
                        if let Some(rule) = current_cf_rule.as_mut() {
                            let mut db = DataBar::new();
                            db.show_value = cf_show_value;
                            if let Some((t, v)) = cf_cfvos.first() {
                                db.min_type = t.clone();
                                db.min_value = (!v.is_empty()).then(|| v.clone());
                            }
                            if let Some((t, v)) = cf_cfvos.get(1) {
                                db.max_type = t.clone();
                                db.max_value = (!v.is_empty()).then(|| v.clone());
                            }
                            if let Some(color) = cf_colors.first() {
                                db.fill_color = color.clone();
                            }
                            rule.data_bar = Some(db);
                        }
                        cf_container = 0;
                    } else if name == b"iconSet" {
                        if let Some(rule) = current_cf_rule.as_mut() {
                            if let Some(mut icon) = cf_icon.take() {
                                icon.thresholds = std::mem::take(&mut cf_cfvos);
                                rule.icon_set = Some(icon);
                            }
                        }
                        cf_container = 0;
                    } else if name == b"cfRule" {
                        if let (Some(cf), Some(rule)) =
                            (current_cf.as_mut(), current_cf_rule.take())
                        {
                            cf.rules.push(rule);
                        }
                    } else if name == b"conditionalFormatting" {
                        if let Some(cf) = current_cf.take() {
                            if !cf.rules.is_empty() {
                                worksheet.add_conditional_formatting(cf);
                            }
                        }
                    } else if name == b"oddHeader" {
                        in_odd_header = false;
                    } else if name == b"oddFooter" {
                        in_odd_footer = false;
                    } else if name == b"formula1" {
                        in_formula1 = false;
                    } else if name == b"formula2" {
                        in_formula2 = false;
                    } else if name == b"dataValidation" {
                        if let Some((dv, sqref)) = current_validation.take() {
                            Self::insert_data_validation(worksheet, dv, sqref);
                        }
                        in_formula1 = false;
                        in_formula2 = false;
                    } else if name == b"hyperlinks" {
                        _in_hyperlinks = false;
                        for ((row, col), url) in &hyperlinks {
                            if let Some(cell_data) = worksheet.cells.get_mut(&cell_key(*row, *col))
                            {
                                cell_data.hyperlink = Some(url.clone());
                            } else {
                                let cell_data = CellData {
                                    value: CellValue::Empty,
                                    hyperlink: Some(url.clone()),
                                    ..Default::default()
                                };
                                worksheet.set_cell_data(*row, *col, cell_data);
                            }
                        }
                    } else if name == b"c" {
                        if let (Some(row), Some(col)) = (current_row, current_col) {
                            let mut cached_formula_value: Option<String> = None;
                            let cell_value = if let Some(formula) = current_formula.take() {
                                // Preserve the cached <v> so a save doesn't
                                // blank the cell in viewers that don't recalc
                                let parsed = current_value.take().map(|v| match v {
                                    TempValue::SharedIdx(idx) => shared_strings
                                        .get(idx)
                                        .map(|s| s.to_string())
                                        .unwrap_or_default(),
                                    TempValue::Bool(b) => (if b { "1" } else { "0" }).to_string(),
                                    TempValue::Number(n) => {
                                        let mut buf = ryu::Buffer::new();
                                        buf.format(n).to_string()
                                    }
                                    TempValue::Date(d) => d,
                                    TempValue::String(s) => s,
                                });
                                cached_formula_value = current_v_raw.take().or(parsed);
                                CellValue::Formula(formula)
                            } else if let Some(value) = current_value.take() {
                                match value {
                                    TempValue::SharedIdx(idx) => {
                                        if idx < shared_strings.len() {
                                            CellValue::String(shared_strings[idx].clone())
                                        } else {
                                            // A dangling index means a corrupt file; an empty
                                            // string is less misleading than fabricating the
                                            // index number as cell text.
                                            CellValue::String(std::sync::Arc::from(""))
                                        }
                                    }
                                    TempValue::Bool(b) => CellValue::Boolean(b),
                                    TempValue::Number(n) => CellValue::Number(n),
                                    TempValue::Date(d) => CellValue::Date(d),
                                    TempValue::String(s) => {
                                        CellValue::String(std::sync::Arc::from(s))
                                    }
                                }
                            } else {
                                // If it was marked as a string type but has no value,
                                // treat it as an empty string (openpyxl writes empty strings this way)
                                if matches!(current_type, b'i' | b's' | b'f') {
                                    CellValue::String(std::sync::Arc::from(""))
                                } else {
                                    CellValue::Empty
                                }
                            };

                            let style = current_style_id.and_then(|id| styles.get(&id).cloned());

                            let num_format = current_number_format
                                .take()
                                .or_else(|| style.as_ref().and_then(|s| s.number_format.clone()));

                            let data_type_str = Self::data_type_code(current_type);

                            let cell_data = CellData {
                                value: cell_value,
                                style,
                                style_index: current_style_id,
                                number_format: num_format,
                                data_type: data_type_str,
                                cached_formula_value,
                                ..Default::default()
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
                        for attr in e.attributes().flatten() {
                            let attr_key = attr.key.as_ref();
                            if attr_key == b"ref" {
                                current_cell_ref =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
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

        let (sheets, _, _, _) = Workbook::parse_workbook_xml(Cursor::new(workbook_xml)).unwrap();

        assert_eq!(sheets.len(), 2);
        assert_eq!(
            sheets[0],
            (
                "Data".to_string(),
                8,
                "rId1".to_string(),
                SheetVisibility::Visible
            )
        );
        assert_eq!(
            sheets[1],
            (
                "Summary".to_string(),
                2,
                "rId2".to_string(),
                SheetVisibility::Visible
            )
        );
    }

    /// `<dimension>` is untrusted. A few-byte ref claiming five million cells
    /// must not make us reserve five million entries: the reserve is bounded by
    /// what the sheet XML actually has room for.
    #[test]
    fn test_dimension_reserve_is_bounded_by_the_sheet_size() {
        // The attack: a tiny sheet whose dimension claims 5M cells
        let tiny_sheet_len = 200;
        let cap = Workbook::dimension_reserve("A1:E1000000", tiny_sheet_len);
        assert_eq!(
            cap,
            Some(tiny_sheet_len / Workbook::MIN_CELL_XML_BYTES),
            "a 200-byte sheet cannot hold more than 50 cells"
        );

        // A dimension larger than the hard cap is still rejected outright
        assert_eq!(
            Workbook::dimension_reserve("A1:XFD1048576", 10_000_000),
            None
        );

        // An honest dimension in a sheet big enough to hold it reserves in full
        assert_eq!(Workbook::dimension_reserve("A1:J100", 100_000), Some(1000));

        // An empty sheet reserves nothing rather than zero-capacity churn
        assert_eq!(Workbook::dimension_reserve("A1:J100", 0), None);
    }

    /// The active tab must follow the sheet it pointed at, not the index.
    #[test]
    fn test_remove_sheet_tracks_the_active_tab() {
        let mut wb = Workbook::new();
        for name in ["A", "B", "C"] {
            wb.create_sheet(Some(name.to_string())).unwrap();
        }
        wb.active_sheet = 1; // B

        // Removing a sheet before the active one keeps the same sheet active
        wb.remove_sheet("A").unwrap();
        assert_eq!(wb.active_sheet, 0);
        assert_eq!(wb.sheet_names[wb.active_sheet], "B");

        // Removing one after the active sheet leaves it alone
        wb.remove_sheet("C").unwrap();
        assert_eq!(wb.active_sheet, 0);
        assert_eq!(wb.sheet_names[wb.active_sheet], "B");
    }

    /// Removing the active sheet itself leaves the index in place, so it lands
    /// on the next sheet -- and clamps when there is no next sheet.
    #[test]
    fn test_remove_active_sheet_lands_on_the_next_one() {
        let mut wb = Workbook::new();
        for name in ["A", "B", "C"] {
            wb.create_sheet(Some(name.to_string())).unwrap();
        }
        wb.active_sheet = 1;

        wb.remove_sheet("B").unwrap();
        assert_eq!(wb.sheet_names[wb.active_sheet], "C");

        wb.remove_sheet("C").unwrap();
        assert_eq!(wb.sheet_names[wb.active_sheet], "A", "clamped to the end");

        wb.remove_sheet("A").unwrap();
        assert_eq!(wb.active_sheet, 0, "no sheets left");
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

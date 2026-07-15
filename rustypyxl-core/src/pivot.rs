//! Read-only pivot-table model. Pivot tables are preserved verbatim on save
//! (see [`crate::workbook::PivotArtifacts`]); this module parses those raw parts
//! into a structured view -- source range, cache fields, and the field
//! placements (row/column/data/page areas) -- matching openpyxl's read support.
//! It does not build or write pivot tables.

use crate::workbook::{resolve_rel_target, PivotArtifacts};
use quick_xml::events::Event;
use std::collections::HashMap;

/// A data field in a pivot table (the values area).
#[derive(Clone, Debug, PartialEq)]
pub struct PivotDataField {
    /// Display name, e.g. "Sum of Sales".
    pub name: String,
    /// The cache field it aggregates.
    pub source_field: String,
    /// Aggregation: sum, count, average, max, min, product, countNums, stdDev,
    /// stdDevp, var, varp. Excel's default when unspecified is "sum".
    pub subtotal: String,
}

/// A read-only view of one pivot table.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct PivotTableInfo {
    /// Pivot table name.
    pub name: String,
    /// The cache id it draws from, if declared.
    pub cache_id: Option<u32>,
    /// The cell range the pivot occupies on its sheet, e.g. "A3:D12".
    pub location: Option<String>,
    /// Source data sheet name.
    pub source_sheet: Option<String>,
    /// Source data range, e.g. "A1:D100".
    pub source_ref: Option<String>,
    /// The cache field names, in source-column order.
    pub cache_fields: Vec<String>,
    /// Fields placed in the row area (a data placeholder shows as "Values").
    pub row_fields: Vec<String>,
    /// Fields placed in the column area.
    pub col_fields: Vec<String>,
    /// Fields placed in the report-filter (page) area.
    pub page_fields: Vec<String>,
    /// Fields in the values area with their aggregation.
    pub data_fields: Vec<PivotDataField>,
}

/// Parse every pivot table in a workbook's preserved artifacts into a read-only
/// model. Returns an empty vector when there are no pivot tables.
pub fn parse_pivot_tables(artifacts: &PivotArtifacts) -> Vec<PivotTableInfo> {
    let by_path: HashMap<&str, &[u8]> = artifacts
        .parts
        .iter()
        .map(|(p, b)| (p.as_str(), b.as_slice()))
        .collect();

    let mut out = Vec::new();
    // Deterministic order: iterate the pivotTable parts by path.
    let mut table_paths: Vec<&str> = by_path
        .keys()
        .copied()
        .filter(|p| p.starts_with("xl/pivotTables/pivotTable") && p.ends_with(".xml"))
        .collect();
    table_paths.sort_unstable();

    for path in table_paths {
        let bytes = by_path[path];
        let raw = parse_pivot_table_xml(bytes);

        // Resolve this pivot table's cache definition through its .rels part.
        let rels_path = table_rels_path(path);
        let (source_sheet, source_ref, cache_fields) = by_path
            .get(rels_path.as_str())
            .and_then(|rels| first_cache_definition_target(rels))
            .map(|target| resolve_rel_target(path, &target))
            .and_then(|cache_path| by_path.get(cache_path.as_str()).copied())
            .map(parse_cache_definition_xml)
            .unwrap_or_default();

        out.push(combine(raw, source_sheet, source_ref, cache_fields));
    }
    out
}

/// The .rels path for a pivotTable part.
fn table_rels_path(table_path: &str) -> String {
    match table_path.rfind('/') {
        Some(idx) => format!(
            "{}/_rels/{}.rels",
            &table_path[..idx],
            &table_path[idx + 1..]
        ),
        None => format!("_rels/{}.rels", table_path),
    }
}

/// Read the pivotCacheDefinition target from a pivotTable .rels part.
fn first_cache_definition_target(rels_xml: &[u8]) -> Option<String> {
    let mut reader = quick_xml::Reader::from_reader(rels_xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if e.local_name().as_ref() == b"Relationship" =>
            {
                let (mut typ, mut target) = (None, None);
                for attr in e.attributes().flatten() {
                    let val = attr.unescape_value().ok().map(|v| v.into_owned());
                    match attr.key.local_name().as_ref() {
                        b"Type" => typ = val,
                        b"Target" => target = val,
                        _ => {}
                    }
                }
                if typ.is_some_and(|t| t.ends_with("pivotCacheDefinition")) {
                    return target;
                }
            }
            Ok(Event::Eof) | Err(_) => return None,
            _ => {}
        }
        buf.clear();
    }
}

/// Extract the source sheet/range and the ordered cache field names from a
/// pivotCacheDefinition part.
fn parse_cache_definition_xml(xml: &[u8]) -> (Option<String>, Option<String>, Vec<String>) {
    let mut reader = quick_xml::Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut sheet = None;
    let mut reference = None;
    let mut fields = Vec::new();
    let mut in_cache_fields = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"worksheetSource" => {
                    for attr in e.attributes().flatten() {
                        let val = attr.unescape_value().ok().map(|v| v.into_owned());
                        match attr.key.local_name().as_ref() {
                            b"sheet" => sheet = val,
                            b"ref" => reference = val,
                            _ => {}
                        }
                    }
                }
                b"cacheFields" => in_cache_fields = true,
                b"cacheField" if in_cache_fields => {
                    let name = e
                        .attributes()
                        .flatten()
                        .find(|a| a.key.local_name().as_ref() == b"name")
                        .and_then(|a| a.unescape_value().ok().map(|v| v.into_owned()))
                        .unwrap_or_default();
                    fields.push(name);
                }
                _ => {}
            },
            Ok(Event::End(e)) if e.local_name().as_ref() == b"cacheFields" => {
                in_cache_fields = false;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    (sheet, reference, fields)
}

/// The placement of a pivot field, as read from the pivotTable part before cache
/// field names are joined in.
#[derive(Default)]
struct RawPivotTable {
    name: String,
    cache_id: Option<u32>,
    location: Option<String>,
    /// Row-area field indices (a `-2` is the data-values placeholder).
    row_x: Vec<i32>,
    col_x: Vec<i32>,
    /// Report-filter field indices.
    page_fld: Vec<i32>,
    /// (display name, source field index, subtotal) for each data field.
    data: Vec<(String, i32, String)>,
}

/// Which `<*Fields>` container the parser is currently inside.
enum Container {
    None,
    Row,
    Col,
    Page,
}

/// Parse a pivotTable part into raw field placements.
fn parse_pivot_table_xml(xml: &[u8]) -> RawPivotTable {
    let mut reader = quick_xml::Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut raw = RawPivotTable::default();
    let mut container = Container::None;

    let attr = |e: &quick_xml::events::BytesStart, key: &[u8]| -> Option<String> {
        e.attributes()
            .flatten()
            .find(|a| a.key.local_name().as_ref() == key)
            .and_then(|a| a.unescape_value().ok().map(|v| v.into_owned()))
    };

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"pivotTableDefinition" => {
                    raw.name = attr(&e, b"name").unwrap_or_default();
                    raw.cache_id = attr(&e, b"cacheId").and_then(|v| v.parse().ok());
                }
                b"location" => raw.location = attr(&e, b"ref"),
                b"rowFields" => container = Container::Row,
                b"colFields" => container = Container::Col,
                b"pageFields" => container = Container::Page,
                b"field" => {
                    if let Some(x) = attr(&e, b"x").and_then(|v| v.parse::<i32>().ok()) {
                        match container {
                            Container::Row => raw.row_x.push(x),
                            Container::Col => raw.col_x.push(x),
                            _ => {}
                        }
                    }
                }
                b"pageField" => {
                    if let Some(fld) = attr(&e, b"fld").and_then(|v| v.parse::<i32>().ok()) {
                        raw.page_fld.push(fld);
                    }
                }
                b"dataField" => {
                    let name = attr(&e, b"name").unwrap_or_default();
                    let fld = attr(&e, b"fld")
                        .and_then(|v| v.parse::<i32>().ok())
                        .unwrap_or(-1);
                    let subtotal = attr(&e, b"subtotal").unwrap_or_else(|| "sum".to_string());
                    raw.data.push((name, fld, subtotal));
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"rowFields" | b"colFields" | b"pageFields" => container = Container::None,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    raw
}

// ---------- creation (phase 3) ----------

const NS_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

/// A request to build a new pivot table's parts. Indices are 0-based positions
/// into `field_names`; `col_indices` may include `-2`, Excel's data-values
/// placeholder.
pub struct PivotBuildRequest<'a> {
    /// Cache id, shared between the workbook `<pivotCache>` and the pivotTable.
    pub cache_id: u32,
    /// Part-file number for the cache (pivotCacheDefinition{n}, records{n}).
    pub cache_num: u32,
    /// Part-file number for the pivotTable.
    pub table_num: u32,
    /// Pivot table name.
    pub name: &'a str,
    /// Source data sheet name.
    pub source_sheet: &'a str,
    /// Source data range, e.g. "A1:C100".
    pub source_ref: &'a str,
    /// Where the pivot is placed, e.g. "F1:J20".
    pub location_ref: &'a str,
    /// The source header field names, in column order.
    pub field_names: &'a [String],
    /// Row-area field indices.
    pub row_indices: &'a [i32],
    /// Column-area field indices (may include -2 for the data placeholder).
    pub col_indices: &'a [i32],
    /// (field index, display name, subtotal) for each data field.
    pub data_fields: &'a [(usize, String, String)],
    /// The workbook relationship id placeholder for the cache (phase-1 renumbers).
    pub cache_rel_id: &'a str,
    /// The sheet relationship id for the pivotTable.
    pub sheet_rel_id: &'a str,
}

/// The parts and wiring produced for a new pivot table.
pub struct BuiltPivot {
    /// The five package parts as (path, bytes).
    pub parts: Vec<(String, Vec<u8>)>,
    /// The `<pivotCache …/>` child to add to the workbook `<pivotCaches>`.
    pub caches_child: String,
    /// The workbook.xml.rels entry as (id, target).
    pub workbook_rel: (String, String),
    /// The source sheet's rels entry as (id, type, target).
    pub sheet_rel: (String, String, String),
}

fn escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

/// Generate the parts for a new pivot table. Excel rebuilds the cache from the
/// source on open (`refreshOnLoad`), so the records part is empty and pivot
/// field items are omitted.
pub fn build_pivot(req: PivotBuildRequest) -> BuiltPivot {
    let cache_def_path = format!("xl/pivotCache/pivotCacheDefinition{}.xml", req.cache_num);
    let cache_rec_path = format!("xl/pivotCache/pivotCacheRecords{}.xml", req.cache_num);
    let cache_rels_path = format!(
        "xl/pivotCache/_rels/pivotCacheDefinition{}.xml.rels",
        req.cache_num
    );
    let table_path = format!("xl/pivotTables/pivotTable{}.xml", req.table_num);
    let table_rels_path = format!("xl/pivotTables/_rels/pivotTable{}.xml.rels", req.table_num);

    // cacheFields
    let mut cache_fields = String::new();
    for name in req.field_names {
        cache_fields.push_str(&format!(
            r#"<cacheField name="{}" numFmtId="0"><sharedItems/></cacheField>"#,
            escape(name)
        ));
    }
    let cache_def = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<pivotCacheDefinition xmlns="{m}" xmlns:r="{r}" r:id="rId1" refreshOnLoad="1" recordCount="0">"#,
            r#"<cacheSource type="worksheet"><worksheetSource ref="{sref}" sheet="{sheet}"/></cacheSource>"#,
            r#"<cacheFields count="{fcount}">{fields}</cacheFields>"#,
            r#"</pivotCacheDefinition>"#
        ),
        m = NS_MAIN,
        r = NS_REL,
        sref = escape(req.source_ref),
        sheet = escape(req.source_sheet),
        fcount = req.field_names.len(),
        fields = cache_fields
    );

    let cache_rec = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheRecords xmlns="{}" count="0"/>"#,
        NS_MAIN
    );

    let cache_rels = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<Relationships xmlns="{p}">"#,
            r#"<Relationship Id="rId1" Type="{r}/pivotCacheRecords" Target="pivotCacheRecords{n}.xml"/>"#,
            r#"</Relationships>"#
        ),
        p = NS_PKG_REL,
        r = NS_REL,
        n = req.cache_num
    );

    // pivotFields: one per source field, tagged by placement.
    let row_set: std::collections::HashSet<i32> = req.row_indices.iter().copied().collect();
    let col_set: std::collections::HashSet<i32> = req.col_indices.iter().copied().collect();
    let data_set: std::collections::HashSet<usize> =
        req.data_fields.iter().map(|(i, _, _)| *i).collect();
    let mut pivot_fields = String::new();
    for i in 0..req.field_names.len() {
        let idx = i as i32;
        if row_set.contains(&idx) {
            pivot_fields.push_str(r#"<pivotField axis="axisRow" showAll="0"/>"#);
        } else if col_set.contains(&idx) {
            pivot_fields.push_str(r#"<pivotField axis="axisCol" showAll="0"/>"#);
        } else if data_set.contains(&i) {
            pivot_fields.push_str(r#"<pivotField dataField="1" showAll="0"/>"#);
        } else {
            pivot_fields.push_str(r#"<pivotField showAll="0"/>"#);
        }
    }

    let row_fields = if req.row_indices.is_empty() {
        String::new()
    } else {
        let inner: String = req
            .row_indices
            .iter()
            .map(|x| format!(r#"<field x="{}"/>"#, x))
            .collect();
        format!(
            r#"<rowFields count="{}">{}</rowFields>"#,
            req.row_indices.len(),
            inner
        )
    };
    let col_fields = if req.col_indices.is_empty() {
        String::new()
    } else {
        let inner: String = req
            .col_indices
            .iter()
            .map(|x| format!(r#"<field x="{}"/>"#, x))
            .collect();
        format!(
            r#"<colFields count="{}">{}</colFields>"#,
            req.col_indices.len(),
            inner
        )
    };
    let data_fields = if req.data_fields.is_empty() {
        String::new()
    } else {
        let inner: String = req
            .data_fields
            .iter()
            .map(|(fld, name, subtotal)| {
                format!(
                    r#"<dataField name="{}" fld="{}" subtotal="{}" baseField="0" baseItem="0"/>"#,
                    escape(name),
                    fld,
                    escape(subtotal)
                )
            })
            .collect();
        format!(
            r#"<dataFields count="{}">{}</dataFields>"#,
            req.data_fields.len(),
            inner
        )
    };

    let first_data_col = req.row_indices.len().max(1);
    let table = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<pivotTableDefinition xmlns="{m}" name="{name}" cacheId="{cid}" applyNumberFormats="0" "#,
            r#"applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" "#,
            r#"applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" "#,
            r#"useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" "#,
            r#"outlineData="1" multipleFieldFilters="0">"#,
            r#"<location ref="{loc}" firstHeaderRow="1" firstDataRow="2" firstDataCol="{fdc}"/>"#,
            r#"<pivotFields count="{fcount}">{pfields}</pivotFields>"#,
            r#"{rowf}{colf}{dataf}"#,
            r#"<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" "#,
            r#"showRowStripes="0" showColStripes="0" showLastColumn="1"/>"#,
            r#"</pivotTableDefinition>"#
        ),
        m = NS_MAIN,
        name = escape(req.name),
        cid = req.cache_id,
        loc = escape(req.location_ref),
        fdc = first_data_col,
        fcount = req.field_names.len(),
        pfields = pivot_fields,
        rowf = row_fields,
        colf = col_fields,
        dataf = data_fields
    );

    let table_rels = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<Relationships xmlns="{p}">"#,
            r#"<Relationship Id="rId1" Type="{r}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition{n}.xml"/>"#,
            r#"</Relationships>"#
        ),
        p = NS_PKG_REL,
        r = NS_REL,
        n = req.cache_num
    );

    BuiltPivot {
        parts: vec![
            (cache_def_path, cache_def.into_bytes()),
            (cache_rec_path, cache_rec.into_bytes()),
            (cache_rels_path, cache_rels.into_bytes()),
            (table_path, table.into_bytes()),
            (table_rels_path, table_rels.into_bytes()),
        ],
        caches_child: format!(
            r#"<pivotCache cacheId="{}" r:id="{}"/>"#,
            req.cache_id, req.cache_rel_id
        ),
        workbook_rel: (
            req.cache_rel_id.to_string(),
            format!("pivotCache/pivotCacheDefinition{}.xml", req.cache_num),
        ),
        sheet_rel: (
            req.sheet_rel_id.to_string(),
            format!("{}/pivotTable", NS_REL),
            format!("../pivotTables/pivotTable{}.xml", req.table_num),
        ),
    }
}

/// Map a field index to its cache field name; `-2` is Excel's data-values
/// placeholder, other out-of-range indices fall back to a positional label.
fn field_name(index: i32, fields: &[String]) -> String {
    if index == -2 {
        "Values".to_string()
    } else if index >= 0 && (index as usize) < fields.len() {
        fields[index as usize].clone()
    } else {
        format!("Field{}", index)
    }
}

/// Join raw field placements with the cache field names into the public model.
fn combine(
    raw: RawPivotTable,
    source_sheet: Option<String>,
    source_ref: Option<String>,
    cache_fields: Vec<String>,
) -> PivotTableInfo {
    let row_fields = raw
        .row_x
        .iter()
        .map(|x| field_name(*x, &cache_fields))
        .collect();
    let col_fields = raw
        .col_x
        .iter()
        .map(|x| field_name(*x, &cache_fields))
        .collect();
    let page_fields = raw
        .page_fld
        .iter()
        .map(|x| field_name(*x, &cache_fields))
        .collect();
    let data_fields = raw
        .data
        .into_iter()
        .map(|(name, fld, subtotal)| {
            let source_field = field_name(fld, &cache_fields);
            PivotDataField {
                name: if name.is_empty() {
                    format!("{} of {}", subtotal, source_field)
                } else {
                    name
                },
                source_field,
                subtotal,
            }
        })
        .collect();

    PivotTableInfo {
        name: raw.name,
        cache_id: raw.cache_id,
        location: raw.location,
        source_sheet,
        source_ref,
        cache_fields,
        row_fields,
        col_fields,
        page_fields,
        data_fields,
    }
}

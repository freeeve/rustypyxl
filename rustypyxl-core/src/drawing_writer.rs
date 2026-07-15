//! Serialize a sheet's drawing part (`xl/drawings/drawingN.xml`): the anchors
//! for its charts (a `graphicFrame` referencing a chart part) and its images (a
//! `pic` referencing an embedded media part), plus the relationship rows the
//! drawing's `.rels` part binds those references to.
//!
//! Charts and images on one sheet share a single drawing part, so both flow
//! through here with a shared drawing-local relationship-id counter.

use crate::chart::{Chart, ChartAnchor};
use crate::image::{Image, ImageAnchorType};
use crate::writer::escape_xml;

const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const C_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const XDR_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

/// One relationship the drawing's `.rels` part must carry: a drawing-local id
/// (`rId1`, `rId2`, ...) bound to a chart part or an embedded media part.
pub struct DrawingRel {
    /// Drawing-local relationship id, e.g. "rId1".
    pub rel_id: String,
    /// Relationship target relative to the drawing part, e.g. "../charts/chart1.xml".
    pub target: String,
    /// True for a chart relationship, false for an image (media) relationship.
    pub is_chart: bool,
}

/// Parse an `A1`-style cell into 0-based (col, row) for drawing anchors.
fn anchor_cell(cell: &str) -> (u32, u32) {
    match crate::utils::parse_coordinate(cell) {
        Ok((row, col)) => (col.saturating_sub(1), row.saturating_sub(1)),
        Err(_) => (0, 0),
    }
}

/// `<xdr:from>` for the given cell and EMU offsets.
fn from_marker(cell: &str, col_off: u32, row_off: u32) -> String {
    let (col, row) = anchor_cell(cell);
    format!(
        r#"<xdr:from><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:from>"#,
        col, col_off, row, row_off
    )
}

/// `<xdr:to>` for the given cell and EMU offsets.
fn to_marker(cell: &str, col_off: u32, row_off: u32) -> String {
    let (col, row) = anchor_cell(cell);
    format!(
        r#"<xdr:to><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:to>"#,
        col, col_off, row, row_off
    )
}

/// A chart's `<xdr:graphicFrame>` referencing `rId{rel}`.
fn chart_frame(shape_id: u32, rel: u32) -> String {
    format!(
        concat!(
            r#"<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>"#,
            r#"<xdr:cNvPr id="{id}" name="Chart {n}"/><xdr:cNvGraphicFramePr/>"#,
            r#"</xdr:nvGraphicFramePr>"#,
            r#"<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>"#,
            r#"<a:graphic><a:graphicData uri="{c}">"#,
            r#"<c:chart xmlns:c="{c}" xmlns:r="{r}" r:id="rId{rel}"/>"#,
            r#"</a:graphicData></a:graphic></xdr:graphicFrame>"#
        ),
        id = shape_id,
        n = shape_id,
        c = C_NS,
        r = R_NS,
        rel = rel
    )
}

/// One anchor framing a chart at its position (default one-cell at A1).
fn chart_anchor(chart: &Chart, shape_id: u32, rel: u32) -> String {
    let default = ChartAnchor::at("A1");
    let a = chart.anchor.as_ref().unwrap_or(&default);
    let from = from_marker(&a.from_cell, a.from_col_offset, a.from_row_offset);
    let frame = chart_frame(shape_id, rel);

    match &a.to_cell {
        Some(to) => format!(
            r#"<xdr:twoCellAnchor>{from}{to}{frame}<xdr:clientData/></xdr:twoCellAnchor>"#,
            from = from,
            to = to_marker(to, a.to_col_offset, a.to_row_offset),
            frame = frame
        ),
        None => format!(
            r#"<xdr:oneCellAnchor>{from}<xdr:ext cx="{cx}" cy="{cy}"/>{frame}<xdr:clientData/></xdr:oneCellAnchor>"#,
            from = from,
            cx = chart.width,
            cy = chart.height,
            frame = frame
        ),
    }
}

/// An image's `<xdr:pic>` referencing embedded media `rId{rel}`.
fn image_pic(image: &Image, shape_id: u32, rel: u32) -> String {
    let name = image
        .name
        .clone()
        .unwrap_or_else(|| format!("Picture {}", shape_id));
    let descr = image
        .alt_text
        .as_deref()
        .or(image.description.as_deref())
        .unwrap_or("");
    format!(
        concat!(
            r#"<xdr:pic><xdr:nvPicPr>"#,
            r#"<xdr:cNvPr id="{id}" name="{name}" descr="{descr}"/>"#,
            r#"<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>"#,
            r#"</xdr:nvPicPr>"#,
            r#"<xdr:blipFill><a:blip xmlns:r="{r}" r:embed="rId{rel}"/>"#,
            r#"<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>"#,
            r#"<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>"#,
            r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>"#,
            r#"</xdr:pic>"#
        ),
        id = shape_id,
        name = escape_xml(&name),
        descr = escape_xml(descr),
        r = R_NS,
        rel = rel,
        cx = image.width,
        cy = image.height
    )
}

/// One anchor placing an image per its anchor type (one-cell, two-cell, absolute).
fn image_anchor(image: &Image, shape_id: u32, rel: u32) -> String {
    let a = &image.anchor;
    let pic = image_pic(image, shape_id, rel);

    match a.anchor_type {
        ImageAnchorType::TwoCell => {
            let to = a.to_cell.as_deref().unwrap_or(&a.from_cell);
            format!(
                r#"<xdr:twoCellAnchor editAs="oneCell">{from}{to}{pic}<xdr:clientData/></xdr:twoCellAnchor>"#,
                from = from_marker(&a.from_cell, a.from_col_offset, a.from_row_offset),
                to = to_marker(to, a.to_col_offset, a.to_row_offset),
                pic = pic
            )
        }
        ImageAnchorType::Absolute => format!(
            r#"<xdr:absoluteAnchor><xdr:pos x="{x}" y="{y}"/><xdr:ext cx="{cx}" cy="{cy}"/>{pic}<xdr:clientData/></xdr:absoluteAnchor>"#,
            x = a.from_col_offset,
            y = a.from_row_offset,
            cx = image.width,
            cy = image.height,
            pic = pic
        ),
        ImageAnchorType::OneCell => format!(
            r#"<xdr:oneCellAnchor>{from}<xdr:ext cx="{cx}" cy="{cy}"/>{pic}<xdr:clientData/></xdr:oneCellAnchor>"#,
            from = from_marker(&a.from_cell, a.from_col_offset, a.from_row_offset),
            cx = image.width,
            cy = image.height,
            pic = pic
        ),
    }
}

/// Build a sheet's drawing XML and the relationship rows its `.rels` part needs.
///
/// `charts` and `images` pair each item with its workbook-unique part id (chart
/// part number, media number) so anchors can bind to `../charts/chartN.xml` and
/// `../media/imageN.ext`. Drawing-local relationship ids are assigned in order:
/// charts first, then images.
pub fn drawing_for_sheet(
    charts: &[(&Chart, u32)],
    images: &[(&Image, u32)],
) -> (String, Vec<DrawingRel>) {
    let mut body = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="{xdr}" xmlns:a="{a}">"#,
        xdr = XDR_NS,
        a = A_NS
    );
    let mut rels = Vec::with_capacity(charts.len() + images.len());
    // Drawing shape ids and relationship ids are both 1-based and unique within
    // the drawing; run one counter across charts then images.
    let mut next = 1u32;

    for (chart, chart_id) in charts {
        body.push_str(&chart_anchor(chart, next, next));
        rels.push(DrawingRel {
            rel_id: format!("rId{}", next),
            target: format!("../charts/chart{}.xml", chart_id),
            is_chart: true,
        });
        next += 1;
    }

    for (image, media_id) in images {
        body.push_str(&image_anchor(image, next, next));
        rels.push(DrawingRel {
            rel_id: format!("rId{}", next),
            target: format!("../media/image{}.{}", media_id, image.format.extension()),
            is_chart: false,
        });
        next += 1;
    }

    body.push_str("</xdr:wsDr>");
    (body, rels)
}

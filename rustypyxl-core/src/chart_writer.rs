//! Serialize charts to the OOXML parts a saved workbook needs: the chart
//! definition (xl/charts/chartN.xml) and the drawing that anchors it on a sheet
//! (xl/drawings/drawingM.xml). Write-path only; reading charts back on load is a
//! separate concern.

use crate::chart::{Chart, ChartSeries, ChartType};
use crate::writer::escape_xml;

const CAT_AX_ID: &str = "111111111";
const VAL_AX_ID: &str = "222222222";
// For scatter both axes are value axes.
const X_VAL_AX_ID: &str = "111111112";

const C_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XDR_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

/// `<c:tx>` for a series name: a reference when it looks like one (`Sheet!$A$1`),
/// otherwise a literal string value.
fn series_tx(name: &str) -> String {
    if name.contains('!') {
        format!(
            r#"<c:tx><c:strRef><c:f>{}</c:f></c:strRef></c:tx>"#,
            escape_xml(name)
        )
    } else {
        format!(r#"<c:tx><c:v>{}</c:v></c:tx>"#, escape_xml(name))
    }
}

/// A category axis reference `<c:cat>`.
fn cat_ref(categories: &str) -> String {
    format!(
        r#"<c:cat><c:strRef><c:f>{}</c:f></c:strRef></c:cat>"#,
        escape_xml(categories)
    )
}

/// A numeric value reference wrapped in the given element (`val`, `xVal`, `yVal`).
fn num_ref(tag: &str, reference: &str) -> String {
    format!(
        r#"<c:{t}><c:numRef><c:f>{r}</c:f></c:numRef></c:{t}>"#,
        t = tag,
        r = escape_xml(reference)
    )
}

fn solid_fill(color: &str) -> String {
    let hex = color.strip_prefix('#').unwrap_or(color);
    format!(
        r#"<c:spPr><a:solidFill><a:srgbClr val="{}"/></a:solidFill></c:spPr>"#,
        escape_xml(hex)
    )
}

/// A `<c:ser>` for a category-based chart (bar/line/area/pie).
fn category_series(idx: usize, s: &ChartSeries) -> String {
    let mut out = format!(r#"<c:ser><c:idx val="{i}"/><c:order val="{i}"/>"#, i = idx);
    if let Some(name) = &s.name {
        out.push_str(&series_tx(name));
    }
    if let Some(fill) = &s.fill_color {
        out.push_str(&solid_fill(fill));
    }
    if let Some(cats) = &s.categories {
        out.push_str(&cat_ref(cats));
    }
    out.push_str(&num_ref("val", &s.values));
    out.push_str("</c:ser>");
    out
}

/// A `<c:ser>` for a scatter chart (xVal/yVal).
fn scatter_series(idx: usize, s: &ChartSeries) -> String {
    let mut out = format!(r#"<c:ser><c:idx val="{i}"/><c:order val="{i}"/>"#, i = idx);
    if let Some(name) = &s.name {
        out.push_str(&series_tx(name));
    }
    // x from categories, y from values
    if let Some(cats) = &s.categories {
        out.push_str(&num_ref("xVal", cats));
    }
    out.push_str(&num_ref("yVal", &s.values));
    out.push_str("</c:ser>");
    out
}

/// Grouping token for the chart's grouping element. Bar/column charts spell the
/// default "clustered"; line/area charts spell it "standard".
fn grouping(chart: &Chart, is_bar: bool) -> &'static str {
    use crate::chart::BarGrouping::*;
    match chart.bar_grouping {
        Clustered => {
            if is_bar {
                "clustered"
            } else {
                "standard"
            }
        }
        Stacked => "stacked",
        PercentStacked => "percentStacked",
    }
}

fn title_xml(chart: &Chart) -> String {
    match &chart.title {
        Some(t) => {
            let text = t.text.as_deref().unwrap_or("");
            format!(
                r#"<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>{}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>"#,
                escape_xml(text)
            )
        }
        None => r#"<c:autoTitleDeleted val="1"/>"#.to_string(),
    }
}

fn legend_xml(chart: &Chart) -> String {
    match &chart.legend {
        Some(l) if l.visible => format!(
            r#"<c:legend><c:legendPos val="{}"/><c:overlay val="0"/></c:legend>"#,
            escape_xml(&l.position)
        ),
        _ => String::new(),
    }
}

/// A `<c:catAx>` or `<c:valAx>`.
fn axis(kind: &str, ax_id: &str, cross_id: &str, pos: &str) -> String {
    format!(
        r#"<c:{kind}><c:axId val="{id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="{pos}"/><c:crossAx val="{cross}"/></c:{kind}>"#,
        kind = kind,
        id = ax_id,
        pos = pos,
        cross = cross_id
    )
}

/// The `xl/charts/chartN.xml` content for a chart.
pub fn chart_xml(chart: &Chart) -> String {
    let plot = match chart.chart_type {
        ChartType::Pie | ChartType::Doughnut => {
            let mut body = String::from(r#"<c:pieChart><c:varyColors val="1"/>"#);
            for (i, s) in chart.series.iter().enumerate() {
                body.push_str(&category_series(i, s));
            }
            if matches!(chart.chart_type, ChartType::Doughnut) {
                // doughnutChart is the same shape plus a hole; reuse pieChart body but
                // emit as doughnutChart with holeSize.
                body = body.replacen("<c:pieChart>", "<c:doughnutChart>", 1);
                body.push_str(r#"<c:holeSize val="50"/></c:doughnutChart>"#);
            } else {
                body.push_str("</c:pieChart>");
            }
            body
        }
        ChartType::Scatter | ChartType::Bubble => {
            let mut body = String::from(r#"<c:scatterChart><c:scatterStyle val="lineMarker"/>"#);
            for (i, s) in chart.series.iter().enumerate() {
                body.push_str(&scatter_series(i, s));
            }
            body.push_str(&format!(r#"<c:axId val="{}"/>"#, X_VAL_AX_ID));
            body.push_str(&format!(r#"<c:axId val="{}"/>"#, VAL_AX_ID));
            body.push_str("</c:scatterChart>");
            // scatter uses two value axes
            body.push_str(&axis("valAx", X_VAL_AX_ID, VAL_AX_ID, "b"));
            body.push_str(&axis("valAx", VAL_AX_ID, X_VAL_AX_ID, "l"));
            body
        }
        _ => {
            // bar / column / line / area share a category + value axis pair.
            let (elem, extra): (&str, String) = match chart.chart_type {
                ChartType::Line | ChartType::Radar | ChartType::Stock => (
                    "lineChart",
                    format!(
                        r#"<c:grouping val="{}"/><c:varyColors val="0"/>"#,
                        grouping(chart, false)
                    ),
                ),
                ChartType::Area => (
                    "areaChart",
                    format!(
                        r#"<c:grouping val="{}"/><c:varyColors val="0"/>"#,
                        grouping(chart, false)
                    ),
                ),
                _ => {
                    // Bar / Column (and any unsupported type) -> barChart.
                    let dir = match chart.bar_direction {
                        crate::chart::BarDirection::Bar => "bar",
                        crate::chart::BarDirection::Col => "col",
                    };
                    (
                        "barChart",
                        format!(
                            r#"<c:barDir val="{}"/><c:grouping val="{}"/><c:varyColors val="0"/>"#,
                            dir,
                            grouping(chart, true)
                        ),
                    )
                }
            };
            let mut body = format!("<c:{}>{}", elem, extra);
            for (i, s) in chart.series.iter().enumerate() {
                body.push_str(&category_series(i, s));
            }
            body.push_str(&format!(r#"<c:axId val="{}"/>"#, CAT_AX_ID));
            body.push_str(&format!(r#"<c:axId val="{}"/>"#, VAL_AX_ID));
            body.push_str(&format!("</c:{}>", elem));
            body.push_str(&axis("catAx", CAT_AX_ID, VAL_AX_ID, "b"));
            body.push_str(&axis("valAx", VAL_AX_ID, CAT_AX_ID, "l"));
            body
        }
    };

    format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<c:chartSpace xmlns:c="{c}" xmlns:a="{a}" xmlns:r="{r}">"#,
            r#"<c:chart>{title}<c:plotArea><c:layout/>{plot}</c:plotArea>{legend}<c:plotVisOnly val="1"/></c:chart>"#,
            r#"</c:chartSpace>"#
        ),
        c = C_NS,
        a = A_NS,
        r = R_NS,
        title = title_xml(chart),
        plot = plot,
        legend = legend_xml(chart)
    )
}

/// Parse an `A1`-style cell into 0-based (col, row) for drawing anchors.
fn anchor_cell(cell: &str) -> (u32, u32) {
    match crate::utils::parse_coordinate(cell) {
        Ok((row, col)) => (col.saturating_sub(1), row.saturating_sub(1)),
        Err(_) => (0, 0),
    }
}

/// One `<xdr:*Anchor>` framing a chart that references `rId{rel_idx}`.
fn chart_anchor(chart: &Chart, rel_idx: usize) -> String {
    let default = crate::chart::ChartAnchor::at("A1");
    let a = chart.anchor.as_ref().unwrap_or(&default);
    let (from_col, from_row) = anchor_cell(&a.from_cell);

    let frame = format!(
        concat!(
            r#"<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>"#,
            r#"<xdr:cNvPr id="{id}" name="Chart {n}"/><xdr:cNvGraphicFramePr/>"#,
            r#"</xdr:nvGraphicFramePr>"#,
            r#"<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>"#,
            r#"<a:graphic><a:graphicData uri="{c}">"#,
            r#"<c:chart xmlns:c="{c}" xmlns:r="{r}" r:id="rId{rel}"/>"#,
            r#"</a:graphicData></a:graphic></xdr:graphicFrame>"#
        ),
        id = rel_idx + 1,
        n = rel_idx,
        c = C_NS,
        r = R_NS,
        rel = rel_idx
    );

    let from = format!(
        r#"<xdr:from><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:from>"#,
        from_col, a.from_col_offset, from_row, a.from_row_offset
    );

    match &a.to_cell {
        Some(to) => {
            let (to_col, to_row) = anchor_cell(to);
            format!(
                r#"<xdr:twoCellAnchor>{from}<xdr:to><xdr:col>{tc}</xdr:col><xdr:colOff>{tco}</xdr:colOff><xdr:row>{tr}</xdr:row><xdr:rowOff>{tro}</xdr:rowOff></xdr:to>{frame}<xdr:clientData/></xdr:twoCellAnchor>"#,
                from = from,
                tc = to_col,
                tco = a.to_col_offset,
                tr = to_row,
                tro = a.to_row_offset,
                frame = frame
            )
        }
        None => format!(
            r#"<xdr:oneCellAnchor>{from}<xdr:ext cx="{cx}" cy="{cy}"/>{frame}<xdr:clientData/></xdr:oneCellAnchor>"#,
            from = from,
            cx = chart.width,
            cy = chart.height,
            frame = frame
        ),
    }
}

/// The `xl/drawings/drawingM.xml` content anchoring a sheet's charts. Each chart
/// is referenced by `rId{i+1}` (matching the drawing's .rels part).
pub fn drawing_xml(charts: &[Chart]) -> String {
    let mut body = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="{xdr}" xmlns:a="{a}">"#,
        xdr = XDR_NS,
        a = A_NS
    );
    for (i, chart) in charts.iter().enumerate() {
        body.push_str(&chart_anchor(chart, i + 1));
    }
    body.push_str("</xdr:wsDr>");
    body
}

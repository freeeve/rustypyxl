//! A chart added to a worksheet is serialized on save: the chart part, the
//! drawing that anchors it, the relationships tying them together, and the
//! content-type overrides all appear and are well-formed XML.

use rustypyxl::chart::{Chart, ChartLegend, ChartSeries, ChartType};
use rustypyxl::{CellValue, Workbook};
use std::io::{Cursor, Read};
use zip::ZipArchive;

fn read_part(bytes: &[u8], name: &str) -> Option<String> {
    let mut zip = ZipArchive::new(Cursor::new(bytes.to_vec())).unwrap();
    let mut file = zip.by_name(name).ok()?;
    let mut s = String::new();
    file.read_to_string(&mut s).unwrap();
    Some(s)
}

fn workbook_with_bar_chart() -> Workbook {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("Data".to_string())).unwrap();
    for (i, v) in [("A", 10.0), ("B", 20.0), ("C", 30.0)].iter().enumerate() {
        let row = (i + 1) as u32;
        wb.set_cell_value_in_sheet("Data", row, 1, CellValue::from(v.0))
            .unwrap();
        wb.set_cell_value_in_sheet("Data", row, 2, CellValue::Number(v.1))
            .unwrap();
    }

    let mut chart = Chart::column().with_title("Sales");
    chart.add_series(
        ChartSeries::new("Data!$B$1:$B$3")
            .with_name("Revenue")
            .with_categories("Data!$A$1:$A$3"),
    );
    chart = chart.with_legend(ChartLegend::new().with_position("b"));
    wb.get_sheet_by_name_mut("Data").unwrap().add_chart(chart);
    wb
}

#[test]
fn bar_chart_emits_parts_and_relationships() {
    let bytes = workbook_with_bar_chart().save_to_bytes().unwrap();

    let chart = read_part(&bytes, "xl/charts/chart1.xml").expect("chart part present");
    assert!(chart.contains("<c:barChart>"), "column -> barChart");
    assert!(chart.contains(r#"<c:barDir val="col"/>"#));
    assert!(chart.contains("Data!$B$1:$B$3"), "value ref present");
    assert!(chart.contains("Data!$A$1:$A$3"), "category ref present");
    assert!(chart.contains("Revenue"), "series name present");
    assert!(chart.contains(r#"<c:legendPos val="b"/>"#));
    assert!(chart.contains("Sales"), "title present");

    let drawing = read_part(&bytes, "xl/drawings/drawing1.xml").expect("drawing part present");
    assert!(drawing.contains("<xdr:oneCellAnchor>"), "default anchor");
    assert!(
        drawing.contains(r#"r:id="rId1""#),
        "frame references chart rel"
    );

    let drawing_rels =
        read_part(&bytes, "xl/drawings/_rels/drawing1.xml.rels").expect("drawing rels present");
    assert!(drawing_rels.contains("../charts/chart1.xml"));

    let sheet_rels =
        read_part(&bytes, "xl/worksheets/_rels/sheet1.xml.rels").expect("sheet rels present");
    assert!(sheet_rels.contains("rIdDrawing"));
    assert!(sheet_rels.contains("../drawings/drawing1.xml"));

    let sheet = read_part(&bytes, "xl/worksheets/sheet1.xml").unwrap();
    assert!(sheet.contains(r#"<drawing r:id="rIdDrawing"/>"#));

    let content_types = read_part(&bytes, "[Content_Types].xml").unwrap();
    assert!(content_types.contains("/xl/charts/chart1.xml"));
    assert!(content_types.contains("/xl/drawings/drawing1.xml"));
}

#[test]
fn scatter_chart_uses_two_value_axes() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("S".to_string())).unwrap();
    let mut chart = Chart::scatter();
    chart.add_series(ChartSeries::new("S!$B$1:$B$3").with_categories("S!$A$1:$A$3"));
    wb.get_sheet_by_name_mut("S").unwrap().add_chart(chart);

    let bytes = wb.save_to_bytes().unwrap();
    let chart = read_part(&bytes, "xl/charts/chart1.xml").unwrap();
    assert!(chart.contains("<c:scatterChart>"));
    assert!(chart.contains("<c:xVal>"));
    assert!(chart.contains("<c:yVal>"));
    assert_eq!(chart.matches("<c:valAx>").count(), 2, "two value axes");
}

#[test]
fn pie_chart_has_no_axes() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("S".to_string())).unwrap();
    let mut chart = Chart::pie();
    chart.add_series(ChartSeries::new("S!$B$1:$B$3"));
    wb.get_sheet_by_name_mut("S").unwrap().add_chart(chart);

    let bytes = wb.save_to_bytes().unwrap();
    let chart = read_part(&bytes, "xl/charts/chart1.xml").unwrap();
    assert!(chart.contains("<c:pieChart>"));
    assert!(!chart.contains("catAx"));
    assert!(!chart.contains("valAx"));
}

#[test]
fn charts_across_sheets_get_unique_ids() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("One".to_string())).unwrap();
    wb.create_sheet(Some("Two".to_string())).unwrap();

    let mut c1 = Chart::column();
    c1.add_series(ChartSeries::new("One!$A$1:$A$3"));
    wb.get_sheet_by_name_mut("One").unwrap().add_chart(c1);

    let mut c2 = Chart::line();
    c2.add_series(ChartSeries::new("Two!$A$1:$A$3"));
    wb.get_sheet_by_name_mut("Two").unwrap().add_chart(c2);

    let bytes = wb.save_to_bytes().unwrap();
    // Sheet One is sheet1 -> chart1/drawing1; Sheet Two is sheet2 -> chart2/drawing2.
    assert!(read_part(&bytes, "xl/charts/chart1.xml")
        .unwrap()
        .contains("<c:barChart>"));
    assert!(read_part(&bytes, "xl/charts/chart2.xml")
        .unwrap()
        .contains("<c:lineChart>"));
    assert!(read_part(&bytes, "xl/drawings/drawing2.xml").is_some());
    let dr2 = read_part(&bytes, "xl/drawings/_rels/drawing2.xml.rels").unwrap();
    assert!(dr2.contains("../charts/chart2.xml"));
}

#[test]
fn chart_survives_load_save_round_trip() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("Data".to_string())).unwrap();
    let mut chart = Chart::column().with_title("Sales");
    chart.add_series(
        ChartSeries::new("Data!$B$1:$B$3")
            .with_name("Revenue")
            .with_categories("Data!$A$1:$A$3"),
    );
    chart = chart.with_legend(ChartLegend::new().with_position("b"));
    chart = chart.with_anchor(rustypyxl::chart::ChartAnchor::at("D2"));
    wb.get_sheet_by_name_mut("Data").unwrap().add_chart(chart);

    let bytes = wb.save_to_bytes().unwrap();
    let reloaded = Workbook::load_from_bytes(&bytes).unwrap();
    let ws = reloaded.get_sheet_by_name("Data").unwrap();

    assert_eq!(ws.charts.len(), 1, "chart read back on load");
    let got = &ws.charts[0];
    assert_eq!(got.chart_type, ChartType::Column);
    assert_eq!(got.series.len(), 1);
    assert_eq!(got.series[0].values, "Data!$B$1:$B$3");
    assert_eq!(got.series[0].categories.as_deref(), Some("Data!$A$1:$A$3"));
    assert_eq!(got.series[0].name.as_deref(), Some("Revenue"));
    assert_eq!(
        got.title.as_ref().and_then(|t| t.text.as_deref()),
        Some("Sales")
    );
    assert_eq!(got.legend.as_ref().map(|l| l.position.as_str()), Some("b"));
    assert_eq!(
        got.anchor.as_ref().map(|a| a.from_cell.as_str()),
        Some("D2")
    );

    // Re-saving the reloaded workbook still carries the chart.
    let bytes2 = reloaded.save_to_bytes().unwrap();
    assert!(read_part(&bytes2, "xl/charts/chart1.xml")
        .unwrap()
        .contains("<c:barChart>"));
}

#[test]
fn round_trips_line_pie_and_scatter_types() {
    for (make, label) in [
        (Chart::line as fn() -> Chart, ChartType::Line),
        (Chart::pie as fn() -> Chart, ChartType::Pie),
        (Chart::scatter as fn() -> Chart, ChartType::Scatter),
    ] {
        let mut wb = Workbook::new();
        wb.create_sheet(Some("S".to_string())).unwrap();
        let mut chart = make();
        chart.add_series(ChartSeries::new("S!$B$1:$B$3").with_categories("S!$A$1:$A$3"));
        wb.get_sheet_by_name_mut("S").unwrap().add_chart(chart);

        let bytes = wb.save_to_bytes().unwrap();
        let reloaded = Workbook::load_from_bytes(&bytes).unwrap();
        let ws = reloaded.get_sheet_by_name("S").unwrap();
        assert_eq!(ws.charts.len(), 1, "{label:?} chart read back");
        assert_eq!(ws.charts[0].chart_type, label);
        assert_eq!(ws.charts[0].series[0].values, "S!$B$1:$B$3");
        assert_eq!(
            ws.charts[0].series[0].categories.as_deref(),
            Some("S!$A$1:$A$3")
        );
    }
}

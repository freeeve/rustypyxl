//! An image added to a worksheet is embedded on save (media part, drawing pic
//! anchor, relationships, content type), and images present in a loaded file
//! survive being saved again.

use rustypyxl::chart::{Chart, ChartSeries};
use rustypyxl::image::{Image, ImageAnchor, ImageAnchorType};
use rustypyxl::Workbook;
use std::io::{Cursor, Read};
use zip::ZipArchive;

/// A minimal valid 1x1 PNG.
const PNG_1X1: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
    0x42, 0x60, 0x82,
];

fn read_part(bytes: &[u8], name: &str) -> Option<String> {
    let mut zip = ZipArchive::new(Cursor::new(bytes.to_vec())).unwrap();
    let mut file = zip.by_name(name).ok()?;
    let mut s = String::new();
    file.read_to_string(&mut s).unwrap();
    Some(s)
}

fn part_bytes(bytes: &[u8], name: &str) -> Option<Vec<u8>> {
    let mut zip = ZipArchive::new(Cursor::new(bytes.to_vec())).unwrap();
    let mut file = zip.by_name(name).ok()?;
    let mut v = Vec::new();
    file.read_to_end(&mut v).unwrap();
    Some(v)
}

#[test]
fn image_emits_media_drawing_and_content_type() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("S".to_string())).unwrap();
    let img = Image::from_bytes(PNG_1X1.to_vec(), ImageAnchor::one_cell("B2"))
        .unwrap()
        .with_name("Logo")
        .with_alt_text("company logo");
    wb.get_sheet_by_name_mut("S").unwrap().add_image(img);

    let bytes = wb.save_to_bytes().unwrap();

    // The media blob is stored verbatim.
    assert_eq!(
        part_bytes(&bytes, "xl/media/image1.png").as_deref(),
        Some(PNG_1X1)
    );

    let drawing = read_part(&bytes, "xl/drawings/drawing1.xml").expect("drawing present");
    assert!(drawing.contains("<xdr:pic>"));
    assert!(drawing.contains(r#"r:embed="rId1""#));
    assert!(drawing.contains("Logo"));
    assert!(drawing.contains("company logo"));
    // B2 -> 0-based col 1, row 1
    assert!(drawing.contains("<xdr:col>1</xdr:col>"));

    let drels = read_part(&bytes, "xl/drawings/_rels/drawing1.xml.rels").unwrap();
    assert!(drels.contains("../media/image1.png"));
    assert!(drels.contains("/relationships/image"));

    let content_types = read_part(&bytes, "[Content_Types].xml").unwrap();
    assert!(content_types.contains(r#"Extension="png""#));
    assert!(content_types.contains("image/png"));
}

#[test]
fn image_survives_load_save_round_trip() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("S".to_string())).unwrap();
    let img = Image::from_bytes(PNG_1X1.to_vec(), ImageAnchor::two_cell("A1", "C5"))
        .unwrap()
        .with_size_px(120, 80)
        .with_name("Pic");
    wb.get_sheet_by_name_mut("S").unwrap().add_image(img);
    let bytes = wb.save_to_bytes().unwrap();

    // Load it back: the image must be present on the sheet.
    let reloaded = Workbook::load_from_bytes(&bytes).unwrap();
    let ws = reloaded.get_sheet_by_name("S").unwrap();
    assert_eq!(ws.images.len(), 1, "image preserved on load");
    let got = &ws.images[0];
    assert_eq!(got.data, PNG_1X1, "media bytes intact");
    assert_eq!(got.format, rustypyxl::image::ImageFormat::Png);
    assert_eq!(got.anchor.anchor_type, ImageAnchorType::TwoCell);
    assert_eq!(got.anchor.from_cell, "A1");
    assert_eq!(got.anchor.to_cell.as_deref(), Some("C5"));
    assert_eq!(got.name.as_deref(), Some("Pic"));

    // And saving the reloaded workbook still carries the media.
    let bytes2 = reloaded.save_to_bytes().unwrap();
    assert_eq!(
        part_bytes(&bytes2, "xl/media/image1.png").as_deref(),
        Some(PNG_1X1)
    );
}

#[test]
fn chart_and_image_share_one_drawing() {
    let mut wb = Workbook::new();
    wb.create_sheet(Some("S".to_string())).unwrap();
    let mut chart = Chart::column();
    chart.add_series(ChartSeries::new("S!$A$1:$A$3"));
    wb.get_sheet_by_name_mut("S").unwrap().add_chart(chart);
    let img = Image::from_bytes(PNG_1X1.to_vec(), ImageAnchor::one_cell("E1")).unwrap();
    wb.get_sheet_by_name_mut("S").unwrap().add_image(img);

    let bytes = wb.save_to_bytes().unwrap();

    // One drawing part carries both the chart frame and the picture.
    let drawing = read_part(&bytes, "xl/drawings/drawing1.xml").expect("drawing present");
    assert!(drawing.contains("<xdr:graphicFrame"), "chart frame present");
    assert!(drawing.contains("<xdr:pic>"), "picture present");

    // Its rels bind both a chart part and a media part, with distinct rIds.
    let drels = read_part(&bytes, "xl/drawings/_rels/drawing1.xml.rels").unwrap();
    assert!(drels.contains("../charts/chart1.xml"));
    assert!(drels.contains("../media/image1.png"));
    assert!(drels.contains(r#"Id="rId1""#));
    assert!(drels.contains(r#"Id="rId2""#));

    // The image round-trips even alongside a chart (the chart does not).
    let reloaded = Workbook::load_from_bytes(&bytes).unwrap();
    assert_eq!(reloaded.get_sheet_by_name("S").unwrap().images.len(), 1);
}

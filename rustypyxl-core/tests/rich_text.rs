//! Rich-text runs within a cell survive a save/load round-trip, and the cell's
//! plain value stays the concatenated text.

use rustypyxl::style::Color;
use rustypyxl::{CellValue, RichText, RunFont, TextRun, Workbook};

fn bold() -> RunFont {
    RunFont {
        bold: true,
        ..Default::default()
    }
}

fn red() -> RunFont {
    RunFont {
        color: Some(Color::rgb("FFFF0000")),
        size: Some(14.0),
        ..Default::default()
    }
}

#[test]
fn rich_text_runs_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.create_sheet(Some("S".to_string())).unwrap();
    // "Total: 42" where "Total:" is bold and " 42" is red 14pt.
    let rich = RichText::new(vec![
        TextRun::formatted("Total:", bold()),
        TextRun::formatted(" 42", red()),
    ]);
    ws.set_cell_rich_text(1, 1, rich.clone());

    // plain value is the concatenation
    assert_eq!(
        ws.get_cell_value(1, 1),
        Some(&CellValue::String("Total: 42".into()))
    );

    let reloaded = Workbook::load_from_bytes(&wb.save_to_bytes().unwrap()).unwrap();
    let ws = reloaded.get_sheet_by_name("S").unwrap();
    let cell = ws.get_cell(1, 1).unwrap();

    assert_eq!(
        cell.value,
        CellValue::String("Total: 42".into()),
        "plain value preserved"
    );
    let got = cell.rich_text.as_ref().expect("rich runs preserved");
    assert_eq!(got.runs.len(), 2);
    assert_eq!(got.runs[0].text, "Total:");
    assert!(got.runs[0].font.as_ref().unwrap().bold);
    assert_eq!(got.runs[1].text, " 42");
    let f = got.runs[1].font.as_ref().unwrap();
    assert_eq!(f.size, Some(14.0));
    assert_eq!(f.color.as_ref().unwrap().rgb.as_deref(), Some("#FFFF0000"));
}

#[test]
fn plain_string_has_no_rich_runs() {
    let mut wb = Workbook::new();
    let ws = wb.create_sheet(Some("S".to_string())).unwrap();
    ws.set_cell_value(1, 1, CellValue::String("plain".into()));

    let reloaded = Workbook::load_from_bytes(&wb.save_to_bytes().unwrap()).unwrap();
    let cell = reloaded
        .get_sheet_by_name("S")
        .unwrap()
        .get_cell(1, 1)
        .unwrap();
    assert_eq!(cell.value, CellValue::String("plain".into()));
    assert!(
        cell.rich_text.is_none(),
        "a plain string must not gain rich runs"
    );
}

#[test]
fn run_with_no_font_inherits() {
    // A run with no rPr (font = None) must survive as an inherit-the-cell run.
    let mut wb = Workbook::new();
    let ws = wb.create_sheet(Some("S".to_string())).unwrap();
    let rich = RichText::new(vec![TextRun::plain("a"), TextRun::formatted("b", bold())]);
    ws.set_cell_rich_text(1, 1, rich);

    let reloaded = Workbook::load_from_bytes(&wb.save_to_bytes().unwrap()).unwrap();
    let cell = reloaded
        .get_sheet_by_name("S")
        .unwrap()
        .get_cell(1, 1)
        .unwrap();
    let runs = &cell.rich_text.as_ref().unwrap().runs;
    assert_eq!(runs.len(), 2);
    assert_eq!(runs[0].text, "a");
    assert!(runs[0].font.is_none(), "unformatted run stays font-less");
    assert!(runs[1].font.as_ref().unwrap().bold);
}

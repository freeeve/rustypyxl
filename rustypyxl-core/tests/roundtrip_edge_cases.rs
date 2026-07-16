//! Curated edge-case round-trip: values that a random generator won't reliably
//! hit -- empty and whitespace strings, embedded tab/newline, XML-special
//! characters, unicode, a very long string, and float extremes -- must all
//! survive save -> load. Complements the randomized property test.

use rustypyxl::{CellValue, Workbook};

#[test]
fn curated_edge_cases_roundtrip() {
    let values: Vec<CellValue> = vec![
        CellValue::from("hello"),
        CellValue::from(""),
        CellValue::from(" "),
        CellValue::from("  trailing  "),
        CellValue::from("line\nbreak"),
        CellValue::from("tab\tchar"),
        CellValue::from("<>&\"'"),
        CellValue::from("emoji 😀 unicode ✓"),
        CellValue::from("x".repeat(40000)),
        CellValue::Number(0.0),
        CellValue::Number(-3.5),
        CellValue::Number(1e20),
        CellValue::Number(1e-20),
        CellValue::Number(f64::MAX),
        CellValue::Number(1.0_f64 / 3.0),
        CellValue::Number(2.0_f64.sqrt()),
        CellValue::Number(0.1 + 0.2),
        CellValue::Number(9007199254740993.0),
        CellValue::Boolean(true),
        CellValue::Boolean(false),
        CellValue::Formula("SUM(A1:A2)".to_string()),
        CellValue::Date("2023-01-15".to_string()),
    ];

    let mut wb = Workbook::new();
    wb.create_sheet(Some("S".to_string())).unwrap();
    for (i, v) in values.iter().enumerate() {
        wb.set_cell_value_in_sheet("S", (i + 1) as u32, 1, v.clone())
            .unwrap();
    }
    let bytes = wb.save_to_bytes().unwrap();
    let back = Workbook::load_from_bytes(&bytes).unwrap();
    let ws = back.get_sheet_by_name("S").unwrap();

    let mut mismatches = 0;
    for (i, v) in values.iter().enumerate() {
        let got = ws.get_cell_value((i + 1) as u32, 1);
        let ok = match (v, got) {
            (CellValue::Empty, None) => true,
            (a, Some(b)) => a == b,
            _ => false,
        };
        if !ok {
            mismatches += 1;
            let show = |c: &CellValue| match c {
                CellValue::String(s) if s.len() > 30 => format!("String(len {})", s.len()),
                other => format!("{:?}", other),
            };
            println!(
                "MISMATCH [{i}] sent {} -> got {}",
                show(v),
                got.map(show).unwrap_or_else(|| "None".into())
            );
        }
    }
    assert_eq!(
        mismatches, 0,
        "some edge-case values did not survive a round-trip"
    );
}

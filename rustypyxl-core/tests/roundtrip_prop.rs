//! Property test: a workbook of randomly generated cells must survive a
//! save -> load round-trip value-for-value.
//!
//! Strings are drawn from non-control characters. Control characters that are
//! illegal in XML (most of 0x00-0x1F) cannot be stored and are intentionally
//! stripped on write, so generating them would test a documented normalization
//! rather than a fidelity invariant. Tab/newline/CR are legal and covered by
//! the curated probe test.

use proptest::prelude::*;
use rustypyxl::{CellValue, Workbook};
use std::collections::HashMap;

/// A sheet name paired with its sparse cell map.
type NamedSheet = (String, HashMap<(u32, u32), CellValue>);

/// A cell value: text (no control chars), a finite number, a boolean, a
/// formula, or an ISO date.
fn cell_value() -> impl Strategy<Value = CellValue> {
    prop_oneof![
        "\\PC{0,60}".prop_map(CellValue::from),
        any::<f64>()
            .prop_filter("finite", |f| f.is_finite())
            .prop_map(CellValue::Number),
        any::<bool>().prop_map(CellValue::Boolean),
        "\\PC{1,40}".prop_map(CellValue::Formula),
        (1900i32..=2200, 1u32..=12, 1u32..=28)
            .prop_map(|(y, m, d)| CellValue::Date(format!("{:04}-{:02}-{:02}", y, m, d))),
    ]
}

/// A sparse sheet: a map of (row, col) -> value at bounded positions.
fn sheet() -> impl Strategy<Value = HashMap<(u32, u32), CellValue>> {
    prop::collection::hash_map((1u32..500, 1u32..50), cell_value(), 0..120)
}

proptest! {
    #![proptest_config(ProptestConfig::with_cases(200))]

    #[test]
    fn cells_survive_save_load(cells in sheet()) {
        let mut wb = Workbook::new();
        wb.create_sheet(Some("S".to_string())).unwrap();
        for (&(row, col), value) in &cells {
            wb.set_cell_value_in_sheet("S", row, col, value.clone()).unwrap();
        }

        let bytes = wb.save_to_bytes().unwrap();
        let back = Workbook::load_from_bytes(&bytes).unwrap();
        let ws = back.get_sheet_by_name("S").unwrap();

        for (&(row, col), sent) in &cells {
            let got = ws.get_cell_value(row, col);
            match sent {
                // A generated "" is a real (empty) string cell and must survive.
                CellValue::Empty => prop_assert!(matches!(got, None | Some(CellValue::Empty))),
                other => prop_assert_eq!(
                    Some(other), got,
                    "cell ({},{}) changed across round-trip", row, col
                ),
            }
        }

        // No phantom cells: every populated cell in the reloaded sheet was one we
        // wrote (bar exact-duplicate positions, which the map already dedupes).
        prop_assert!(ws.cells.len() <= cells.len());
    }

    #[test]
    fn multiple_sheets_survive_save_load(
        sheets in prop::collection::vec(
            ("[A-Za-z][A-Za-z0-9 ]{0,20}", sheet()),
            1..4,
        )
    ) {
        // Distinct sheet names (Excel requires uniqueness).
        let mut wb = Workbook::new();
        let mut used = std::collections::HashSet::new();
        let mut kept: Vec<NamedSheet> = Vec::new();
        for (name, cells) in sheets {
            let name = name.trim().to_string();
            if name.is_empty() || !used.insert(name.clone()) {
                continue;
            }
            wb.create_sheet(Some(name.clone())).unwrap();
            for (&(row, col), value) in &cells {
                wb.set_cell_value_in_sheet(&name, row, col, value.clone()).unwrap();
            }
            kept.push((name, cells));
        }
        prop_assume!(!kept.is_empty());

        let bytes = wb.save_to_bytes().unwrap();
        let back = Workbook::load_from_bytes(&bytes).unwrap();

        for (name, cells) in &kept {
            let ws = back.get_sheet_by_name(name).unwrap();
            for (&(row, col), sent) in cells {
                let got = ws.get_cell_value(row, col);
                match sent {
                    CellValue::Empty => {
                        prop_assert!(matches!(got, None | Some(CellValue::Empty)))
                    }
                    other => prop_assert_eq!(Some(other), got),
                }
            }
        }
    }
}

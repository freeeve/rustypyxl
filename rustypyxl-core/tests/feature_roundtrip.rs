//! Save->load round-trips for the feature modules. The writer emits the full
//! model for tables and autofilters; the parser has to read all of it back, or
//! loading and re-saving an Excel file silently drops what it did not model.

use rustypyxl_core::autofilter::{
    AutoFilter, CustomFilter, DynamicFilterType, FilterColumn, FilterOperator, FilterType,
    Top10Filter,
};
use rustypyxl_core::table::{Table, TableColumn};
use rustypyxl_core::worksheet::DataValidation;
use rustypyxl_core::{CellValue, Workbook};

fn roundtrip(wb: &Workbook) -> Workbook {
    Workbook::load_from_bytes(&wb.save_to_bytes().unwrap()).unwrap()
}

/// The writer emitted only type/allowBlank/show*/sqref, so a validation with a
/// custom error dialog saved without one: Excel showed its generic message
/// instead of the rule's own wording.
#[test]
fn data_validation_messages_survive_roundtrip() {
    let mut wb = Workbook::new();
    let ws = wb.create_sheet(Some("Sheet1".to_string())).unwrap();

    let validation = DataValidation {
        validation_type: "whole".to_string(),
        operator: Some("greaterThan".to_string()),
        formula1: Some("10".to_string()),
        formula2: None,
        error_style: Some("warning".to_string()),
        error_title: Some("Too small".to_string()),
        error_message: Some("Enter a number above 10.".to_string()),
        prompt_title: Some("Quantity".to_string()),
        prompt_message: Some("How many units?".to_string()),
        allow_blank: false,
        show_error: true,
        show_input: true,
        sqref: None,
    };
    ws.add_data_validation(1, 1, validation);

    let reloaded = roundtrip(&wb);
    let ws = reloaded.get_sheet_by_name("Sheet1").unwrap();
    let (_, dv) = ws
        .data_validations
        .iter()
        .next()
        .expect("validation preserved");

    assert_eq!(dv.validation_type, "whole");
    assert_eq!(dv.operator.as_deref(), Some("greaterThan"));
    assert_eq!(dv.formula1.as_deref(), Some("10"));
    assert_eq!(dv.error_style.as_deref(), Some("warning"));
    assert_eq!(dv.error_title.as_deref(), Some("Too small"));
    assert_eq!(
        dv.error_message.as_deref(),
        Some("Enter a number above 10.")
    );
    assert_eq!(dv.prompt_title.as_deref(), Some("Quantity"));
    assert_eq!(dv.prompt_message.as_deref(), Some("How many units?"));
    assert!(!dv.allow_blank);
}

/// A validation with no dialog text must not grow empty attributes.
#[test]
fn data_validation_without_messages_omits_the_attributes() {
    let mut wb = Workbook::new();
    let ws = wb.create_sheet(Some("Sheet1".to_string())).unwrap();
    ws.add_data_validation(
        1,
        1,
        DataValidation {
            validation_type: "list".to_string(),
            formula1: Some("\"Yes,No\"".to_string()),
            ..Default::default()
        },
    );

    let reloaded = roundtrip(&wb);
    let (_, dv) = reloaded
        .get_sheet_by_name("Sheet1")
        .unwrap()
        .data_validations
        .iter()
        .next()
        .unwrap();

    assert_eq!(dv.validation_type, "list");
    assert_eq!(dv.formula1.as_deref(), Some("\"Yes,No\""));
    assert_eq!(dv.operator, None);
    assert_eq!(dv.error_title, None);
    assert_eq!(dv.error_style, None);
}

/// The writer emits calculatedColumnFormula as a child element (correct
/// OOXML); the parser used to look only for an attribute of that name, so the
/// formula came back as None.
#[test]
fn table_calculated_column_formula_survives_roundtrip() {
    let mut wb = Workbook::new();
    let ws = wb.create_sheet(Some("Sheet1".to_string())).unwrap();
    ws.set_cell_value(1, 1, CellValue::String("Qty".into()));
    ws.set_cell_value(1, 2, CellValue::String("Total".into()));
    ws.set_cell_value(2, 1, CellValue::Number(2.0));
    ws.set_cell_value(2, 2, CellValue::Number(4.0));

    let mut table = Table::new(1, "Table1", "A1:B2");
    table.columns = vec![
        TableColumn::new(1, "Qty"),
        TableColumn::new(2, "Total").with_formula("Table1[[#This Row],[Qty]]*2"),
    ];
    ws.add_table(table);

    let reloaded = roundtrip(&wb);
    let table = &reloaded.get_sheet_by_name("Sheet1").unwrap().tables[0];

    assert_eq!(table.columns.len(), 2);
    assert_eq!(table.columns[0].calculated_column_formula, None);
    assert_eq!(
        table.columns[1].calculated_column_formula.as_deref(),
        Some("Table1[[#This Row],[Qty]]*2"),
        "calculated column formula must survive a save/load cycle"
    );
}

/// Value filters ("show only these entries") are the common case.
#[test]
fn autofilter_value_criteria_survive_roundtrip() {
    let mut wb = Workbook::new();
    let ws = wb.create_sheet(Some("Sheet1".to_string())).unwrap();
    ws.set_cell_value(1, 1, CellValue::String("Fruit".into()));

    let mut af = AutoFilter::new("A1:B10");
    af.add_filter(FilterColumn::values(
        0,
        vec!["Apple".to_string(), "Orange".to_string()],
    ));
    ws.auto_filter = Some(af);

    let reloaded = roundtrip(&wb);
    let af = reloaded
        .get_sheet_by_name("Sheet1")
        .unwrap()
        .auto_filter
        .as_ref()
        .expect("autofilter preserved");

    assert_eq!(af.range, "A1:B10");
    assert_eq!(af.columns.len(), 1, "filter criteria must not be dropped");
    assert_eq!(af.columns[0].column_id, 0);
    match &af.columns[0].filter {
        FilterType::Values(values) => {
            assert_eq!(values, &["Apple".to_string(), "Orange".to_string()])
        }
        other => panic!("expected a value filter, got {:?}", other),
    }
}

/// Custom filters carry two operators joined by AND or OR.
#[test]
fn autofilter_custom_criteria_survive_roundtrip() {
    let mut wb = Workbook::new();
    let ws = wb.create_sheet(Some("Sheet1".to_string())).unwrap();

    let mut af = AutoFilter::new("A1:C20");
    af.add_filter(FilterColumn::custom(
        1,
        CustomFilter::new(FilterOperator::GreaterThan, "100")
            .or(FilterOperator::LessThanOrEqual, "5"),
    ));
    ws.auto_filter = Some(af);

    let reloaded = roundtrip(&wb);
    let af = reloaded
        .get_sheet_by_name("Sheet1")
        .unwrap()
        .auto_filter
        .as_ref()
        .unwrap();

    assert_eq!(af.columns.len(), 1);
    assert_eq!(af.columns[0].column_id, 1);
    match &af.columns[0].filter {
        FilterType::Custom(custom) => {
            assert_eq!(custom.operator1, FilterOperator::GreaterThan);
            assert_eq!(custom.value1, "100");
            assert!(!custom.and, "the OR join must survive");
            assert_eq!(custom.operator2, Some(FilterOperator::LessThanOrEqual));
            assert_eq!(custom.value2.as_deref(), Some("5"));
        }
        other => panic!("expected a custom filter, got {:?}", other),
    }
}

/// Dynamic and top-10 filters, plus the sort state and hidden-button flag.
#[test]
fn autofilter_dynamic_top10_and_sort_state_survive_roundtrip() {
    let mut wb = Workbook::new();
    let ws = wb.create_sheet(Some("Sheet1".to_string())).unwrap();

    let mut af = AutoFilter::new("A1:D50");
    af.add_filter(FilterColumn {
        column_id: 0,
        filter: FilterType::DynamicFilter(DynamicFilterType::ThisMonth),
        show_button: true,
    });
    af.add_filter(FilterColumn {
        column_id: 2,
        filter: FilterType::Top10Filter(Top10Filter::top_percent(25.0)),
        show_button: false,
    });
    af.sort_by(1, true);
    ws.auto_filter = Some(af);

    let reloaded = roundtrip(&wb);
    let af = reloaded
        .get_sheet_by_name("Sheet1")
        .unwrap()
        .auto_filter
        .as_ref()
        .unwrap();

    assert_eq!(af.columns.len(), 2);
    assert_eq!(
        af.columns[0].filter,
        FilterType::DynamicFilter(DynamicFilterType::ThisMonth)
    );

    match &af.columns[1].filter {
        FilterType::Top10Filter(top10) => {
            assert!(top10.top);
            assert!(top10.percent);
            assert_eq!(top10.value, 25.0);
        }
        other => panic!("expected a top10 filter, got {:?}", other),
    }
    assert!(
        !af.columns[1].show_button,
        "hiddenButton must survive the round-trip"
    );

    assert_eq!(af.sort_column, Some(1));
    assert!(af.sort_descending);
}

/// A plain autofilter with no criteria still round-trips as a bare range.
#[test]
fn autofilter_without_criteria_roundtrips_as_range() {
    let mut wb = Workbook::new();
    let ws = wb.create_sheet(Some("Sheet1".to_string())).unwrap();
    ws.auto_filter = Some(AutoFilter::new("A1:C5"));

    let reloaded = roundtrip(&wb);
    let af = reloaded
        .get_sheet_by_name("Sheet1")
        .unwrap()
        .auto_filter
        .as_ref()
        .unwrap();

    assert_eq!(af.range, "A1:C5");
    assert!(af.columns.is_empty());
    assert_eq!(af.sort_column, None);
}

# Table and autofilter data lost on load->save round-trip

Round-trip losses in the feature-module wiring (parser side).

- Table calculated-column formulas are lost on reload. The writer emits
  `<calculatedColumnFormula>` as a child element with text content, which is
  correct OOXML (writer.rs:2182-2193), but the parser only looks for a
  `calculatedColumnFormula` attribute on `<tableColumn>` (workbook.rs:2202)
  and never reads the child element. Any table saved with
  `TableColumn::with_formula(...)`, or any Excel-authored table with
  calculated columns, comes back with `calculated_column_formula == None`.
  Fix the parser to read the child element text; add a save->load round-trip
  test for tables with formulas.

- AutoFilter criteria are dropped on load: parse_autofilter_attrs captures
  only the `ref` range (workbook.rs:1888-1898), while the writer emits full
  `<filterColumn>`/`<customFilters>`/`<sortState>` (writer.rs:1547-1632).
  Loading an Excel file with an active filter and saving it silently clears
  the filter criteria. The code comments note filter columns are "not yet
  modeled"--model them, or at minimum preserve the raw child XML verbatim so
  round-trips are lossless.

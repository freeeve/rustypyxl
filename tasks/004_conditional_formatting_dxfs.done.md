# Conditional formatting rules never apply their formats (no dxfs) — DONE

Completed:

- `ConditionalRule.format` is now serialized: `collect_dxfs` gathers the
  deduplicated differential formats across all sheets, styles.xml writes a
  `<dxfs>` section (font, numFmt, fill, border), and each cfRule carries the
  matching `dxfId`. Rules now actually apply their formatting.
- Conditional formatting round-trips: `<dxfs>` and `<conditionalFormatting>`
  (cellIs/expression/colorScale/dataBar/iconSet/top10/text/timePeriod rules,
  cfvo thresholds, colors, formulas) are parsed on load, with dxfId resolved
  back into `ConditionalRule.format`.
- `IconSet` with empty thresholds now emits Excel-default percent bands
  (e.g. 0/33/66 for 3 icons) instead of schema-invalid `<iconSet/>` that
  triggered Excel repair.
- Text rules (contains/notContains/beginsWith/endsWith), blank/error rules,
  and time-period rules write the hidden anchor formula Excel requires to
  evaluate them; timePeriod rules write the required `timePeriod` attribute
  with generated formulas for all ten ST_TimePeriod values.
- aboveAverage rules write `equalAverage` and `stdDev`; constructors added
  for text/time-period/average/duplicate/unique rules, which previously
  could not be built without struct literals.

Verified by openpyxl differential tests (tests/test_roundtrip_structure.py
TestConditionalFormattingRoundtrip): openpyxl-authored rules survive a
rustypyxl load+save with their differential styles intact.

Not done (Excel 2010 x14 extensions, out of scope for the strict OOXML
namespace): DataBar border_color / gradient / negative_color — fields are
documented as not serialized.

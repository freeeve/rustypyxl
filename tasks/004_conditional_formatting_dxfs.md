# Conditional formatting rules never apply their formats (no dxfs)

`ConditionalRule.format` is collected but never serialized: the writer never
emits `dxfId` and styles.xml contains no `<dxfs>` section, so every
cellIs/expression/text/top10 rule applies no formatting (writer.rs
`write_conditional_formatting`).

Related corruption/completeness bugs in the same module:
- `IconSet::new()` defaults to zero thresholds and the writer emits
  `<iconSet>` with no `<cfvo>` children — schema requires >= 2, Excel flags
  the file as corrupt. Auto-generate default percent thresholds.
- Text rules (ContainsText/BeginsWith/...) emit the `text` attribute but not
  the companion `<formula>` Excel needs to evaluate the rule.
- `TimePeriod` rules never write the required `timePeriod` attribute;
  `equal_average`/`std_dev` fields are dead.
- `DataBar.border_color`/`gradient`/`negative_color` are dead fields.

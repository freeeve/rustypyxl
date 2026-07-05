# Load parser: cell placement and inline-string correctness

Correctness bugs in the worksheet XML parser (rustypyxl-core/src/workbook.rs).

## High priority

- Cells without an `r` attribute are mis-placed or silently overwrite each
  other. On `<c>` Start only value/type/style/formula are reset
  (workbook.rs:2568-2574); `current_col` is set only when `r` is present and
  parses, and is never reset per cell. `r` is optional in OOXML (column is
  inferred from position), so `<row r="2"><c><v>1</v></c><c><v>2</v></c></row>`
  reuses the stale column from the previous cell and both cells collide on one
  key--the second overwrites the first. The self-closing `<c/>` path instead
  drops such cells entirely, so the two paths are also inconsistent. Fix:
  maintain a position counter per row and infer row/col when `r` is absent.

- Multi-run inline strings (`<is>` rich text) lose all runs but the last.
  Each `<t>` overwrites `current_value` (workbook.rs:2778-2779):
  `current_value = Some(TempValue::String(text.into_owned()))`. The shared
  strings parser correctly concatenates runs via push_str
  (workbook.rs:1155-1159); the inline path should do the same.
  `<c t="inlineStr"><is><r><t>Hello </t></r><r><t>World</t></r></is></c>`
  currently yields "World".

## Low priority

- Row height depends on attribute order: `ht` is applied using `current_row`,
  which is only set if `r` was already seen in the same element
  (workbook.rs:2559-2565). If `ht` precedes `r`, height is dropped.
- Namespace-prefixed element names (`<x:t>`, `<x:si>`, `<x:c>`, ...) are not
  matched in the shared-strings and worksheet parsers (e.name() compared to
  bare bytes); workbook.xml and rels parsing already fall back to
  local_name()--make the others consistent.
- Cached formula `<v>` values are reformatted through ryu on save
  (workbook.rs:2920-2923), so `<v>5</v>` round-trips as "5.0". Cosmetic but
  alters output text.
- `<workbookPr date1904="1">` is never read. No data loss today (serials stay
  as Number), but any consumer interpreting serial dates is ~4 years off for
  1904-system files; at minimum preserve the flag on round-trip.

Add regression tests for each (spec-legal XML without `r`, inline rich text,
attribute-order permutations).

# Writer: XML sanitization and whitespace preservation gaps

Output-correctness gaps in rustypyxl-core/src/writer.rs. quick_xml's
BytesText/push_attribute entity-escape `< > & " '` but do NOT strip C0
control chars (0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F), which are illegal in
XML 1.0 even as entities. Cell values and shared strings are guarded
(escape_xml / strip_illegal_xml_chars, writer.rs:61, :585-587) but these
paths bypass the guard, so a stray `\x01` in dirty source data produces a
file Excel rejects as corrupt:

- Comment text (writer.rs:1517)
- Data-validation formula1/formula2 (writer.rs:1347, :1353)
- Defined-name range text (writer.rs:482)
- Header/footer text (writer.rs:2016, :2021)
- Table names/column names/calculated formulas (writer.rs:2131-2132, :2173, :2187)
- Conditional-formatting formulas (writer.rs:1826-1828)

Fix: route all user-supplied strings through the same sanitizer (element
text and attribute values), and add a test that writes each of these parts
with a control char and re-opens the file.

Also:

- `<t>` elements are emitted without `xml:space="preserve"` (shared strings
  writer.rs:582-590, inline strings :172-179, comments :1516-1520). The
  reader deliberately preserves leading/trailing whitespace, but conforming
  consumers may trim `<t>  hello  </t>` without the attribute; openpyxl and
  Excel add it whenever text has significant surrounding whitespace. Emit it
  conditionally.
- sharedStrings `count` is set equal to `uniqueCount` (writer.rs:577-579);
  `count` should be total references. Excel tolerates it--fix opportunistically.
- dxf numFmt ids use 200+idx (writer.rs:832), sharing the numbering space
  with workbook custom formats at 164+; a loaded file already using ids >=200
  could collide. Allocate from the same registry instead.

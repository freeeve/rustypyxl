# Lossy round-trip: load+save strips workbook structure

The loader only parses cells, merged cells, shared strings, styles, column
widths, row heights, sheet protection, comments, and (partially) hyperlinks.
Save regenerates the package from scratch, so loading any real-world xlsx and
saving it silently drops everything else. For an openpyxl-compatible library
this is silent data loss.

## Subtasks

- [ ] Hidden sheet state + active tab: parse `state` on `<sheet>` and
      `activeTab` on `<workbookView>`; write them back (currently hidden
      sheets become visible on save — data-exposure flavor).
- [ ] Freeze panes: parse `<pane topLeftCell=... state="frozen">` on load
      (writer already emits it).
- [ ] External hyperlinks: parse `xl/worksheets/_rels/sheetN.xml.rels` and
      resolve `r:id` to the real URL on load; on save write hyperlink
      relationships (`r:id` + TargetMode="External") instead of stuffing the
      URL into `location`, and add `xmlns:r` to the worksheet root.
      Currently any load+save destroys all external hyperlinks
      (workbook.rs hyperlink parse, writer.rs `write_worksheet_xml`).
- [ ] Comments: resolve the comments part via the sheet rels instead of the
      hardcoded `xl/comments/comment{sheet_id}.xml` path (real files use
      `xl/comments1.xml`); add the `[Content_Types].xml` override. The VML
      `legacyDrawing` part needed for Excel to display comments is a separate
      follow-up.
- [ ] AutoFilter: parse `<autoFilter ref>` on load (writer exists).
- [ ] Data validations: parse `<dataValidation>` (+ formulas, sqref ranges)
      on load; preserve multi-cell sqref through the writer.
- [ ] Page setup: parse `pageMargins`, `pageSetup`, `printOptions` on load
      (writer exists).
- [ ] Tables: `write_table_xml` exists but has zero call sites — wire table
      parts into save (`xl/tables/tableN.xml`, sheet `<tableParts>`, sheet
      rels, content-type override) and parse them on load. `add_table()` is
      currently a silent no-op.
- [ ] Conditional formatting: parse `<conditionalFormatting>` on load, and
      write a `<dxfs>` section in styles.xml wired to `dxfId` so rule formats
      actually apply (see 006).
- [ ] Sheet-scoped defined names: preserve `localSheetId`/`hidden` on
      defined names (they currently become global on save).
- [ ] Cached formula values: `<f>` is written without `<v>`, so viewers that
      don't recalculate show blanks; preserve the cached value.

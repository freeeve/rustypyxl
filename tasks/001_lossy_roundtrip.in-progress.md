# Lossy round-trip: load+save strips workbook structure

The loader only parsed cells and a few basics; save regenerates the package
from scratch, so loading any real-world xlsx and saving it silently dropped
everything else. For an openpyxl-compatible library this is silent data loss.

## Subtasks

- [x] Hidden sheet state + active tab: parsed and written back; exposed as
      `ws.sheet_state` / honored by `wb.active` in Python.
- [x] Freeze panes: `<pane>` parsed on load.
- [x] External hyperlinks: sheet .rels parsed, r:id resolved to real URLs on
      load; writer emits proper hyperlink relationships
      (TargetMode="External") with `xmlns:r` on the worksheet root.
- [x] Comments: part resolved via sheet rels (handles `xl/comments1.xml`);
      content-type override added. Remaining: the VML `legacyDrawing` part
      Excel needs to actually display comments.
- [x] AutoFilter: range parsed on load. Remaining: individual filter-column
      criteria are not modeled on load.
- [x] Data validations: parsed with formulas; multi-cell sqref preserved via
      `DataValidation::sqref`.
- [x] Page setup: pageMargins, pageSetup, printOptions parsed. Remaining:
      headerFooter text is written but not parsed back into sections.
- [x] Tables: write side fully wired (parts, tableParts, rels, content
      types, workbook-unique ids, totals-row-aware autoFilter ref) and
      parsed on load.
- [ ] Conditional formatting: parse `<conditionalFormatting>` on load, and
      write a `<dxfs>` section in styles.xml wired to `dxfId` so rule formats
      actually apply (see tasks/004).
- [ ] Sheet-scoped defined names: preserve `localSheetId`/`hidden` on
      defined names (they currently become global on save).
- [ ] Cached formula values: `<f>` is written without `<v>`, so viewers that
      don't recalculate show blanks; preserve the cached value.
- [ ] Comments VML (`legacyDrawing`) so Excel displays comment boxes.

Verified by differential tests (tests/test_roundtrip_structure.py): a
feature-rich openpyxl-authored file survives a rustypyxl load+save cycle
with openpyxl reading every feature back intact.

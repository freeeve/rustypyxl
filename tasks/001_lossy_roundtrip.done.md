# Lossy round-trip: load+save strips workbook structure — DONE

The loader only parsed cells and a few basics; save regenerated the package
from scratch, so loading any real-world xlsx and saving it silently dropped
everything else.

## Completed

- [x] Hidden sheet state + active tab (exposed as `ws.sheet_state`,
      honored by `wb.active`).
- [x] Freeze panes parsed on load.
- [x] External hyperlinks: sheet .rels parsed, r:id resolved on load;
      writer emits proper hyperlink relationships (TargetMode="External").
- [x] Comments: part resolved via sheet rels; content-type override;
      legacy VML drawing part + legacyDrawing element + rel written so
      Excel actually displays comment boxes.
- [x] AutoFilter range parsed on load.
- [x] Data validations parsed with formulas; multi-cell sqref preserved.
- [x] Page setup: pageMargins, pageSetup, printOptions parsed; oddHeader/
      oddFooter text parsed back into &L/&C/&R sections.
- [x] Tables: written (parts, tableParts, rels, content types, unique ids,
      totals-aware autoFilter ref) and parsed on load.
- [x] Conditional formatting + dxfs both directions (tasks/004, done).
- [x] Sheet-scoped defined names: localSheetId/hidden preserved.
- [x] Cached formula values: `<v>` next to `<f>` captured into
      `CellData::cached_formula_value` and written back with the right
      t attribute.

## Known remaining gaps (deliberately out of scope, tracked elsewhere)

- AutoFilter per-column filter criteria are not modeled on load (the range
  round-trips; criteria from external files are dropped) — see tasks/009.
- DataBar border/gradient/negative colors are x14 extensions, documented
  as not serialized (tasks/004 notes).
- Charts, images, pivot tables: model-only stubs, never written or parsed.

Verified by differential tests (tests/test_roundtrip_structure.py, 18
tests): feature-rich openpyxl/Excel-authored files survive a rustypyxl
load+save cycle with openpyxl reading every feature back intact, including
data_only cached formula reads.

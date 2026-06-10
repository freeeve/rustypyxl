# PyO3 worksheet handles silently target the wrong sheet — DONE

Fixed by introducing stable sheet uids:

- Core: `Worksheet::uid` stamped from a monotonic, never-reused counter on
  the workbook (`allocate_sheet_uid` / `sheet_index_by_uid`); assigned on
  create_sheet and load, re-stamped on copy.
- PyWorksheet and PyCell store the uid and resolve index/name through it on
  every operation — a single lookup path replaced the index/name split.
- Stale handles (sheet removed) raise ValueError on both reads and writes
  instead of silently hitting whatever sheet shifted into the old index.
- Handles survive remove / create_sheet(index=...) / move_sheet / rename /
  copy_worksheet; `ws.title` returns the live name after renames through
  any handle.
- Cell setters propagate errors (previously `let _ =` discarded them while
  getters raised — write-silently-fails/read-raises asymmetry).
- `cell.number_format = None` clears the format including the style-xf side.

Verified by tests/test_handle_stability.py (10 tests covering every
mutation scenario from the original review repro).

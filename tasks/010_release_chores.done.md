# Release and repo chores — DONE (v0.5.0)

- [x] Version bumped to 0.5.0 (0.4.0 was already on PyPI and breaking
      changes landed since); tagged and released as v0.5.0.
- [x] cargo fmt'd the workspace as a standalone commit; ci.yml gates on
      `cargo fmt --all --check`.
- [x] Default CompressionLevel is Default (deflate 6), matching CLAUDE.md;
      loaded workbooks no longer forced to uncompressed saves.
- [x] publish.yml runs the full Rust+Python test suites before building
      any wheel.
- [x] sonar-project.properties: removed invalid sonar.language, version
      bumped.
- [x] pyproject author email fixed to match the git identity.
- [x] Removed tracked TODO_MEMORY_OPTIMIZATION.md, empty root src/ and
      benches/ dirs; CLAUDE.md benchmark references corrected.

Deferred to tasks/011: wheel matrix expansion, scheduled fuzzing in CI,
sonar tarpaulin handling, fuzz_xml retargeting.

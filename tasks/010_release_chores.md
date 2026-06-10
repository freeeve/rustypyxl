# Release and repo chores

- [ ] Tag v0.4.0 (workspace/pyproject/sonar all say 0.4.0 but tags stop at
      v0.3.1; publish.yml triggers on releases, so 0.4.0 likely never
      shipped). Per project convention, always tag semantic versions.
- [ ] `cargo fmt` the workspace (currently a ~3,900-line diff) as a
      standalone mechanical commit, then add `cargo fmt --check` to CI.
- [ ] Decide the default CompressionLevel: workbook.rs defaults to `None`
      ("for benchmarking"), shipping uncompressed files by default and
      contradicting CLAUDE.md which documents Deflate-6.
- [ ] publish.yml wheel matrix: no macOS x86_64, Linux aarch64, or musl
      wheels; no test gate inside the publish workflow.
- [ ] sonarcloud.yml: tarpaulin runs with continue-on-error (coverage
      failures invisible), cargo-tarpaulin uncached; sonar.language=rust is
      not a valid property.
- [ ] pyproject author email `eve.freeman@gmail.com` vs git
      `eve.freema@gmail.com` — confirm which is right.
- [ ] Remove tracked stale `TODO_MEMORY_OPTIMIZATION.md`; delete empty root
      `src/` and `benches/` dirs; CLAUDE.md references
      `benchmarks/benchmark_read.py`/`benchmark_write.py` which don't exist.
- [ ] Wire fuzzing into CI (short scheduled run); fuzz_xml still fuzzes
      quick-xml directly rather than library code.

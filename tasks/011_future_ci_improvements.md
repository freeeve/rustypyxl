# Future CI / packaging improvements (deferred from 010)

- Wheel matrix: add macOS x86_64 (macos-13), Linux aarch64, and musl
  targets to publish.yml (via maturin-action with manylinux containers);
  consider abi3 wheels to collapse the 15 per-version builds.
- Wire fuzzing into CI as a scheduled short-run smoke (cargo +nightly
  fuzz run, ~60s per target); fuzz_xml still fuzzes quick-xml directly
  rather than library code - retarget or remove it.
- sonarcloud.yml: tarpaulin runs with continue-on-error so coverage
  failures are invisible; cargo-tarpaulin install is uncached.
- Rust CI tests run release-only on ubuntu only; consider a debug run
  (overflow checks) and a Windows/macOS leg.

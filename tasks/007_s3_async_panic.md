# S3 helpers panic when called from an async context

Doc comments on `load_from_s3`/`save_to_s3` claim a runtime is created "if
one is not already running", but the code unconditionally does
`Runtime::new()` + `block_on` (s3.rs:128-142) — inside an existing tokio
runtime this panics ("Cannot block the current thread from within a
runtime"). Use `Handle::try_current()` and spawn-blocking when already
inside a runtime, or document the restriction honestly.

Also: full AWS config resolution + a fresh runtime per call (no client
reuse); SdkError formatted with `format!("{}", e)` loses the underlying S3
error code (use DisplayErrorContext or into_service_error()).

# S3 helpers panic when called from an async context — DONE

- block_on_s3 detects an ambient tokio runtime (Handle::try_current) and
  runs the future on a dedicated thread with its own runtime instead of
  panicking with "Cannot block the current thread from within a runtime".
  Regression test calls load_from_s3 from inside a runtime.
- SdkError formatted via DisplayErrorContext, so errors include the
  underlying S3 service code/message.
- Doc comments now honestly describe per-call config/credential
  resolution and point hot paths at the async functions with their own
  client management.

Not done (acceptable for the API shape): client caching across calls -
each call takes its own S3Config, so resolution per call is inherent;
documented instead.

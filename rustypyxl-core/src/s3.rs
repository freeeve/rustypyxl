//! S3 support for loading and saving workbooks.
//!
//! This module provides functionality to load and save Excel workbooks
//! directly from/to Amazon S3 buckets.

use crate::error::{Result, RustypyxlError};
use crate::workbook::Workbook;

use aws_config::BehaviorVersion;
use aws_sdk_s3::Client;
use aws_smithy_types::error::display::DisplayErrorContext;

/// Configuration for S3 operations.
#[derive(Clone, Debug, Default)]
pub struct S3Config {
    /// AWS region (e.g., "us-east-1"). If None, uses default region.
    pub region: Option<String>,
    /// Custom endpoint URL (for S3-compatible services like MinIO).
    pub endpoint_url: Option<String>,
    /// Force path-style addressing (required for some S3-compatible services).
    pub force_path_style: bool,
}

impl S3Config {
    /// Create a new S3Config with default settings.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the AWS region.
    pub fn with_region(mut self, region: impl Into<String>) -> Self {
        self.region = Some(region.into());
        self
    }

    /// Set a custom endpoint URL (for S3-compatible services).
    pub fn with_endpoint_url(mut self, url: impl Into<String>) -> Self {
        self.endpoint_url = Some(url.into());
        self
    }

    /// Enable path-style addressing.
    pub fn with_path_style(mut self) -> Self {
        self.force_path_style = true;
        self
    }
}

/// Create an S3 client with the given configuration.
async fn create_s3_client(config: Option<&S3Config>) -> Result<Client> {
    let mut aws_config_loader = aws_config::defaults(BehaviorVersion::latest());

    if let Some(cfg) = config {
        if let Some(ref region) = cfg.region {
            aws_config_loader =
                aws_config_loader.region(aws_sdk_s3::config::Region::new(region.clone()));
        }
    }

    let aws_config = aws_config_loader.load().await;

    let mut s3_config_builder = aws_sdk_s3::config::Builder::from(&aws_config);

    if let Some(cfg) = config {
        if let Some(ref endpoint) = cfg.endpoint_url {
            s3_config_builder = s3_config_builder.endpoint_url(endpoint);
        }
        if cfg.force_path_style {
            s3_config_builder = s3_config_builder.force_path_style(true);
        }
    }

    Ok(Client::from_conf(s3_config_builder.build()))
}

/// Load a workbook from S3.
pub async fn load_from_s3_async(
    bucket: &str,
    key: &str,
    config: Option<&S3Config>,
) -> Result<Workbook> {
    let client = create_s3_client(config).await?;

    let response = client
        .get_object()
        .bucket(bucket)
        .key(key)
        .send()
        .await
        .map_err(|e| {
            RustypyxlError::S3Error(format!(
                "Failed to get object from S3: {}",
                DisplayErrorContext(&e)
            ))
        })?;

    let data = response.body.collect().await.map_err(|e| {
        RustypyxlError::S3Error(format!(
            "Failed to read S3 response body: {}",
            DisplayErrorContext(&e)
        ))
    })?;

    Workbook::load_from_bytes(&data.into_bytes())
}

/// Save a workbook to S3.
pub async fn save_to_s3_async(
    workbook: &Workbook,
    bucket: &str,
    key: &str,
    config: Option<&S3Config>,
) -> Result<()> {
    let client = create_s3_client(config).await?;

    let data = workbook.save_to_bytes()?;

    client
        .put_object()
        .bucket(bucket)
        .key(key)
        .body(data.into())
        .content_type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .send()
        .await
        .map_err(|e| {
            RustypyxlError::S3Error(format!(
                "Failed to put object to S3: {}",
                DisplayErrorContext(&e)
            ))
        })?;

    Ok(())
}

/// Run an S3 future to completion from synchronous code. Calling
/// Runtime::block_on inside an existing tokio runtime panics ("Cannot
/// block the current thread from within a runtime"), so when already
/// inside one the future runs on a dedicated thread with its own runtime.
fn block_on_s3<F, T>(future: F) -> Result<T>
where
    F: std::future::Future<Output = Result<T>> + Send,
    T: Send,
{
    let run = || -> Result<T> {
        let rt = tokio::runtime::Runtime::new().map_err(|e| {
            RustypyxlError::S3Error(format!("Failed to create tokio runtime: {}", e))
        })?;
        rt.block_on(future)
    };

    if tokio::runtime::Handle::try_current().is_ok() {
        std::thread::scope(|scope| {
            scope.spawn(run).join().unwrap_or_else(|_| {
                Err(RustypyxlError::S3Error(
                    "S3 worker thread panicked".to_string(),
                ))
            })
        })
    } else {
        run()
    }
}

impl Workbook {
    /// Load a workbook from S3.
    ///
    /// Blocking wrapper around [`load_from_s3_async`]; safe to call both
    /// from plain synchronous code and from within a tokio runtime. AWS
    /// configuration and credentials are resolved on every call; use the
    /// async function with your own client-managing code for hot paths.
    pub fn load_from_s3(bucket: &str, key: &str, config: Option<S3Config>) -> Result<Self> {
        block_on_s3(load_from_s3_async(bucket, key, config.as_ref()))
    }

    /// Save the workbook to S3.
    ///
    /// Blocking wrapper around [`save_to_s3_async`]; safe to call both from
    /// plain synchronous code and from within a tokio runtime.
    pub fn save_to_s3(&self, bucket: &str, key: &str, config: Option<S3Config>) -> Result<()> {
        block_on_s3(save_to_s3_async(self, bucket, key, config.as_ref()))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The blocking wrappers previously panicked with "Cannot block the
    /// current thread from within a runtime" when called from async code.
    #[test]
    fn test_blocking_wrapper_inside_tokio_runtime_returns_error_not_panic() {
        let rt = tokio::runtime::Runtime::new().unwrap();
        rt.block_on(async {
            let config = S3Config::new()
                .with_region("us-east-1")
                // Nothing listens here, so the call fails fast at connect
                .with_endpoint_url("http://127.0.0.1:9")
                .with_path_style();
            let result = Workbook::load_from_s3("no-such-bucket", "key", Some(config));
            assert!(result.is_err(), "expected an S3 error, not a panic");
        });
    }

    #[test]
    fn test_s3_config() {
        let config = S3Config::new()
            .with_region("us-west-2")
            .with_endpoint_url("http://localhost:9000")
            .with_path_style();

        assert_eq!(config.region, Some("us-west-2".to_string()));
        assert_eq!(
            config.endpoint_url,
            Some("http://localhost:9000".to_string())
        );
        assert!(config.force_path_style);
    }
}

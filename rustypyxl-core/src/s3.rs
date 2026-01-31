//! S3 support for loading and saving workbooks.
//!
//! This module provides functionality to load and save Excel workbooks
//! directly from/to Amazon S3 buckets.

use crate::error::{Result, RustypyxlError};
use crate::workbook::Workbook;

use aws_config::BehaviorVersion;
use aws_sdk_s3::Client;

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
            aws_config_loader = aws_config_loader.region(aws_sdk_s3::config::Region::new(region.clone()));
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
        .map_err(|e| RustypyxlError::S3Error(format!("Failed to get object from S3: {}", e)))?;

    let data = response
        .body
        .collect()
        .await
        .map_err(|e| RustypyxlError::S3Error(format!("Failed to read S3 response body: {}", e)))?;

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
        .map_err(|e| RustypyxlError::S3Error(format!("Failed to put object to S3: {}", e)))?;

    Ok(())
}

impl Workbook {
    /// Load a workbook from S3.
    ///
    /// This is a blocking wrapper around the async S3 load operation.
    /// It creates a tokio runtime internally if one is not already running.
    pub fn load_from_s3(bucket: &str, key: &str, config: Option<S3Config>) -> Result<Self> {
        let rt = tokio::runtime::Runtime::new()
            .map_err(|e| RustypyxlError::S3Error(format!("Failed to create tokio runtime: {}", e)))?;
        rt.block_on(load_from_s3_async(bucket, key, config.as_ref()))
    }

    /// Save the workbook to S3.
    ///
    /// This is a blocking wrapper around the async S3 save operation.
    /// It creates a tokio runtime internally if one is not already running.
    pub fn save_to_s3(&self, bucket: &str, key: &str, config: Option<S3Config>) -> Result<()> {
        let rt = tokio::runtime::Runtime::new()
            .map_err(|e| RustypyxlError::S3Error(format!("Failed to create tokio runtime: {}", e)))?;
        rt.block_on(save_to_s3_async(self, bucket, key, config.as_ref()))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_s3_config() {
        let config = S3Config::new()
            .with_region("us-west-2")
            .with_endpoint_url("http://localhost:9000")
            .with_path_style();

        assert_eq!(config.region, Some("us-west-2".to_string()));
        assert_eq!(config.endpoint_url, Some("http://localhost:9000".to_string()));
        assert!(config.force_path_style);
    }
}

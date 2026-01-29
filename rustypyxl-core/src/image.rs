//! Image support for Excel workbooks.
//!
//! This module provides structures for inserting and managing images in Excel files.

use std::path::PathBuf;

/// Supported image formats.
#[derive(Clone, Debug, PartialEq)]
pub enum ImageFormat {
    /// PNG image.
    Png,
    /// JPEG image.
    Jpeg,
    /// GIF image.
    Gif,
    /// Windows Enhanced Metafile.
    Emf,
    /// Windows Metafile.
    Wmf,
    /// BMP image.
    Bmp,
    /// TIFF image.
    Tiff,
}

impl ImageFormat {
    /// Get the file extension for this format.
    pub fn extension(&self) -> &'static str {
        match self {
            ImageFormat::Png => "png",
            ImageFormat::Jpeg => "jpeg",
            ImageFormat::Gif => "gif",
            ImageFormat::Emf => "emf",
            ImageFormat::Wmf => "wmf",
            ImageFormat::Bmp => "bmp",
            ImageFormat::Tiff => "tiff",
        }
    }

    /// Get the content type for this format.
    pub fn content_type(&self) -> &'static str {
        match self {
            ImageFormat::Png => "image/png",
            ImageFormat::Jpeg => "image/jpeg",
            ImageFormat::Gif => "image/gif",
            ImageFormat::Emf => "image/x-emf",
            ImageFormat::Wmf => "image/x-wmf",
            ImageFormat::Bmp => "image/bmp",
            ImageFormat::Tiff => "image/tiff",
        }
    }

    /// Detect format from file extension.
    pub fn from_extension(ext: &str) -> Option<Self> {
        match ext.to_lowercase().as_str() {
            "png" => Some(ImageFormat::Png),
            "jpg" | "jpeg" => Some(ImageFormat::Jpeg),
            "gif" => Some(ImageFormat::Gif),
            "emf" => Some(ImageFormat::Emf),
            "wmf" => Some(ImageFormat::Wmf),
            "bmp" => Some(ImageFormat::Bmp),
            "tif" | "tiff" => Some(ImageFormat::Tiff),
            _ => None,
        }
    }

    /// Detect format from magic bytes.
    pub fn from_bytes(data: &[u8]) -> Option<Self> {
        if data.len() < 8 {
            return None;
        }

        // PNG: 89 50 4E 47 0D 0A 1A 0A
        if data.starts_with(&[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]) {
            return Some(ImageFormat::Png);
        }

        // JPEG: FF D8 FF
        if data.starts_with(&[0xFF, 0xD8, 0xFF]) {
            return Some(ImageFormat::Jpeg);
        }

        // GIF: GIF87a or GIF89a
        if data.starts_with(b"GIF87a") || data.starts_with(b"GIF89a") {
            return Some(ImageFormat::Gif);
        }

        // BMP: BM
        if data.starts_with(b"BM") {
            return Some(ImageFormat::Bmp);
        }

        // TIFF: II or MM
        if data.starts_with(&[0x49, 0x49, 0x2A, 0x00]) || data.starts_with(&[0x4D, 0x4D, 0x00, 0x2A])
        {
            return Some(ImageFormat::Tiff);
        }

        None
    }
}

/// Anchor type for positioning images.
#[derive(Clone, Debug, PartialEq)]
pub enum ImageAnchorType {
    /// One cell anchor - image moves with cell but doesn't resize.
    OneCell,
    /// Two cell anchor - image moves and resizes with cells.
    TwoCell,
    /// Absolute anchor - fixed position.
    Absolute,
}

/// Position anchor for an image.
#[derive(Clone, Debug)]
pub struct ImageAnchor {
    /// Anchor type.
    pub anchor_type: ImageAnchorType,
    /// Starting cell (e.g., "A1").
    pub from_cell: String,
    /// Column offset from cell in EMUs.
    pub from_col_offset: u32,
    /// Row offset from cell in EMUs.
    pub from_row_offset: u32,
    /// Ending cell for two-cell anchors.
    pub to_cell: Option<String>,
    /// Column offset for ending cell.
    pub to_col_offset: u32,
    /// Row offset for ending cell.
    pub to_row_offset: u32,
}

impl ImageAnchor {
    /// Create a one-cell anchor at the specified cell.
    pub fn one_cell<S: Into<String>>(cell: S) -> Self {
        ImageAnchor {
            anchor_type: ImageAnchorType::OneCell,
            from_cell: cell.into(),
            from_col_offset: 0,
            from_row_offset: 0,
            to_cell: None,
            to_col_offset: 0,
            to_row_offset: 0,
        }
    }

    /// Create a two-cell anchor spanning from one cell to another.
    pub fn two_cell<S1: Into<String>, S2: Into<String>>(from: S1, to: S2) -> Self {
        ImageAnchor {
            anchor_type: ImageAnchorType::TwoCell,
            from_cell: from.into(),
            from_col_offset: 0,
            from_row_offset: 0,
            to_cell: Some(to.into()),
            to_col_offset: 0,
            to_row_offset: 0,
        }
    }

    /// Set offsets in pixels (converted to EMUs).
    pub fn with_offset_px(mut self, col_px: u32, row_px: u32) -> Self {
        // Approximately 9525 EMUs per pixel
        self.from_col_offset = col_px * 9525;
        self.from_row_offset = row_px * 9525;
        self
    }
}

/// An image in an Excel worksheet.
#[derive(Clone, Debug)]
pub struct Image {
    /// Image data (bytes).
    pub data: Vec<u8>,
    /// Image format.
    pub format: ImageFormat,
    /// Original file path (if loaded from file).
    pub source_path: Option<PathBuf>,
    /// Position anchor.
    pub anchor: ImageAnchor,
    /// Width in EMUs (914400 EMUs = 1 inch).
    pub width: u32,
    /// Height in EMUs.
    pub height: u32,
    /// Alternative text for accessibility.
    pub alt_text: Option<String>,
    /// Description.
    pub description: Option<String>,
    /// Name/title.
    pub name: Option<String>,
}

impl Image {
    /// Create a new image from bytes.
    pub fn from_bytes(data: Vec<u8>, anchor: ImageAnchor) -> Option<Self> {
        let format = ImageFormat::from_bytes(&data)?;
        Some(Image {
            data,
            format,
            source_path: None,
            anchor,
            width: 914400,  // 1 inch default
            height: 914400, // 1 inch default
            alt_text: None,
            description: None,
            name: None,
        })
    }

    /// Create a new image from a file path.
    pub fn from_file(path: &std::path::Path, anchor: ImageAnchor) -> std::io::Result<Self> {
        let data = std::fs::read(path)?;

        let format = path
            .extension()
            .and_then(|e| e.to_str())
            .and_then(ImageFormat::from_extension)
            .or_else(|| ImageFormat::from_bytes(&data))
            .ok_or_else(|| {
                std::io::Error::new(std::io::ErrorKind::InvalidData, "Unknown image format")
            })?;

        Ok(Image {
            data,
            format,
            source_path: Some(path.to_path_buf()),
            anchor,
            width: 914400,
            height: 914400,
            alt_text: None,
            description: None,
            name: None,
        })
    }

    /// Set the image size in inches.
    pub fn with_size_inches(mut self, width: f64, height: f64) -> Self {
        self.width = (width * 914400.0) as u32;
        self.height = (height * 914400.0) as u32;
        self
    }

    /// Set the image size in pixels (at 96 DPI).
    pub fn with_size_px(mut self, width: u32, height: u32) -> Self {
        // 9525 EMUs per pixel at 96 DPI
        self.width = width * 9525;
        self.height = height * 9525;
        self
    }

    /// Set alternative text.
    pub fn with_alt_text<S: Into<String>>(mut self, text: S) -> Self {
        self.alt_text = Some(text.into());
        self
    }

    /// Set the image name.
    pub fn with_name<S: Into<String>>(mut self, name: S) -> Self {
        self.name = Some(name.into());
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_image_format_detection() {
        // PNG magic bytes
        let png_data = vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        assert_eq!(ImageFormat::from_bytes(&png_data), Some(ImageFormat::Png));

        // JPEG magic bytes (need at least 8 bytes for detection)
        let jpeg_data = vec![0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46];
        assert_eq!(ImageFormat::from_bytes(&jpeg_data), Some(ImageFormat::Jpeg));

        // GIF magic bytes (need at least 8 bytes for detection)
        let gif_data = b"GIF89a\x00\x00".to_vec();
        assert_eq!(ImageFormat::from_bytes(&gif_data), Some(ImageFormat::Gif));
    }

    #[test]
    fn test_format_from_extension() {
        assert_eq!(ImageFormat::from_extension("png"), Some(ImageFormat::Png));
        assert_eq!(ImageFormat::from_extension("PNG"), Some(ImageFormat::Png));
        assert_eq!(ImageFormat::from_extension("jpg"), Some(ImageFormat::Jpeg));
        assert_eq!(ImageFormat::from_extension("jpeg"), Some(ImageFormat::Jpeg));
    }

    #[test]
    fn test_image_anchor() {
        let anchor = ImageAnchor::one_cell("A1");
        assert_eq!(anchor.anchor_type, ImageAnchorType::OneCell);
        assert_eq!(anchor.from_cell, "A1");

        let anchor = ImageAnchor::two_cell("A1", "C5");
        assert_eq!(anchor.anchor_type, ImageAnchorType::TwoCell);
        assert_eq!(anchor.to_cell, Some("C5".to_string()));
    }

    #[test]
    fn test_image_size() {
        let png_data = vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        let image = Image::from_bytes(png_data, ImageAnchor::one_cell("A1"))
            .unwrap()
            .with_size_inches(2.0, 1.5);

        assert_eq!(image.width, 1828800);  // 2 * 914400
        assert_eq!(image.height, 1371600); // 1.5 * 914400
    }

    #[test]
    fn test_content_type() {
        assert_eq!(ImageFormat::Png.content_type(), "image/png");
        assert_eq!(ImageFormat::Jpeg.content_type(), "image/jpeg");
    }
}

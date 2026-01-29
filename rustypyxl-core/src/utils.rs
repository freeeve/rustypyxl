//! Utility functions for coordinate parsing and conversion.

use crate::error::{Result, RustypyxlError};

/// Maximum column number in Excel (XFD = 16384).
pub const MAX_COLUMN: u32 = 16384;
/// Maximum row number in Excel.
pub const MAX_ROW: u32 = 1_048_576;

/// Parse an Excel cell coordinate from bytes (e.g., b"A1", b"AB123") into (row, column).
/// Row and column are 1-indexed. This is the fast path that avoids string allocation.
#[inline]
pub fn parse_coordinate_bytes(bytes: &[u8]) -> Option<(u32, u32)> {
    if bytes.is_empty() {
        return None;
    }

    let mut i = 0usize;
    let mut column: u32 = 0;

    // Parse column letters with overflow protection
    while i < bytes.len() {
        let b = bytes[i];
        let upper = match b {
            b'a'..=b'z' => b - 32,
            b'A'..=b'Z' => b,
            _ => break,
        };
        // Use checked arithmetic to prevent overflow
        column = column.checked_mul(26)?.checked_add((upper - b'A' + 1) as u32)?;
        // Early exit if column exceeds Excel's max
        if column > MAX_COLUMN {
            return None;
        }
        i += 1;
    }

    if i == 0 || i >= bytes.len() || column == 0 {
        return None;
    }

    // Parse row number with overflow protection
    let mut row: u32 = 0;
    while i < bytes.len() {
        let b = bytes[i];
        if !b.is_ascii_digit() {
            return None;
        }
        // Use checked arithmetic to prevent overflow
        row = row.checked_mul(10)?.checked_add((b - b'0') as u32)?;
        // Early exit if row exceeds Excel's max
        if row > MAX_ROW {
            return None;
        }
        i += 1;
    }

    if row == 0 {
        return None;
    }

    Some((row, column))
}

/// Parse an Excel cell coordinate (e.g., "A1", "AB123") into (row, column).
/// Row and column are 1-indexed.
pub fn parse_coordinate(coord: &str) -> Result<(u32, u32)> {
    let coord = coord.trim();
    parse_coordinate_bytes(coord.as_bytes())
        .ok_or_else(|| RustypyxlError::InvalidCoordinate(format!("Invalid coordinate: {}", coord)))
}

/// Parse a u32 directly from bytes without string allocation.
#[inline]
pub fn parse_u32_bytes(bytes: &[u8]) -> Option<u32> {
    if bytes.is_empty() {
        return None;
    }
    let mut result: u32 = 0;
    for &b in bytes {
        if !b.is_ascii_digit() {
            return None;
        }
        // Use checked arithmetic to prevent overflow
        result = result.checked_mul(10)?.checked_add((b - b'0') as u32)?;
    }
    Some(result)
}

/// Parse an f64 directly from bytes without string allocation.
/// Falls back to string parsing for complex cases.
#[inline]
pub fn parse_f64_bytes(bytes: &[u8]) -> Option<f64> {
    // Fast path for simple integers
    if !bytes.is_empty() && bytes.iter().all(|&b| b.is_ascii_digit()) {
        let mut result: f64 = 0.0;
        for &b in bytes {
            result = result * 10.0 + (b - b'0') as f64;
        }
        return Some(result);
    }
    // Fall back to standard parsing for floats
    std::str::from_utf8(bytes).ok()?.parse().ok()
}

/// Convert column letters (e.g., "A", "AB", "XFD") to column number (1-indexed).
pub fn letter_to_column(letters: &str) -> Result<u32> {
    let mut result: u32 = 0;
    let mut saw_letter = false;

    for &b in letters.as_bytes() {
        let upper = match b {
            b'a'..=b'z' => b - 32,
            b'A'..=b'Z' => b,
            _ => {
                return Err(RustypyxlError::InvalidCoordinate(
                    format!("Invalid character in column: {}", b as char)
                ))
            }
        };
        saw_letter = true;
        // Use checked arithmetic to prevent overflow
        result = result
            .checked_mul(26)
            .and_then(|r| r.checked_add((upper - b'A' + 1) as u32))
            .ok_or_else(|| RustypyxlError::InvalidCoordinate(
                format!("Column '{}' exceeds maximum", letters)
            ))?;
        // Validate against Excel's max column
        if result > MAX_COLUMN {
            return Err(RustypyxlError::InvalidCoordinate(
                format!("Column '{}' exceeds Excel maximum (XFD = {})", letters, MAX_COLUMN)
            ));
        }
    }

    if !saw_letter || result == 0 {
        return Err(RustypyxlError::InvalidCoordinate(
            "Empty column letters".to_string()
        ));
    }

    Ok(result)
}

/// Convert column number (1-indexed) to letters (e.g., 1 -> "A", 28 -> "AB").
pub fn column_to_letter(column: u32) -> String {
    let mut result = String::new();
    let mut col = column;

    while col > 0 {
        col -= 1;
        let letter = (b'A' + (col % 26) as u8) as char;
        result.insert(0, letter);
        col /= 26;
    }

    result
}

/// Create a cell coordinate string from row and column (1-indexed).
pub fn coordinate_from_row_col(row: u32, column: u32) -> String {
    format!("{}{}", column_to_letter(column), row)
}

/// Parse a range reference (e.g., "A1:B10") into start and end coordinates.
pub fn parse_range(range: &str) -> Result<((u32, u32), (u32, u32))> {
    let parts: Vec<&str> = range.split(':').collect();

    if parts.len() != 2 {
        return Err(RustypyxlError::InvalidCoordinate(
            format!("Invalid range format: {}", range)
        ));
    }

    let start = parse_coordinate(parts[0])?;
    let end = parse_coordinate(parts[1])?;

    Ok((start, end))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_coordinate() {
        assert_eq!(parse_coordinate("A1").unwrap(), (1, 1));
        assert_eq!(parse_coordinate("B2").unwrap(), (2, 2));
        assert_eq!(parse_coordinate("Z1").unwrap(), (1, 26));
        assert_eq!(parse_coordinate("AA1").unwrap(), (1, 27));
        assert_eq!(parse_coordinate("AB10").unwrap(), (10, 28));
        assert_eq!(parse_coordinate("XFD1048576").unwrap(), (1048576, 16384));
    }

    #[test]
    fn test_parse_coordinate_case_insensitive() {
        assert_eq!(parse_coordinate("a1").unwrap(), (1, 1));
        assert_eq!(parse_coordinate("Ab10").unwrap(), (10, 28));
    }

    #[test]
    fn test_parse_coordinate_errors() {
        assert!(parse_coordinate("").is_err());
        assert!(parse_coordinate("A").is_err());
        assert!(parse_coordinate("1").is_err());
        assert!(parse_coordinate("A0").is_err());
    }

    #[test]
    fn test_letter_to_column() {
        assert_eq!(letter_to_column("A").unwrap(), 1);
        assert_eq!(letter_to_column("Z").unwrap(), 26);
        assert_eq!(letter_to_column("AA").unwrap(), 27);
        assert_eq!(letter_to_column("AB").unwrap(), 28);
        assert_eq!(letter_to_column("XFD").unwrap(), 16384);
    }

    #[test]
    fn test_column_to_letter() {
        assert_eq!(column_to_letter(1), "A");
        assert_eq!(column_to_letter(26), "Z");
        assert_eq!(column_to_letter(27), "AA");
        assert_eq!(column_to_letter(28), "AB");
        assert_eq!(column_to_letter(16384), "XFD");
    }

    #[test]
    fn test_column_roundtrip() {
        for col in 1..=16384 {
            let letters = column_to_letter(col);
            assert_eq!(letter_to_column(&letters).unwrap(), col);
        }
    }

    #[test]
    fn test_parse_range() {
        let ((r1, c1), (r2, c2)) = parse_range("A1:B10").unwrap();
        assert_eq!((r1, c1), (1, 1));
        assert_eq!((r2, c2), (10, 2));
    }

    #[test]
    fn test_coordinate_from_row_col() {
        assert_eq!(coordinate_from_row_col(1, 1), "A1");
        assert_eq!(coordinate_from_row_col(10, 28), "AB10");
    }

    #[test]
    fn test_overflow_protection_column() {
        // Test the fuzz-discovered crash input: very long column names
        assert!(parse_coordinate("CCCccccc0").is_err());
        assert!(parse_coordinate("AAAAAAAAAA1").is_err());
        assert!(parse_coordinate("ZZZZZZZZZ1").is_err());
        // Column exceeds Excel max (XFD = 16384)
        assert!(parse_coordinate("XFE1").is_err());
        assert!(parse_coordinate("XFDA1").is_err());
    }

    #[test]
    fn test_overflow_protection_row() {
        // Row exceeds Excel max (1048576)
        assert!(parse_coordinate("A1048577").is_err());
        assert!(parse_coordinate("A9999999999").is_err());
        // Numeric overflow
        assert!(parse_coordinate("A99999999999999999999").is_err());
    }

    #[test]
    fn test_letter_to_column_overflow() {
        // Very long column name should fail, not overflow
        assert!(letter_to_column("AAAAAAAAAA").is_err());
        assert!(letter_to_column("ZZZZZZZZZ").is_err());
        // Column exceeds Excel max
        assert!(letter_to_column("XFE").is_err());
    }

    #[test]
    fn test_parse_u32_bytes_overflow() {
        // Should return None on overflow, not wrap
        assert!(parse_u32_bytes(b"99999999999999999999").is_none());
        // Valid numbers should work
        assert_eq!(parse_u32_bytes(b"123"), Some(123));
        assert_eq!(parse_u32_bytes(b"4294967295"), Some(u32::MAX));
    }
}

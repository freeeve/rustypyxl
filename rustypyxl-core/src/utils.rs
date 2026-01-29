//! Utility functions for coordinate parsing and conversion.

use crate::error::{Result, RustypyxlError};

/// Parse an Excel cell coordinate from bytes (e.g., b"A1", b"AB123") into (row, column).
/// Row and column are 1-indexed. This is the fast path that avoids string allocation.
#[inline]
pub fn parse_coordinate_bytes(bytes: &[u8]) -> Option<(u32, u32)> {
    if bytes.is_empty() {
        return None;
    }

    let mut i = 0usize;
    let mut column: u32 = 0;

    // Parse column letters
    while i < bytes.len() {
        let b = bytes[i];
        let upper = match b {
            b'a'..=b'z' => b - 32,
            b'A'..=b'Z' => b,
            _ => break,
        };
        column = column * 26 + (upper - b'A' + 1) as u32;
        i += 1;
    }

    if i == 0 || i >= bytes.len() || column == 0 {
        return None;
    }

    // Parse row number
    let mut row: u32 = 0;
    while i < bytes.len() {
        let b = bytes[i];
        if !b.is_ascii_digit() {
            return None;
        }
        row = row.wrapping_mul(10).wrapping_add((b - b'0') as u32);
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
        result = result.wrapping_mul(10).wrapping_add((b - b'0') as u32);
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
        result = result * 26 + (upper - b'A' + 1) as u32;
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
}

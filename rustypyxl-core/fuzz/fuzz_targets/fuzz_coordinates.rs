#![no_main]

//! Fuzz target for cell coordinate parsing in rustypyxl-core.
//!
//! Tests the coordinate parsing functions with arbitrary input to ensure:
//! - No panics on any input (graceful error handling)
//! - No integer overflow issues
//! - Consistent behavior between string and bytes variants
//! - Correct roundtrip behavior when applicable

use libfuzzer_sys::fuzz_target;
use arbitrary::Arbitrary;
use rustypyxl_core::{
    parse_coordinate,
    parse_coordinate_bytes,
    parse_u32_bytes,
    parse_f64_bytes,
    letter_to_column,
    column_to_letter,
    coordinate_from_row_col,
    parse_range,
};

/// Structured input for coordinate fuzzing that allows more targeted testing.
#[derive(Arbitrary, Debug)]
struct CoordinateFuzzInput {
    /// Raw bytes to test byte-based parsers
    raw_bytes: Vec<u8>,
    /// String input for string-based parsers
    string_input: String,
    /// Row number for roundtrip testing
    row: u32,
    /// Column number for roundtrip testing
    column: u32,
}

/// Test parse_coordinate with arbitrary string input.
/// Valid formats: "A1", "AB123", "XFD1048576" (max Excel coordinate)
fn fuzz_parse_coordinate(input: &str) {
    let result = parse_coordinate(input);

    // If parsing succeeds, verify the result is sensible
    if let Ok((row, col)) = result {
        // Row and column should be non-zero (1-indexed)
        assert!(row > 0, "Row should be 1-indexed, got 0");
        assert!(col > 0, "Column should be 1-indexed, got 0");

        // Excel max is XFD1048576 (row=1048576, col=16384)
        // But we shouldn't crash even if values exceed this
    }
}

/// Test parse_coordinate_bytes with arbitrary byte input.
/// This is the fast path used internally.
fn fuzz_parse_coordinate_bytes(input: &[u8]) {
    let result = parse_coordinate_bytes(input);

    if let Some((row, col)) = result {
        assert!(row > 0, "Row should be 1-indexed");
        assert!(col > 0, "Column should be 1-indexed");
    }
}

/// Verify consistency between string and bytes parsing.
fn fuzz_coordinate_consistency(input: &str) {
    let string_result = parse_coordinate(input);
    let bytes_result = parse_coordinate_bytes(input.as_bytes());

    // Both should succeed or fail together
    match (string_result, bytes_result) {
        (Ok(s), Some(b)) => {
            // When both succeed, results must match
            assert_eq!(s, b, "String and bytes parsing gave different results for {:?}", input);
        }
        (Err(_), None) => {
            // Both failed - expected for invalid input
        }
        (Ok(s), None) => {
            // This could happen if parse_coordinate trims whitespace
            // but parse_coordinate_bytes doesn't - that's acceptable
            let trimmed = input.trim();
            if trimmed == input {
                panic!("String parsing succeeded ({:?}) but bytes parsing failed for {:?}", s, input);
            }
        }
        (Err(_), Some(b)) => {
            panic!("Bytes parsing succeeded ({:?}) but string parsing failed for {:?}", b, input);
        }
    }
}

/// Test parse_u32_bytes with arbitrary byte input.
fn fuzz_parse_u32_bytes(input: &[u8]) {
    let result = parse_u32_bytes(input);

    if let Some(val) = result {
        // The implementation uses wrapping arithmetic, so any u32 is valid
        // Just ensure we got a result without panicking
        let _ = val;
    }
}

/// Test parse_f64_bytes with arbitrary byte input.
fn fuzz_parse_f64_bytes(input: &[u8]) {
    let result = parse_f64_bytes(input);

    if let Some(val) = result {
        // NaN and infinity are valid f64 values
        // Just ensure no panic occurred
        let _ = val.is_nan();
        let _ = val.is_infinite();
    }
}

/// Test letter_to_column with arbitrary string input.
fn fuzz_letter_to_column(input: &str) {
    let result = letter_to_column(input);

    if let Ok(col) = result {
        assert!(col > 0, "Column should be 1-indexed");
    }
}

/// Test column_to_letter with arbitrary column numbers.
fn fuzz_column_to_letter(column: u32) {
    // column_to_letter should handle any u32 without panicking
    let letters = column_to_letter(column);

    if column > 0 {
        // Valid column, verify roundtrip
        if let Ok(back) = letter_to_column(&letters) {
            assert_eq!(back, column, "Roundtrip failed for column {}", column);
        }
    } else {
        // column=0 produces empty string, which is invalid
        assert!(letters.is_empty());
    }
}

/// Test coordinate_from_row_col with arbitrary row/column values.
fn fuzz_coordinate_from_row_col(row: u32, column: u32) {
    let coord = coordinate_from_row_col(row, column);

    // If both row and column are valid, test roundtrip
    if row > 0 && column > 0 {
        if let Ok((back_row, back_col)) = parse_coordinate(&coord) {
            assert_eq!(back_row, row, "Row roundtrip failed");
            assert_eq!(back_col, column, "Column roundtrip failed");
        }
    }
}

/// Test parse_range with arbitrary string input.
fn fuzz_parse_range(input: &str) {
    let result = parse_range(input);

    if let Ok(((r1, c1), (r2, c2))) = result {
        // All coordinates should be valid (1-indexed)
        assert!(r1 > 0 && c1 > 0 && r2 > 0 && c2 > 0);
    }
}

/// Test edge cases with specific patterns that might cause issues.
fn fuzz_edge_cases(input: &[u8]) {
    // Very long column letters (could overflow)
    if input.iter().all(|&b| b.is_ascii_alphabetic()) {
        let s = std::str::from_utf8(input).unwrap_or("");
        let _ = letter_to_column(s);
    }

    // Input with null bytes
    if input.contains(&0) {
        let _ = parse_coordinate_bytes(input);
    }

    // Input with only digits
    if input.iter().all(|&b| b.is_ascii_digit()) {
        let _ = parse_u32_bytes(input);
        let _ = parse_f64_bytes(input);
    }

    // Mixed case letters
    let _ = parse_coordinate_bytes(input);
}

fuzz_target!(|input: CoordinateFuzzInput| {
    // Test all coordinate parsing functions
    fuzz_parse_coordinate(&input.string_input);
    fuzz_parse_coordinate_bytes(&input.raw_bytes);
    fuzz_coordinate_consistency(&input.string_input);
    fuzz_parse_u32_bytes(&input.raw_bytes);
    fuzz_parse_f64_bytes(&input.raw_bytes);
    fuzz_letter_to_column(&input.string_input);
    fuzz_column_to_letter(input.column);
    fuzz_coordinate_from_row_col(input.row, input.column);
    fuzz_parse_range(&input.string_input);
    fuzz_edge_cases(&input.raw_bytes);
});

// A1 notation parser: cell reference and range parsing utilities

use crate::error::ExcelMcpError;

/// Maximum Excel column index (0-based). Column XFD = 16383.
const MAX_COL: u16 = 16383;
/// Maximum Excel row (1-based). Row 1048576 → 0-based index 1048575.
const MAX_ROW: u32 = 1_048_575;

/// A zero-based cell position.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellPos {
    /// 0-based row index (A1 row "1" → 0).
    pub row: u32,
    /// 0-based column index (A → 0, B → 1, …, XFD → 16383).
    pub col: u16,
}

/// A rectangular range defined by two cell positions (inclusive).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellRange {
    pub start: CellPos,
    pub end: CellPos,
}

/// Convert a column letter string (e.g. "A", "Z", "AA", "XFD") to a 0-based index.
///
/// Uses base-26 conversion where A=0, B=1, …, Z=25, AA=26, etc.
pub fn col_letter_to_index(s: &str) -> Result<u16, ExcelMcpError> {
    if s.is_empty() {
        return Err(ExcelMcpError::ParseError(
            "Column letter string is empty".into(),
        ));
    }

    let mut index: u32 = 0;
    for ch in s.chars() {
        if !ch.is_ascii_alphabetic() {
            return Err(ExcelMcpError::ParseError(format!(
                "Invalid character '{}' in column letters '{}'",
                ch, s
            )));
        }
        let val = ch.to_ascii_uppercase() as u32 - b'A' as u32;
        index = index * 26 + val + 1;
    }
    // Convert from 1-based to 0-based
    let index = index - 1;

    if index > MAX_COL as u32 {
        return Err(ExcelMcpError::InvalidInput(format!(
            "Column '{}' (index {}) exceeds maximum Excel column XFD ({})",
            s, index, MAX_COL
        )));
    }
    Ok(index as u16)
}

/// Convert a 0-based column index to a column letter string (e.g. 0 → "A", 25 → "Z", 26 → "AA").
pub fn index_to_col_letter(col: u16) -> String {
    let mut result = String::new();
    let mut n = col as u32 + 1; // convert to 1-based
    while n > 0 {
        n -= 1;
        let ch = (b'A' + (n % 26) as u8) as char;
        result.push(ch);
        n /= 26;
    }
    result.chars().rev().collect()
}

/// Convert a `CellPos` back to A1 notation (e.g. `CellPos { row: 0, col: 0 }` → "A1").
pub fn cell_pos_to_a1(pos: &CellPos) -> String {
    format!("{}{}", index_to_col_letter(pos.col), pos.row + 1)
}

/// Parse an A1-notation cell reference (e.g. "B3") into a `CellPos`.
///
/// Splits the string at the boundary between letters and digits, converts
/// the column letters via base-26, and converts the 1-based row to 0-based.
pub fn parse_cell_ref(s: &str) -> Result<CellPos, ExcelMcpError> {
    let s = s.trim();
    if s.is_empty() {
        return Err(ExcelMcpError::ParseError("Cell reference is empty".into()));
    }

    // Find the split point between letters and digits
    let split = s.find(|c: char| c.is_ascii_digit()).ok_or_else(|| {
        ExcelMcpError::ParseError(format!("Cell reference '{}' has no row number", s))
    })?;

    if split == 0 {
        return Err(ExcelMcpError::ParseError(format!(
            "Cell reference '{}' has no column letters",
            s
        )));
    }

    let col_str = &s[..split];
    let row_str = &s[split..];

    // Validate that the remaining part is all digits
    if !row_str.chars().all(|c| c.is_ascii_digit()) {
        return Err(ExcelMcpError::ParseError(format!(
            "Cell reference '{}' contains invalid characters in row portion '{}'",
            s, row_str
        )));
    }

    let col = col_letter_to_index(col_str)?;

    let row_1based: u32 = row_str.parse().map_err(|_| {
        ExcelMcpError::ParseError(format!(
            "Invalid row number '{}' in cell reference '{}'",
            row_str, s
        ))
    })?;

    if row_1based == 0 {
        return Err(ExcelMcpError::InvalidInput(format!(
            "Row number in '{}' must be >= 1",
            s
        )));
    }

    let row = row_1based - 1;
    if row > MAX_ROW {
        return Err(ExcelMcpError::InvalidInput(format!(
            "Row {} in '{}' exceeds maximum Excel row 1048576",
            row_1based, s
        )));
    }

    Ok(CellPos { row, col })
}

/// Parse a range reference in "A1:B2" notation into a `CellRange`.
pub fn parse_range_ref(s: &str) -> Result<CellRange, ExcelMcpError> {
    let s = s.trim();
    let parts: Vec<&str> = s.split(':').collect();
    if parts.len() != 2 {
        return Err(ExcelMcpError::ParseError(format!(
            "Range reference '{}' must contain exactly one ':'",
            s
        )));
    }

    let start = parse_cell_ref(parts[0])?;
    let end = parse_cell_ref(parts[1])?;

    Ok(CellRange { start, end })
}

#[cfg(test)]
mod tests {
    use super::*;
    use proptest::prelude::*;

    proptest! {
        /// **Validates: Requirements 7.5, 8.3**
        #[test]
        fn roundtrip_cell_ref(row in 0u32..1_048_576, col in 0u16..16384) {
            let pos = CellPos { row, col };
            let a1 = cell_pos_to_a1(&pos);
            let parsed = parse_cell_ref(&a1).unwrap();
            prop_assert_eq!(parsed.row, pos.row);
            prop_assert_eq!(parsed.col, pos.col);
        }
    }

    /// **Validates: Requirements 7.5, 8.3**
    #[test]
    fn test_parse_cell_ref_known() {
        let cases = [
            ("A1", 0, 0),
            ("B1", 0, 1),
            ("Z1", 0, 25),
            ("AA1", 0, 26),
            ("AB1", 0, 27),
            ("AZ1", 0, 51),
            ("BA1", 0, 52),
            ("XFD1", 0, 16383),
            ("A2", 1, 0),
            ("A1048576", 1_048_575, 0),
            ("C10", 9, 2),
        ];
        for (input, expected_row, expected_col) in cases {
            let pos = parse_cell_ref(input)
                .unwrap_or_else(|e| panic!("Failed to parse '{}': {}", input, e));
            assert_eq!(pos.row, expected_row, "Row mismatch for '{}'", input);
            assert_eq!(pos.col, expected_col, "Col mismatch for '{}'", input);
        }
    }

    /// **Validates: Requirements 7.5, 8.3**
    #[test]
    fn test_parse_cell_ref_invalid() {
        // Empty string
        assert!(parse_cell_ref("").is_err());

        // Digits only (no column letters)
        assert!(parse_cell_ref("123").is_err());

        // Letters only (no row number)
        assert!(parse_cell_ref("ABC").is_err());

        // Row 0 is invalid (Excel rows are 1-based)
        assert!(parse_cell_ref("A0").is_err());

        // Exceeding max row (1048576 is max, so 1048577 is invalid)
        assert!(parse_cell_ref("A1048577").is_err());

        // Exceeding max column (XFD = 16383, XFE would be 16384)
        assert!(parse_cell_ref("XFE1").is_err());

        // Non-alphabetic characters in column
        assert!(parse_cell_ref("1A1").is_err());
    }

    /// **Validates: Requirements 7.5, 8.3**
    #[test]
    fn test_parse_range_ref() {
        // Simple range
        let range = parse_range_ref("A1:B2").unwrap();
        assert_eq!(range.start, CellPos { row: 0, col: 0 });
        assert_eq!(range.end, CellPos { row: 1, col: 1 });

        // Full-sheet range
        let range = parse_range_ref("A1:XFD1048576").unwrap();
        assert_eq!(range.start, CellPos { row: 0, col: 0 });
        assert_eq!(
            range.end,
            CellPos {
                row: 1_048_575,
                col: 16383
            }
        );

        // Same cell range
        let range = parse_range_ref("C3:C3").unwrap();
        assert_eq!(range.start, CellPos { row: 2, col: 2 });
        assert_eq!(range.end, CellPos { row: 2, col: 2 });

        // Invalid: no colon
        assert!(parse_range_ref("A1B2").is_err());

        // Invalid: too many colons
        assert!(parse_range_ref("A1:B2:C3").is_err());

        // Invalid: bad cell ref in range
        assert!(parse_range_ref("A0:B2").is_err());
    }

    /// **Validates: Requirements 7.5, 8.3**
    #[test]
    fn test_col_letter_to_index() {
        assert_eq!(col_letter_to_index("A").unwrap(), 0);
        assert_eq!(col_letter_to_index("B").unwrap(), 1);
        assert_eq!(col_letter_to_index("Z").unwrap(), 25);
        assert_eq!(col_letter_to_index("AA").unwrap(), 26);
        assert_eq!(col_letter_to_index("AB").unwrap(), 27);
        assert_eq!(col_letter_to_index("AZ").unwrap(), 51);
        assert_eq!(col_letter_to_index("BA").unwrap(), 52);
        assert_eq!(col_letter_to_index("XFD").unwrap(), 16383);

        // Case insensitive
        assert_eq!(col_letter_to_index("a").unwrap(), 0);
        assert_eq!(col_letter_to_index("aa").unwrap(), 26);

        // Empty string
        assert!(col_letter_to_index("").is_err());

        // Invalid character
        assert!(col_letter_to_index("A1").is_err());

        // Exceeds max column
        assert!(col_letter_to_index("XFE").is_err());
    }

    /// **Validates: Requirements 7.5, 8.3**
    #[test]
    fn test_index_to_col_letter() {
        assert_eq!(index_to_col_letter(0), "A");
        assert_eq!(index_to_col_letter(1), "B");
        assert_eq!(index_to_col_letter(25), "Z");
        assert_eq!(index_to_col_letter(26), "AA");
        assert_eq!(index_to_col_letter(27), "AB");
        assert_eq!(index_to_col_letter(51), "AZ");
        assert_eq!(index_to_col_letter(52), "BA");
        assert_eq!(index_to_col_letter(16383), "XFD");
    }
}

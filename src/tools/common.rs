//! Shared helpers for tool modules.

use crate::store::WorkbookStore;
use crate::types::responses::{error, ErrorCategory};

/// Build a not-found error response that includes the list of open workbook IDs.
pub fn workbook_not_found(store: &WorkbookStore, id: &str) -> String {
    let open = store.open_ids();
    let ids_str = if open.is_empty() {
        "none".to_string()
    } else {
        open.join(", ")
    };
    error(
        ErrorCategory::NotFound,
        &format!("Workbook '{}' not found", id),
        &format!(
            "Currently open workbook IDs: {}. The workbook may have been closed or evicted due to inactivity.",
            ids_str
        ),
    )
}

/// Parse a comma-separated range string into individual `(row1, col1, row2, col2)` tuples.
///
/// Trims whitespace around each segment. Returns `Err` identifying the invalid segment.
pub fn parse_multi_range(range: &str) -> Result<Vec<(u32, u16, u32, u16)>, String> {
    let segments: Vec<&str> = range.split(',').map(|s| s.trim()).collect();
    let mut results = Vec::with_capacity(segments.len());
    for seg in &segments {
        let parsed = zavora_xlsx::utility::parse_range_ref(seg)
            .map_err(|e| format!("Invalid range segment '{}': {}", seg, e))?;
        results.push(parsed);
    }
    Ok(results)
}

/// Resolve a semantic format name to an Excel format code.
///
/// Returns the mapped code for known names, or the input unchanged for unknown strings.
pub fn resolve_semantic_format(name: &str) -> &str {
    match name {
        "currency" => "$#,##0.00",
        "percentage" => "0.0%",
        "accounting" => r#"_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)"#,
        "multiple" => r#"0.0"x""#,
        "date" => "yyyy-mm-dd",
        "number" => "#,##0",
        "integer" => "#,##0",
        "text" => "@",
        "decimal" => "#,##0.00",
        other => other,
    }
}

/// Convert a column letter string (e.g. "A", "AA") to a 0-based index.
fn col_letters_to_index(s: &str) -> Option<u16> {
    if s.is_empty() {
        return None;
    }
    let mut index: u32 = 0;
    for ch in s.chars() {
        if !ch.is_ascii_alphabetic() {
            return None;
        }
        let val = ch.to_ascii_uppercase() as u32 - b'A' as u32;
        index = index * 26 + val + 1;
    }
    let index = index.checked_sub(1)?;
    if index > 16383 {
        return None;
    }
    Some(index as u16)
}

/// Convert a 0-based column index back to column letter(s).
fn index_to_col_letters(col: u16) -> String {
    let mut result = String::new();
    let mut n = col as u32 + 1; // 1-based
    while n > 0 {
        n -= 1;
        let ch = (b'A' + (n % 26) as u8) as char;
        result.push(ch);
        n /= 26;
    }
    result.chars().rev().collect()
}

/// Adjust column references in a formula by a given offset.
///
/// Relative column references (e.g., `B10`) are shifted; absolute column references
/// (`$B$10`, `$B10`) are preserved. Row-only absolute (`B$10`) still shifts the column.
/// String literals (text inside double quotes) are left untouched.
/// Sheet references (`Sheet2!B10`) are handled — only the cell part adjusts.
pub fn adjust_formula_col_refs(formula: &str, col_offset: i16) -> String {
    if col_offset == 0 {
        return formula.to_string();
    }

    let chars: Vec<char> = formula.chars().collect();
    let len = chars.len();
    let mut result = String::with_capacity(formula.len());
    let mut i = 0;

    while i < len {
        // Skip string literals
        if chars[i] == '"' {
            result.push('"');
            i += 1;
            while i < len && chars[i] != '"' {
                result.push(chars[i]);
                i += 1;
            }
            if i < len {
                result.push('"');
                i += 1;
            }
            continue;
        }

        // Skip single-quoted sheet names (e.g., 'My Sheet'!A1)
        if chars[i] == '\'' {
            result.push('\'');
            i += 1;
            while i < len && chars[i] != '\'' {
                result.push(chars[i]);
                i += 1;
            }
            if i < len {
                result.push('\'');
                i += 1;
            }
            continue;
        }

        // Check for a cell reference pattern.
        // A cell reference can be preceded by `!` (sheet ref), or appear at start
        // or after a non-alphanumeric/non-$ character.
        let is_ref_start = i == 0
            || !chars[i - 1].is_ascii_alphanumeric()
            || (i > 0 && chars[i - 1] == '!');

        if is_ref_start {
            // Try to parse a cell reference: optional $ + column letters + optional $ + row digits
            let ref_start = i;
            let mut j = i;

            // Check for absolute column marker
            let col_is_absolute = j < len && chars[j] == '$';
            if col_is_absolute {
                j += 1;
            }

            // Collect column letters
            let col_start = j;
            while j < len && chars[j].is_ascii_alphabetic() {
                j += 1;
            }
            let col_end = j;
            let col_letters: String = chars[col_start..col_end].iter().collect();

            if !col_letters.is_empty() {
                // Check for optional absolute row marker
                let row_is_absolute = j < len && chars[j] == '$';
                if row_is_absolute {
                    j += 1;
                }

                // Collect row digits
                let row_start = j;
                while j < len && chars[j].is_ascii_digit() {
                    j += 1;
                }
                let row_end = j;

                // Valid cell reference requires at least one row digit
                if row_end > row_start {
                    // Make sure the character after the reference is not alphanumeric
                    // (to avoid matching partial identifiers like "SUM" in "SUMIF")
                    let after_ok =
                        j >= len || (!chars[j].is_ascii_alphabetic() && chars[j] != '_');

                    // Also ensure we're not in the middle of a sheet name before `!`
                    // If the next char is `!`, this is a sheet name, not a cell ref
                    let is_sheet_prefix = j < len && chars[j] == '!';

                    if after_ok && !is_sheet_prefix {
                        if col_is_absolute {
                            // Absolute column — emit unchanged
                            let original: String =
                                chars[ref_start..row_end].iter().collect();
                            result.push_str(&original);
                        } else {
                            // Relative column — shift by offset
                            if let Some(col_idx) = col_letters_to_index(&col_letters) {
                                let new_col =
                                    (col_idx as i32 + col_offset as i32).max(0) as u16;
                                let new_letters = index_to_col_letters(new_col);
                                result.push_str(&new_letters);
                                if row_is_absolute {
                                    result.push('$');
                                }
                                let row_digits: String =
                                    chars[row_start..row_end].iter().collect();
                                result.push_str(&row_digits);
                            } else {
                                // Can't parse column — emit unchanged
                                let original: String =
                                    chars[ref_start..row_end].iter().collect();
                                result.push_str(&original);
                            }
                        }
                        i = row_end;
                        continue;
                    }
                }
            }
        }

        // Not a cell reference — emit character as-is
        result.push(chars[i]);
        i += 1;
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;

    // ── resolve_semantic_format ──

    #[test]
    fn test_semantic_format_known_names() {
        assert_eq!(resolve_semantic_format("currency"), "$#,##0.00");
        assert_eq!(resolve_semantic_format("percentage"), "0.0%");
        assert_eq!(
            resolve_semantic_format("accounting"),
            r#"_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)"#
        );
        assert_eq!(resolve_semantic_format("multiple"), r#"0.0"x""#);
        assert_eq!(resolve_semantic_format("date"), "yyyy-mm-dd");
        assert_eq!(resolve_semantic_format("number"), "#,##0");
        assert_eq!(resolve_semantic_format("integer"), "#,##0");
        assert_eq!(resolve_semantic_format("text"), "@");
        assert_eq!(resolve_semantic_format("decimal"), "#,##0.00");
    }

    #[test]
    fn test_semantic_format_passthrough() {
        assert_eq!(resolve_semantic_format("$#,##0.00"), "$#,##0.00");
        assert_eq!(resolve_semantic_format("0.00%"), "0.00%");
        assert_eq!(resolve_semantic_format("custom_format"), "custom_format");
        assert_eq!(resolve_semantic_format(""), "");
    }

    // ── parse_multi_range ──

    #[test]
    fn test_comma_range_whitespace_trim() {
        let result = parse_multi_range(" A1:B5 , D1:E5 ").unwrap();
        assert_eq!(result.len(), 2);
        // A1:B5 → (0, 0, 4, 1)
        assert_eq!(result[0], (0, 0, 4, 1));
        // D1:E5 → (0, 3, 4, 4)
        assert_eq!(result[1], (0, 3, 4, 4));
    }

    #[test]
    fn test_comma_range_invalid_segment() {
        let result = parse_multi_range("A1:B5,INVALID,D1:E5");
        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(
            err.contains("INVALID"),
            "Error should identify the invalid segment: {}",
            err
        );
    }

    #[test]
    fn test_comma_range_single() {
        let result = parse_multi_range("A1:C3").unwrap();
        assert_eq!(result.len(), 1);
        assert_eq!(result[0], (0, 0, 2, 2));
    }

    // ── adjust_formula_col_refs ──

    #[test]
    fn test_adjust_basic_relative() {
        // B10 shifted by +1 → C10
        assert_eq!(adjust_formula_col_refs("B10", 1), "C10");
        // B10 shifted by +2 → D10
        assert_eq!(adjust_formula_col_refs("B10", 2), "D10");
    }

    #[test]
    fn test_adjust_absolute_column_preserved() {
        // $B$10 — absolute column, should not shift
        assert_eq!(adjust_formula_col_refs("$B$10", 1), "$B$10");
        // $B10 — absolute column, relative row — column should not shift
        assert_eq!(adjust_formula_col_refs("$B10", 1), "$B10");
    }

    #[test]
    fn test_adjust_relative_col_absolute_row() {
        // B$10 — relative column, absolute row — column DOES shift
        assert_eq!(adjust_formula_col_refs("B$10", 1), "C$10");
    }

    #[test]
    fn test_adjust_formula_expression() {
        assert_eq!(
            adjust_formula_col_refs("B10*(1+0.05)", 1),
            "C10*(1+0.05)"
        );
    }

    #[test]
    fn test_adjust_range_reference() {
        // SUM(B10:B20) shifted by +1 → SUM(C10:C20)
        assert_eq!(
            adjust_formula_col_refs("SUM(B10:B20)", 1),
            "SUM(C10:C20)"
        );
    }

    #[test]
    fn test_adjust_nested_functions() {
        // IF(B10>0, B10*C10, 0) shifted by +1 → IF(C10>0, C10*D10, 0)
        assert_eq!(
            adjust_formula_col_refs("IF(B10>0, B10*C10, 0)", 1),
            "IF(C10>0, C10*D10, 0)"
        );
    }

    #[test]
    fn test_adjust_string_literal_preserved() {
        // Text inside quotes should NOT be adjusted
        assert_eq!(
            adjust_formula_col_refs(r#""Cell B10""#, 1),
            r#""Cell B10""#
        );
    }

    #[test]
    fn test_adjust_sheet_reference() {
        // Sheet2!B10 shifted by +1 → Sheet2!C10
        assert_eq!(
            adjust_formula_col_refs("Sheet2!B10", 1),
            "Sheet2!C10"
        );
    }

    #[test]
    fn test_adjust_zero_offset() {
        assert_eq!(adjust_formula_col_refs("B10+C20", 0), "B10+C20");
    }

    #[test]
    fn test_adjust_negative_offset() {
        // C10 shifted by -1 → B10
        assert_eq!(adjust_formula_col_refs("C10", -1), "B10");
    }

    #[test]
    fn test_adjust_multi_letter_column() {
        // AA10 shifted by +1 → AB10
        assert_eq!(adjust_formula_col_refs("AA10", 1), "AB10");
    }

    #[test]
    fn test_adjust_mixed_absolute_relative() {
        // $A$1*B10 shifted by +1 → $A$1*C10
        assert_eq!(
            adjust_formula_col_refs("$A$1*B10", 1),
            "$A$1*C10"
        );
    }

    #[test]
    fn test_adjust_complex_formula() {
        // A realistic financial formula
        assert_eq!(
            adjust_formula_col_refs("B10-B12+$A$1*B15", 2),
            "D10-D12+$A$1*D15"
        );
    }

    // ── Property-based tests ──

    use proptest::prelude::*;

    /// All recognized semantic format names and their expected Excel codes.
    const KNOWN_FORMATS: &[(&str, &str)] = &[
        ("currency", "$#,##0.00"),
        ("percentage", "0.0%"),
        ("accounting", "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"),
        ("multiple", "0.0\"x\""),
        ("date", "yyyy-mm-dd"),
        ("number", "#,##0"),
        ("integer", "#,##0"),
        ("text", "@"),
        ("decimal", "#,##0.00"),
    ];

    // **Validates: Requirements 5.1, 5.3, 2.6**
    //
    // Property 1: Semantic Format Resolution Consistency
    // For any string input, if recognized semantic name → returns mapped Excel code;
    // if unrecognized → returns input unchanged.
    proptest! {
        #[test]
        fn prop_semantic_format_resolution_consistency(
            input in prop_oneof![
                Just("currency".to_string()),
                Just("percentage".to_string()),
                Just("accounting".to_string()),
                Just("multiple".to_string()),
                Just("date".to_string()),
                Just("number".to_string()),
                Just("integer".to_string()),
                Just("text".to_string()),
                Just("decimal".to_string()),
                "[a-zA-Z0-9#,.]{1,20}".prop_map(|s| s),
            ]
        ) {
            let result = resolve_semantic_format(&input);

            // Check if input is a known semantic name
            let known = KNOWN_FORMATS.iter().find(|(name, _)| *name == input.as_str());

            match known {
                Some((_, expected_code)) => {
                    // Known semantic name → must return the mapped Excel code
                    prop_assert_eq!(result, *expected_code,
                        "Known semantic name '{}' should map to '{}', got '{}'",
                        input, expected_code, result);
                }
                None => {
                    // Unrecognized string → must return input unchanged
                    prop_assert_eq!(result, input.as_str(),
                        "Unrecognized input '{}' should pass through unchanged, got '{}'",
                        input, result);
                }
            }
        }
    }

    // **Validates: Requirements 11.2, 11.3, 12.1, 12.2, 12.3**
    //
    // Property 6: Formula Reference Adjustment
    // For any formula and column offset N, relative column references shift by exactly N;
    // absolute column references ($) remain unchanged. String literals are not modified.

    /// Strategy to generate a single column letter in A-Z range.
    fn col_letter_strategy() -> impl Strategy<Value = String> {
        (0u8..26u8).prop_map(|i| String::from((b'A' + i) as char))
    }

    /// Strategy to generate a row number 1-1000.
    fn row_number_strategy() -> impl Strategy<Value = u32> {
        1u32..=1000u32
    }

    /// Strategy to generate a column offset that keeps references in valid range.
    fn offset_strategy() -> impl Strategy<Value = i16> {
        -10i16..=10i16
    }

    /// A single cell reference component for formula generation.
    #[derive(Debug, Clone)]
    enum CellRefKind {
        /// Relative column, relative row: e.g. B10
        RelRel(String, u32),
        /// Absolute column, absolute row: e.g. $B$10
        AbsAbs(String, u32),
        /// Absolute column, relative row: e.g. $B10
        AbsRel(String, u32),
        /// Relative column, absolute row: e.g. B$10
        RelAbs(String, u32),
    }

    impl CellRefKind {
        fn to_string_repr(&self) -> String {
            match self {
                CellRefKind::RelRel(col, row) => format!("{}{}", col, row),
                CellRefKind::AbsAbs(col, row) => format!("${}${}", col, row),
                CellRefKind::AbsRel(col, row) => format!("${}{}", col, row),
                CellRefKind::RelAbs(col, row) => format!("{}${}", col, row),
            }
        }

        fn col_is_absolute(&self) -> bool {
            matches!(self, CellRefKind::AbsAbs(_, _) | CellRefKind::AbsRel(_, _))
        }

        fn col_letters(&self) -> &str {
            match self {
                CellRefKind::RelRel(c, _)
                | CellRefKind::AbsAbs(c, _)
                | CellRefKind::AbsRel(c, _)
                | CellRefKind::RelAbs(c, _) => c,
            }
        }

        fn row(&self) -> u32 {
            match self {
                CellRefKind::RelRel(_, r)
                | CellRefKind::AbsAbs(_, r)
                | CellRefKind::AbsRel(_, r)
                | CellRefKind::RelAbs(_, r) => *r,
            }
        }

        fn row_is_absolute(&self) -> bool {
            matches!(self, CellRefKind::AbsAbs(_, _) | CellRefKind::RelAbs(_, _))
        }

        /// Compute the expected adjusted reference given an offset.
        fn expected_after_adjust(&self, offset: i16) -> String {
            if self.col_is_absolute() {
                // Absolute column references are unchanged
                return self.to_string_repr();
            }
            // Relative column — shift by offset
            let col_idx = col_letters_to_index(self.col_letters()).unwrap();
            let new_col = (col_idx as i32 + offset as i32).max(0) as u16;
            let new_letters = index_to_col_letters(new_col);
            if self.row_is_absolute() {
                format!("{}${}", new_letters, self.row())
            } else {
                format!("{}{}", new_letters, self.row())
            }
        }
    }

    /// Strategy to generate a cell reference of any kind.
    fn cell_ref_strategy() -> impl Strategy<Value = CellRefKind> {
        (col_letter_strategy(), row_number_strategy(), 0u8..4u8).prop_map(
            |(col, row, kind)| match kind {
                0 => CellRefKind::RelRel(col, row),
                1 => CellRefKind::AbsAbs(col, row),
                2 => CellRefKind::AbsRel(col, row),
                _ => CellRefKind::RelAbs(col, row),
            },
        )
    }

    proptest! {
        #[test]
        fn prop_formula_ref_adjustment_single_ref(
            cell_ref in cell_ref_strategy(),
            offset in offset_strategy(),
        ) {
            // For a formula consisting of a single cell reference,
            // verify the adjustment matches expectations.
            let formula = cell_ref.to_string_repr();
            let expected = cell_ref.expected_after_adjust(offset);
            let result = adjust_formula_col_refs(&formula, offset);
            prop_assert_eq!(result, expected);
        }

        #[test]
        fn prop_formula_ref_adjustment_mixed_refs(
            ref1 in cell_ref_strategy(),
            ref2 in cell_ref_strategy(),
            ref3 in cell_ref_strategy(),
            offset in offset_strategy(),
        ) {
            // Build a formula with mixed references separated by operators:
            // ref1+ref2*ref3
            let formula = format!(
                "{}+{}*{}",
                ref1.to_string_repr(),
                ref2.to_string_repr(),
                ref3.to_string_repr()
            );
            let expected = format!(
                "{}+{}*{}",
                ref1.expected_after_adjust(offset),
                ref2.expected_after_adjust(offset),
                ref3.expected_after_adjust(offset)
            );
            let result = adjust_formula_col_refs(&formula, offset);
            prop_assert_eq!(result, expected);
        }

        #[test]
        fn prop_formula_ref_adjustment_string_literal_preserved(
            col in col_letter_strategy(),
            row in row_number_strategy(),
            offset in offset_strategy(),
        ) {
            // References inside string literals must NOT be adjusted.
            let ref_str = format!("{}{}", col, row);
            let formula = format!("\"Cell {}\"", ref_str);
            let result = adjust_formula_col_refs(&formula, offset);
            prop_assert_eq!(result, formula);
        }

        #[test]
        fn prop_formula_ref_adjustment_sheet_reference(
            col in col_letter_strategy(),
            row in row_number_strategy(),
            offset in offset_strategy(),
        ) {
            // Sheet references: Sheet2!B10 — only the cell part adjusts.
            let col_idx = col_letters_to_index(&col).unwrap();
            let new_col = (col_idx as i32 + offset as i32).max(0) as u16;
            let new_letters = index_to_col_letters(new_col);

            let formula = format!("Sheet2!{}{}", col, row);
            let expected = format!("Sheet2!{}{}", new_letters, row);
            let result = adjust_formula_col_refs(&formula, offset);
            prop_assert_eq!(result, expected);
        }

        #[test]
        fn prop_formula_ref_adjustment_range_reference(
            col in col_letter_strategy(),
            row1 in 1u32..=500u32,
            row2 in 501u32..=1000u32,
            offset in offset_strategy(),
        ) {
            // Range references: SUM(col_row1:col_row2) — both endpoints adjust.
            let col_idx = col_letters_to_index(&col).unwrap();
            let new_col = (col_idx as i32 + offset as i32).max(0) as u16;
            let new_letters = index_to_col_letters(new_col);

            let formula = format!("SUM({}{}:{}{})", col, row1, col, row2);
            let expected = format!("SUM({}{}:{}{})", new_letters, row1, new_letters, row2);
            let result = adjust_formula_col_refs(&formula, offset);
            prop_assert_eq!(result, expected);
        }

        #[test]
        fn prop_formula_ref_adjustment_zero_offset_identity(
            ref1 in cell_ref_strategy(),
            ref2 in cell_ref_strategy(),
        ) {
            // Zero offset should always return the formula unchanged.
            let formula = format!("{}+{}", ref1.to_string_repr(), ref2.to_string_repr());
            let result = adjust_formula_col_refs(&formula, 0);
            prop_assert_eq!(result, formula);
        }

        #[test]
        fn prop_formula_ref_adjustment_absolute_col_never_changes(
            col in col_letter_strategy(),
            row in row_number_strategy(),
            offset in offset_strategy(),
            abs_row in proptest::bool::ANY,
        ) {
            // Absolute column references ($B10, $B$10) must never change.
            let formula = if abs_row {
                format!("${}${}", col, row)
            } else {
                format!("${}{}", col, row)
            };
            let result = adjust_formula_col_refs(&formula, offset);
            prop_assert_eq!(result, formula);
        }
    }

    // **Validates: Requirements 1.6**
    //
    // Property 3: Comma-Separated Range Atomicity
    // For any comma-separated range string with at least one invalid segment,
    // `parse_multi_range` returns an Err and the error identifies the invalid segment.
    // Valid ranges alone parse successfully.

    /// Strategy to generate a valid A1:B2-style range segment.
    fn valid_range_segment_strategy() -> impl Strategy<Value = String> {
        // Generate col letter (A-Z), row (1-100), col2 letter (A-Z), row2 (1-100)
        (
            (0u8..26u8).prop_map(|i| String::from((b'A' + i) as char)),
            1u32..=100u32,
            (0u8..26u8).prop_map(|i| String::from((b'A' + i) as char)),
            1u32..=100u32,
        )
            .prop_map(|(c1, r1, c2, r2)| {
                // Ensure start <= end for both row and col
                let (c1, c2) = if c1 <= c2 { (c1, c2) } else { (c2, c1) };
                let (r1, r2) = if r1 <= r2 { (r1, r2) } else { (r2, r1) };
                format!("{}{}:{}{}", c1, r1, c2, r2)
            })
    }

    /// Strategy to generate an invalid range segment that will fail parsing.
    fn invalid_range_segment_strategy() -> impl Strategy<Value = String> {
        prop_oneof![
            Just("INVALID".to_string()),
            Just("!!!".to_string()),
            Just("not_a_range".to_string()),
            Just("::".to_string()),
            Just("".to_string()),
            Just("ZZZ99999:ZZZ99999".to_string()),
            Just("123".to_string()),
            Just("@#$".to_string()),
        ]
    }

    proptest! {
        #[test]
        fn prop_comma_range_atomicity_invalid_segment_causes_error(
            valid_ranges in proptest::collection::vec(valid_range_segment_strategy(), 1..=4),
            invalid in invalid_range_segment_strategy(),
            insert_pos in 0usize..=4usize,
        ) {
            // Insert the invalid segment at a random position among valid ranges.
            let insert_pos = insert_pos.min(valid_ranges.len());
            let mut segments = valid_ranges.clone();
            segments.insert(insert_pos, invalid.clone());

            let input = segments.join(",");
            let result = parse_multi_range(&input);

            // Must be an error
            prop_assert!(result.is_err(),
                "Expected Err for input '{}' with invalid segment '{}', got Ok",
                input, invalid);

            // Error message must identify the invalid segment
            let err_msg = result.unwrap_err();
            prop_assert!(err_msg.contains(&invalid),
                "Error message '{}' should contain the invalid segment '{}'",
                err_msg, invalid);
        }

        #[test]
        fn prop_comma_range_valid_segments_parse_successfully(
            valid_ranges in proptest::collection::vec(valid_range_segment_strategy(), 1..=5),
        ) {
            // All-valid comma-separated ranges must parse successfully.
            let input = valid_ranges.join(",");
            let result = parse_multi_range(&input);

            prop_assert!(result.is_ok(),
                "Expected Ok for all-valid input '{}', got Err: {:?}",
                input, result.err());

            let parsed = result.unwrap();
            prop_assert_eq!(parsed.len(), valid_ranges.len(),
                "Parsed {} ranges but expected {} for input '{}'",
                parsed.len(), valid_ranges.len(), input);
        }
    }
}

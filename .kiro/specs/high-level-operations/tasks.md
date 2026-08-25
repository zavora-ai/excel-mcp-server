# Implementation Plan: High-Level Operations

## Overview

Implement ~20 new high-level MCP tools across three tiers (Styler, Writer, Data Operations) plus shared infrastructure. Tasks are ordered: shared infrastructure first, then Tier 1 (formatting), Tier 2 (writing), Tier 3 (data operations), with modifications to existing tools woven in where they naturally fit. Each task builds on previous work and ends with wiring into `server.rs`.

## Tasks

- [x] 1. Shared infrastructure: new enums, helpers, and module scaffolding
  - [x] 1.1 Add new enums to `src/types/enums.rs`
    - Add `SortDirection` (ascending, descending), `FillDirection` (down, right), `FillType` (linear, date, copy), `ConditionOperator` (equals, not_equals, contains, greater_than, less_than, starts_with, ends_with, is_empty)
    - All enums use `#[serde(rename_all = "snake_case")]` and derive `Deserialize, JsonSchema`
    - _Requirements: 14.1, 16.1, 17.2_

  - [x] 1.2 Add shared helper functions to `src/tools/common.rs`
    - Implement `parse_multi_range(range: &str) -> Result<Vec<(u32, u16, u32, u16)>, String>` — splits on comma, trims whitespace, parses each segment via `zavora_xlsx::utility::parse_range_ref`, returns error identifying invalid segment
    - Implement `resolve_semantic_format(name: &str) -> &str` — maps "currency" → "$#,##0.00", "percentage" → "0.0%", "accounting", "multiple", "date", "number", "integer", "text", "decimal"; unknown strings pass through unchanged
    - Implement `adjust_formula_col_refs(formula: &str, col_offset: i16) -> String` — tokenizes formula, shifts relative column references by offset, preserves absolute ($) references, handles sheet references, range references, nested functions, and string literals
    - _Requirements: 1.1-1.6, 5.1-5.3, 11.2-11.4, 12.1-12.3_

  - [x] 1.3 Write property test for semantic format resolution (Property 1)
    - **Property 1: Semantic Format Resolution Consistency**
    - For any string input, if recognized semantic name → returns mapped Excel code; if unrecognized → returns input unchanged
    - Generator: `prop_oneof![Just("currency"), Just("percentage"), ..., "[a-zA-Z0-9#,.]{1,20}"]`
    - **Validates: Requirements 5.1, 5.3, 2.6**

  - [x] 1.4 Write property test for formula reference adjustment (Property 6)
    - **Property 6: Formula Reference Adjustment**
    - For any formula and column offset N, relative column references shift by exactly N; absolute references ($) unchanged
    - Generator: formulas with mixed absolute/relative refs, random i16 offsets
    - Test edge cases: sheet references (`Sheet2!B10`), range references (`SUM(B10:B20)`), string literals, mixed `$B10` vs `B$10`
    - **Validates: Requirements 11.2, 11.3, 12.1, 12.2, 12.3**

  - [x] 1.5 Write property test for comma-separated range atomicity (Property 3)
    - **Property 3: Comma-Separated Range Atomicity**
    - For any comma-separated range string with at least one invalid segment, no formatting is applied and error identifies the invalid segment
    - Generator: valid ranges + one malformed range injected at random position
    - **Validates: Requirements 1.6**

  - [x] 1.6 Write unit tests for shared helpers
    - `test_semantic_format_known_names` — each semantic name maps correctly
    - `test_semantic_format_passthrough` — unknown strings pass through
    - `test_comma_range_whitespace_trim` — " A1:B5 , D1:E5 " parsed correctly
    - `test_comma_range_invalid_segment` — one bad segment returns error
    - _Requirements: 1.5, 1.6, 5.1-5.3_

  - [x] 1.3a Create `src/tools/data.rs` module file and register it in `src/tools/mod.rs`
    - Add `pub mod data;` to `src/tools/mod.rs`
    - Create empty `src/tools/data.rs` with module doc comment and standard imports
    - _Requirements: 14-22 (scaffolding)_

- [x] 2. Checkpoint — Ensure shared infrastructure compiles and all tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 3. Tier 1: Modify existing formatting tools for comma-separated ranges and semantic formats
  - [x] 3.1 Update `set_cell_format` in `src/tools/format.rs` to resolve semantic formats
    - Before applying `number_format`, call `resolve_semantic_format()` to map semantic names to Excel codes
    - Existing comma-separated range support is already in place; verify it uses `parse_multi_range` or equivalent with proper error handling per Requirement 1.6
    - _Requirements: 1.1, 1.5, 1.6, 5.1-5.4_

  - [x] 3.2 Update `merge_cells` in `src/tools/format.rs` for comma-separated ranges
    - Split range on comma, merge each range independently
    - Use `parse_multi_range` helper; if any segment invalid, return error without merging any
    - _Requirements: 1.2, 1.5, 1.6_

  - [x] 3.3 Update `add_conditional_format` in `src/tools/conditional.rs` for comma-separated ranges
    - Split range on comma, apply conditional format to each range independently
    - _Requirements: 1.3, 1.5, 1.6_

  - [x] 3.4 Update `set_dimensions` in `src/tools/expanded.rs` for comma-separated ranges
    - Split range on comma, apply dimension setting to each range independently
    - _Requirements: 1.4, 1.5, 1.6_

  - [x] 3.5 Update `write_cells` and `write_formula` tool descriptions in `src/server.rs`
    - Update `write_cells` description to state: "Strings starting with '=' are written as formulas with relative reference support. Use this for batch writes instead of write_formula for individual cells."
    - Update `write_formula` description to state: "For single formulas needing array or dynamic behavior. For regular formulas, prefer write_cells which auto-detects formulas."
    - _Requirements: 13.1, 13.2_

- [x] 4. Tier 1: Add new input/response structs for formatting tools
  - [x] 4.1 Add Tier 1 input structs to `src/types/inputs.rs`
    - Add `FormatOperation`, `BatchFormatInput`, `ApplyThemeInput`, `CopyFormatInput`, `ApplyStyleInput`, `FormatAsTableHeaderInput`, `FormatAsTableRangeInput`, `DescribeFormattingInput`
    - All structs use `#[serde(deny_unknown_fields)]`, standard `workbook_id` + `sheet_name` pattern
    - _Requirements: 2.1-2.7, 3.1-3.8, 4.1-4.5, 6.1-6.6, 7.1-7.6, 8.1-8.4, 9.1-9.4_

  - [x] 4.2 Add Tier 1 response structs to `src/types/responses.rs`
    - Add `BatchFormatResult`, `BatchFormatFailure`, `CopyFormatResult`, `DescribeFormattingResult`, `FormatGroup`
    - All structs derive `Debug, Serialize`
    - _Requirements: 2.4-2.5, 4.5, 9.1-9.4_

- [x] 5. Tier 1: Implement formatting tools
  - [x] 5.1 Implement `batch_format` in `src/tools/format.rs`
    - Accept array of `FormatOperation`, apply each sequentially
    - Each operation supports comma-separated ranges and semantic format resolution
    - On failure of one operation, continue processing remaining; collect failures with `{operation_index, range, error}`
    - Return `BatchFormatResult` with `operations_applied` count and `failures` array
    - _Requirements: 2.1-2.7_

  - [x] 5.2 Implement `apply_theme` in `src/tools/format.rs`
    - Support "financial_professional", "corporate", "minimal" themes
    - Apply header styling to specified `header_rows`, total styling to `total_rows`
    - Financial: bold white-on-dark-blue headers, alternating row shading, autofit
    - Corporate: bold dark headers, light gray bg, subtle borders, autofit
    - Minimal: bold headers, bottom border, no colors, autofit
    - Optional `auto_detect_formats` for currency column detection (opt-in)
    - Return error with valid theme names for unrecognized theme
    - _Requirements: 3.1-3.8_

  - [x] 5.3 Implement `copy_format` in `src/tools/format.rs`
    - Read formatting from source range cell-by-cell
    - Apply to each target range, tiling source formatting if target is larger
    - Copy: bold, italic, underline, font size, font color, background color, number format, alignment, border style
    - Return success note if source has no formatting
    - _Requirements: 4.1-4.5_

  - [x] 5.4 Write property test for copy format fidelity (Property 4)
    - **Property 4: Copy Format Fidelity**
    - Set random formatting on source, copy to target, read back and compare cell-by-cell
    - **Validates: Requirements 4.1, 4.3, 4.4**

  - [x] 5.5 Implement `apply_style` in `src/tools/format.rs`
    - Map style preset names to formatting bundles: "header", "title", "currency", "percentage", "date", "number", "text", "accounting", "total"
    - Support comma-separated ranges
    - Return error with valid preset names for unrecognized style
    - _Requirements: 8.1-8.4_

  - [x] 5.6 Implement `format_as_table_header` in `src/tools/format.rs`
    - Apply bold, white font, dark blue bg, center alignment to header row (default row 1)
    - Freeze panes at row below header
    - Set autofilter spanning header row from column A to last used column
    - Support optional `header_row`, `background_color`, `font_color` overrides
    - Return error if sheet is empty
    - _Requirements: 6.1-6.6_

  - [x] 5.7 Implement `format_as_table_range` in `src/tools/format.rs`
    - Apply header styling to first row of range (bold, bg color, white font)
    - Apply alternating row shading to data rows
    - Apply thin borders to all cells
    - Support color schemes: "blue" (default), "green", "gray", "orange"
    - Autofit columns within range
    - _Requirements: 7.1-7.6_

  - [x] 5.8 Write property test for table range consistency (Property 15)
    - **Property 15: Format as Table Range Consistency**
    - For any range with ≥2 rows, first row has header styling, all cells have borders, data rows have alternating shading
    - **Validates: Requirements 7.1, 7.2, 7.3**

  - [x] 5.9 Implement `describe_formatting` in `src/tools/read.rs`
    - Read formatting properties for each cell in range
    - Group cells with identical formatting, report ranges sharing each format
    - Return `DescribeFormattingResult` with `format_groups`
    - Return empty list if no formatting found
    - _Requirements: 9.1-9.4_

  - [x] 5.10 Write property test for batch format equivalence (Property 2)
    - **Property 2: Batch Format Equivalence**
    - For any array of 1-5 format operations, `batch_format` produces same result as sequential `set_cell_format` calls
    - **Validates: Requirements 2.1, 2.4**

  - [x] 5.11 Write unit tests for Tier 1 formatting tools
    - `test_batch_format_empty_operations`, `test_batch_format_partial_failure`
    - `test_apply_theme_financial`, `test_apply_theme_corporate`, `test_apply_theme_minimal`, `test_apply_theme_invalid_name`
    - `test_copy_format_no_formatting`, `test_copy_format_tiling`
    - `test_apply_style_each_preset`, `test_apply_style_invalid_name`
    - `test_format_table_header_defaults`, `test_format_table_header_custom_row`, `test_format_table_header_empty_sheet`
    - `test_format_table_range_default_blue`, `test_format_table_range_each_style`
    - `test_describe_formatting_empty`, `test_describe_formatting_grouped`
    - _Requirements: 2.1-2.7, 3.1-3.8, 4.1-4.5, 6.1-6.6, 7.1-7.6, 8.1-8.4, 9.1-9.4_

- [x] 6. Tier 1: Register formatting tools in `src/server.rs`
  - Add `batch_format`, `apply_theme`, `copy_format`, `apply_style`, `format_as_table_header`, `format_as_table_range`, `describe_formatting` tool registrations
  - Group under `// ── High-level formatting (7) ──` comment header
  - Each uses `#[tool(description = "...")]` and `tool_fn!` macro
  - _Requirements: 2-9_

- [x] 7. Checkpoint — Ensure Tier 1 compiles and all tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 8. Tier 2: Add input/response structs and implement writer tools
  - [x] 8.1 Add Tier 2 input structs to `src/types/inputs.rs`
    - Add `WriteGridInput`, `WriteRowRangeInput`, `CloneColumnFormulasInput`
    - _Requirements: 10.1-10.4, 11.1-11.6, 12.1-12.6_

  - [x] 8.2 Add Tier 2 response structs to `src/types/responses.rs`
    - Add `WriteGridResult`, `WriteRowRangeResult`, `CloneFormulasResult`
    - _Requirements: 10.3-10.4, 11.5, 12.5_

  - [x] 8.3 Implement `write_grid` in `src/tools/write.rs`
    - Write 2D array starting at `start_cell`, filling rightward and downward
    - Auto-detect types: "=" prefix → formula, numbers, booleans, ISO dates
    - Return `WriteGridResult` with rows_written, columns_written, cells_written
    - On cell write failure, continue and report failures array
    - _Requirements: 10.1-10.4_

  - [x] 8.4 Write property test for write grid round-trip (Property 5)
    - **Property 5: Write Grid Round-Trip**
    - Write random 2D arrays (1-10 rows, 1-10 cols) of numbers/strings/booleans, read back, verify equivalence and dimensions match
    - **Validates: Requirements 10.1, 10.3**

  - [x] 8.5 Implement `write_row_range` in `src/tools/write.rs`
    - Write formula at start cell, fill rightward to end column
    - Accept formula with or without leading "=" (strip if present)
    - Use `adjust_formula_col_refs` to shift relative references for each column
    - Preserve absolute references ($)
    - Return error if start column >= end column
    - _Requirements: 11.1-11.6_

  - [x] 8.6 Implement `clone_column_formulas` in `src/tools/write.rs`
    - Copy formulas from source column to each target column across row range
    - Use `adjust_formula_col_refs` with offset = target_col - source_col
    - Skip non-formula cells
    - Return `CloneFormulasResult` with formulas_cloned and columns_filled counts
    - _Requirements: 12.1-12.6_

  - [x] 8.7 Write unit tests for Tier 2 writer tools
    - `test_write_grid_mixed_types`, `test_write_grid_partial_failure`
    - `test_write_row_range_basic`, `test_write_row_range_absolute_refs`, `test_write_row_range_invalid_columns`
    - `test_clone_column_no_formulas`, `test_clone_column_skips_values`
    - _Requirements: 10.1-10.4, 11.1-11.6, 12.1-12.6_

- [x] 9. Tier 2: Register writer tools in `src/server.rs`
  - Add `write_grid`, `write_row_range`, `clone_column_formulas` tool registrations
  - Group under `// ── High-level writing (3) ──` comment header
  - _Requirements: 10-12_

- [x] 10. Checkpoint — Ensure Tier 2 compiles and all tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 11. Tier 3: Add input/response structs for data operations
  - [x] 11.1 Add Tier 3 input structs to `src/types/inputs.rs`
    - Add `SortKey`, `SortRangeInput`, `FindReplaceInput` (with `default_true` for `match_case`), `FillSeriesInput`, `RowCondition`, `DeleteRowsWhereInput`, `CopySheetInput`, `CopyRangeInput`, `TransposeRangeInput`, `RemoveDuplicatesInput`, `SplitColumnInput` (with `default_comma` for delimiter)
    - _Requirements: 14-22_

  - [x] 11.2 Add Tier 3 response structs to `src/types/responses.rs`
    - Add `SortResult`, `FindReplaceResult`, `FillSeriesResult`, `DeleteRowsResult`, `TransposeResult`, `RemoveDuplicatesResult`, `SplitColumnResult`
    - _Requirements: 14-22_

- [x] 12. Tier 3: Implement data operation tools
  - [x] 12.1 Implement `sort_range` in `src/tools/data.rs`
    - Read all cell values into `Vec<Vec<CellValue>>`, separate header if `has_header`
    - Stable sort with comparator: primary key first, then secondary; numbers numeric, strings lexicographic, empty cells last
    - Write sorted rows back (formatting stays on cell positions)
    - Tool description MUST document that formatting doesn't move with data
    - Return error if sort key column is outside range
    - _Requirements: 14.1-14.5_

  - [x] 12.2 Write property test for sort correctness (Property 7)
    - **Property 7: Sort Correctness**
    - For any data range and sort keys, rows are ordered correctly; header row unchanged if `has_header`
    - Generator: random numeric/string data, 1-3 sort keys
    - **Validates: Requirements 14.1, 14.2, 14.3**

  - [x] 12.3 Implement `find_replace` in `src/tools/data.rs`
    - Replace all occurrences in sheet or optional range
    - Match against displayed cell values, not formulas
    - Support case-insensitive matching via `match_case` flag (default true)
    - Return `FindReplaceResult` with replacement count; count 0 on no matches
    - _Requirements: 15.1-15.5_

  - [x] 12.4 Write property test for find-replace completeness (Property 11)
    - **Property 11: Find-Replace Completeness**
    - After replace, no cell contains the find string (respecting match_case); returned count equals replacements made
    - **Validates: Requirements 15.1, 15.3**

  - [x] 12.5 Implement `fill_series` in `src/tools/data.rs`
    - Read seed values from source range
    - "linear": detect arithmetic step from numeric seeds, continue progression
    - "date": detect date interval, continue
    - "copy": repeat seed values cyclically
    - Default direction "down", default type "linear"
    - _Requirements: 16.1-16.5_

  - [x] 12.6 Write property test for fill series linear continuation (Property 12)
    - **Property 12: Fill Series Linear Continuation**
    - For 2+ numeric seeds with constant step, filled values continue the arithmetic progression
    - **Validates: Requirements 16.2**

  - [x] 12.7 Write property test for fill series copy cycling (Property 13)
    - **Property 13: Fill Series Copy Cycling**
    - For any seed values and fill count, "copy" type repeats seeds cyclically
    - **Validates: Requirements 16.4**

  - [x] 12.8 Implement `delete_rows_where` in `src/tools/data.rs`
    - Evaluate each row against condition (column, operator, value)
    - Support all operators: equals, not_equals, contains, greater_than, less_than, starts_with, ends_with, is_empty
    - Delete matching rows from bottom to top to preserve indices
    - Skip header row if `has_header` is true
    - Return `DeleteRowsResult` with deleted count
    - _Requirements: 17.1-17.5_

  - [x] 12.9 Write property test for delete rows completeness (Property 9)
    - **Property 9: Delete Rows Completeness**
    - After deletion, no remaining non-header row matches the condition; header unchanged if `has_header`
    - **Validates: Requirements 17.1, 17.5**

  - [x] 12.10 Implement `copy_sheet` in `src/tools/sheets.rs`
    - Create copy of source sheet with all values, formulas, formatting, merged ranges
    - Return new sheet name
    - Error if source not found; error if new name already exists
    - _Requirements: 18.1-18.4_

  - [x] 12.11 Implement `copy_range` in `src/tools/data.rs`
    - Copy values, formulas, and formatting from source range to destination cell
    - Support cross-sheet copy (source_sheet + optional destination_sheet)
    - Preserve source range dimensions in destination
    - _Requirements: 19.1-19.3_

  - [x] 12.12 Implement `transpose_range` in `src/tools/data.rs`
    - Read values from source range, write transposed at destination cell
    - If no destination specified, write back to source range origin
    - Return `TransposeResult` with original and transposed dimensions
    - _Requirements: 20.1-20.3_

  - [x] 12.13 Write property test for transpose involution (Property 8)
    - **Property 8: Transpose Involution**
    - Transposing twice restores original data; transposed dimensions = swapped original dimensions
    - Generator: random 1-10 × 1-10 grids
    - **Validates: Requirements 20.1, 20.3**

  - [x] 12.14 Implement `remove_duplicates` in `src/tools/data.rs`
    - Remove duplicate rows keeping first occurrence
    - Compare specified columns, or all columns if none specified
    - Skip header row if `has_header`
    - Return `RemoveDuplicatesResult` with rows_removed and rows_remaining
    - _Requirements: 21.1-21.4_

  - [x] 12.15 Write property test for remove duplicates uniqueness (Property 10)
    - **Property 10: Remove Duplicates Uniqueness**
    - After dedup, no two non-header rows have identical values in specified columns; first occurrence preserved
    - **Validates: Requirements 21.1, 21.4**

  - [x] 12.16 Implement `split_column` in `src/tools/data.rs`
    - Split each cell in column by delimiter, write parts into consecutive columns to the right
    - Default comma delimiter
    - Skip header row if `has_header`
    - Return `SplitColumnResult` with rows_split and output_columns
    - _Requirements: 22.1-22.4_

  - [x] 12.17 Write property test for split column round-trip (Property 14)
    - **Property 14: Split Column Round-Trip**
    - For any cell value containing delimiter, splitting and re-joining with same delimiter reconstructs original value
    - **Validates: Requirements 22.1**

  - [x] 12.18 Write unit tests for Tier 3 data operations
    - `test_sort_single_key`, `test_sort_multi_key`, `test_sort_with_header`, `test_sort_key_outside_range`
    - `test_find_replace_case_insensitive`, `test_find_replace_no_matches`, `test_find_replace_in_range`
    - `test_fill_series_linear_integers`, `test_fill_series_copy`, `test_fill_series_date`
    - `test_delete_rows_each_operator`, `test_delete_rows_with_header`
    - `test_copy_sheet_basic`, `test_copy_sheet_not_found`, `test_copy_sheet_duplicate_name`
    - `test_copy_range_same_sheet`, `test_copy_range_cross_sheet`
    - `test_transpose_basic`, `test_transpose_in_place`
    - `test_remove_duplicates_all_columns`, `test_remove_duplicates_specific_columns`, `test_remove_duplicates_with_header`
    - `test_split_column_comma`, `test_split_column_custom_delimiter`, `test_split_column_with_header`
    - _Requirements: 14-22_

- [x] 13. Tier 3: Register data operation tools in `src/server.rs`
  - Add `sort_range`, `find_replace`, `fill_series`, `delete_rows_where`, `copy_sheet`, `copy_range`, `transpose_range`, `remove_duplicates`, `split_column` tool registrations
  - Group under `// ── Data operations (9) ──` comment header
  - Add `copy_sheet` under existing sheet management section or data operations section
  - Update server instructions string to mention new tool count and capabilities
  - _Requirements: 14-22_

- [x] 14. Final checkpoint — Ensure all tiers compile and all tests pass
  - Ensure all tests pass, ask the user if questions arise.

## Notes

- Tasks marked with `*` are optional and can be skipped for faster MVP
- Each task references specific requirements for traceability
- Checkpoints ensure incremental validation after each tier
- Property tests validate universal correctness properties from the design document (15 properties)
- Unit tests validate specific examples and edge cases (53 tests)
- The project uses Rust with `proptest` crate (already in dev-dependencies)
- All new tools follow existing patterns: `workbook_id` + `sheet_name`, A1 notation, `#[serde(deny_unknown_fields)]`, `tool_fn!` macro, structured JSON responses

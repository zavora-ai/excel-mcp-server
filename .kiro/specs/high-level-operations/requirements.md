# Requirements Document

## Introduction

This feature adds high-level Excel operation tools to the excel-mcp-server, targeting the actual pain points observed in LLM agent workflows. The tools are organized into three categories: **Styler agent optimizations** (reduce formatting tool calls by 10-80x), **Writer agent optimizations** (reduce data/formula writing tool calls by 5-60x), and **Data operations** (sort, find-replace, dedup, etc.). Each tool composes existing zavora-xlsx primitives into single-call operations. All tools follow existing patterns: `workbook_id` handle, `sheet_name` targeting, A1 notation, `#[serde(deny_unknown_fields)]` input structs, and structured JSON responses.

## Glossary

- **Server**: The excel-mcp-server MCP tool router
- **Engine**: The zavora-xlsx library for low-level Excel read/write
- **Workbook_Store**: In-memory store managing open workbooks with TTL eviction
- **Range**: A rectangular cell region in A1:B2 notation, optionally comma-separated for multi-range ("A1:B5,D1:E5")
- **Cell_Reference**: A single cell address in A1 notation
- **Sort_Key**: A column identifier (letter) paired with a direction ("ascending" or "descending")
- **Condition**: A predicate with column, operator (equals, not_equals, contains, greater_than, less_than, starts_with, ends_with, is_empty), and comparison value
- **Style_Preset**: A named formatting bundle ("header", "currency", "percentage", "date", "title", "number", "text", "accounting", "multiple")
- **Theme**: A named complete sheet styling configuration ("financial_professional", "corporate", "minimal") that applies headers, totals, banding, number formats, and borders
- **Border_Preset**: A named border configuration ("bottom_thick", "box", "top_bottom", "accounting_underline", "none")
- **Semantic_Format**: A human-readable number format hint ("currency", "percentage", "accounting", "multiple", "date", "number", "text") that maps to the correct Excel format code
- **Fill_Direction**: "down" or "right"
- **Fill_Type**: "linear", "date", or "copy"
- **Delimiter**: A string used to split cell text into multiple columns
- **Format_Operation**: A single formatting instruction containing a range and formatting properties, used in batch operations

## Requirements

---

### TIER 1: HIGH PRIORITY — Styler Agent Optimizations

---

### Requirement 1: Comma-Separated Ranges in All Formatting Tools

**User Story:** As an LLM Styler agent, I want to pass comma-separated ranges to all formatting tools, so that I can apply the same format to multiple non-contiguous ranges in one call instead of making separate calls for each range.

#### Acceptance Criteria

1. WHEN a comma-separated range string (e.g., "A1:B5,D1:E5,G1:H5") is provided to `set_cell_format`, THE Server SHALL split on commas and apply the format to each range independently
2. WHEN a comma-separated range string is provided to `merge_cells`, THE Server SHALL merge each range independently
3. WHEN a comma-separated range string is provided to `add_conditional_format`, THE Server SHALL apply the conditional format to each range independently
4. WHEN a comma-separated range string is provided to `set_dimensions`, THE Server SHALL apply the dimension setting to each range independently
5. THE Server SHALL trim whitespace around each range segment before parsing
6. IF any individual range segment is invalid, THE Server SHALL return an error identifying the invalid segment without applying the format to any range

### Requirement 2: Batch Format

**User Story:** As an LLM Styler agent, I want to apply multiple different formatting operations in a single tool call, so that I can format an entire sheet in one call instead of 15-20 separate `set_cell_format` calls.

#### Acceptance Criteria

1. WHEN an array of Format_Operations is provided, THE Server SHALL apply each operation sequentially to the specified ranges
2. EACH Format_Operation SHALL support all properties available in `set_cell_format`: bold, italic, underline, font_size, font_color, background_color, number_format, horizontal_alignment, vertical_alignment, border_style
3. EACH Format_Operation SHALL support comma-separated ranges
4. THE Server SHALL return the count of operations successfully applied
5. IF any operation fails, THE Server SHALL continue processing remaining operations and report failures in the response with the failing range and error message
6. THE Server SHALL support Semantic_Format values in the `number_format` field (e.g., "currency" → "$#,##0.00", "percentage" → "0.0%", "accounting" → "$#,##0.00;($#,##0.00)", "multiple" → "0.0x")
7. THE Server SHALL support Border_Preset values in the `border_style` field (e.g., "bottom_thick", "box", "top_bottom", "accounting_underline")

### Requirement 3: Apply Theme

**User Story:** As an LLM Styler agent, I want to apply a complete professional theme to a sheet in one call, specifying which rows are headers and which are totals, so that I can replace 15-20 formatting calls with a single call.

#### Acceptance Criteria

1. WHEN a Theme name, header row numbers, and total row numbers are provided, THE Server SHALL apply the complete theme styling to the sheet
2. THE Server SHALL support the "financial_professional" Theme which applies: bold white-on-dark-blue headers, bold totals with top border, alternating row shading between header/total sections, and autofit columns
3. THE Server SHALL support the "corporate" Theme which applies: bold dark headers with light gray background, subtle borders, and autofit columns
4. THE Server SHALL support the "minimal" Theme which applies: bold headers, bottom border on headers, bottom border on totals, no background colors
5. WHERE optional `header_rows` are specified, THE Server SHALL apply header styling to those specific rows
6. WHERE optional `total_rows` are specified, THE Server SHALL apply total styling (bold, top border) to those specific rows
7. THE Server SHALL auto-detect currency columns (columns where most values are numeric) ONLY WHEN an optional `auto_detect_formats` field is set to true. This is opt-in because numeric columns like years (2020, 2021) would be incorrectly formatted as currency.
8. IF an unrecognized Theme name is provided, THE Server SHALL return an error listing valid theme names

### Requirement 4: Copy Format

**User Story:** As an LLM Styler agent, I want to copy formatting from one range to multiple target ranges in one call, so that after formatting one header row correctly I can replicate it to other header rows without repeating all the format parameters.

#### Acceptance Criteria

1. WHEN a source range and one or more target ranges are provided, THE Server SHALL read the formatting from the source range and apply it to each target range
2. THE Server SHALL copy: bold, italic, underline, font size, font color, background color, number format, alignment, and border style
3. THE Server SHALL map source formatting cell-by-cell to target ranges (first cell of source maps to first cell of each target)
4. WHERE a target range has different dimensions than the source, THE Server SHALL tile the source formatting to fill the target range
5. IF the source range has no formatting, THE Server SHALL return a success response with a note that no formatting was found to copy

### Requirement 5: Semantic Number Formats in set_cell_format

**User Story:** As an LLM Styler agent, I want to use semantic format hints like "currency" or "percentage" instead of exact Excel format codes, so that I don't have to remember or get wrong the exact syntax like `$#,##0.00` or `0.0%`.

#### Acceptance Criteria

1. WHEN a Semantic_Format value is provided in the `number_format` field of `set_cell_format`, THE Server SHALL resolve it to the corresponding Excel format code before applying
2. THE Server SHALL support the following Semantic_Format mappings:
   - "currency" → "$#,##0.00"
   - "percentage" → "0.0%"
   - "accounting" → '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
   - "multiple" → '0.0"x"'
   - "date" → "yyyy-mm-dd"
   - "number" → "#,##0"
   - "integer" → "#,##0"
   - "text" → "@"
   - "decimal" → "#,##0.00"
3. WHEN a value that is not a recognized Semantic_Format is provided, THE Server SHALL pass it through as a literal Excel format code (backward compatible)
4. THE Server SHALL apply Semantic_Format resolution in `batch_format` operations as well

### Requirement 6: Format as Table Header

**User Story:** As an LLM Styler agent, I want to format a row as a table header (freeze, bold, background color, autofilter) in a single call, so that I can apply the most common formatting pattern without four separate tool calls.

#### Acceptance Criteria

1. WHEN a sheet name is provided, THE Server SHALL apply bold font, white font color (#FFFFFF), dark blue background (#4472C4), and center horizontal alignment to row 1
2. THE Server SHALL freeze panes at cell A2 so row 1 remains visible while scrolling
3. THE Server SHALL set an autofilter on the header row spanning from column A to the last used column
4. WHERE an optional `header_row` number is provided, THE Server SHALL apply the formatting to that row instead of row 1, and freeze panes at the row below it
5. WHERE optional `background_color` or `font_color` are provided, THE Server SHALL use those colors instead of the defaults
6. IF the sheet has no data (empty used range), THE Server SHALL return an error indicating the sheet is empty

### Requirement 7: Format as Table Range

**User Story:** As an LLM Styler agent, I want to apply table-like formatting (banded rows, header styling, borders) to a range without creating a formal Excel Table object, so that I can style data regions that don't need autofilter or structured references.

#### Acceptance Criteria

1. WHEN a range is provided, THE Server SHALL apply header styling to the first row of the range (bold, background color, white font)
2. THE Server SHALL apply alternating row shading to data rows below the header
3. THE Server SHALL apply thin borders to all cells in the range
4. WHERE an optional `style` name is provided (e.g., "blue", "green", "gray", "orange"), THE Server SHALL use the corresponding color scheme
5. WHERE no style is specified, THE Server SHALL default to the blue color scheme (#4472C4 header, #D6E4F0 alternating rows)
6. THE Server SHALL autofit columns within the range

### Requirement 8: Apply Style Preset

**User Story:** As an LLM Styler agent, I want to apply a named style preset to a range in one call, so that I can format cells without specifying verbose JSON formatting objects.

#### Acceptance Criteria

1. WHEN a valid Style_Preset name and range are provided, THE Server SHALL apply the corresponding formatting bundle to all cells in the range
2. THE Server SHALL support comma-separated ranges
3. THE Server SHALL support the following Style_Presets:
   - "header": bold, white font (#FFFFFF), dark blue background (#4472C4), center alignment
   - "title": bold, font size 14, dark font (#1F3864)
   - "currency": number format "$#,##0.00"
   - "percentage": number format "0.0%"
   - "date": number format "yyyy-mm-dd"
   - "number": number format "#,##0"
   - "text": number format "@"
   - "accounting": number format accounting style
   - "total": bold, top border thin
4. IF an unrecognized Style_Preset name is provided, THE Server SHALL return an error listing valid preset names

### Requirement 9: Describe Formatting

**User Story:** As an LLM Styler agent, I want to read the current formatting on a range, so that I can make intelligent decisions about what formatting to add or change.

#### Acceptance Criteria

1. WHEN a range is provided, THE Server SHALL return the formatting properties for each unique format found in the range
2. THE response SHALL include: bold, italic, underline, font_size, font_color, background_color, number_format, horizontal_alignment, vertical_alignment, border_style
3. THE Server SHALL group cells with identical formatting and report the ranges that share each format
4. IF the range has no formatting, THE Server SHALL return an empty format list

---

### TIER 2: HIGH PRIORITY — Writer Agent Optimizations

---

### Requirement 10: Write Grid

**User Story:** As an LLM Writer agent, I want to write a 2D block of data (values and formulas) in one call, so that I can populate an entire financial projection section without 20+ individual write calls.

#### Acceptance Criteria

1. WHEN a start cell and a 2D array of rows are provided, THE Server SHALL write all values starting at the specified cell, filling rightward and downward
2. THE Server SHALL auto-detect value types: strings starting with "=" are formulas, numbers stay numeric, booleans are preserved, ISO dates are parsed
3. THE Server SHALL return the dimensions of the written grid (rows × columns) in the response
4. IF any cell write fails, THE Server SHALL continue writing remaining cells and report failures in the response

### Requirement 11: Write Row Range (Drag-Fill Formulas)

**User Story:** As an LLM Writer agent, I want to write a formula in one cell and have it automatically filled across a row with relative references adjusting, so that I can replicate Excel's drag-fill behavior in one call instead of writing each cell individually.

#### Acceptance Criteria

1. WHEN a start cell, end column, and formula are provided, THE Server SHALL write the formula at the start cell and then fill it rightward to the end column, adjusting relative column references in each cell
2. THE Server SHALL accept the formula with or without a leading "=" (both "=B10*(1+0.05)" and "B10*(1+0.05)" are valid; if "=" is present it is stripped before processing)
3. THE Server SHALL adjust relative references: in formula `=B10*(1+0.05)`, column B increments to C, D, E, etc. as the formula fills right
4. THE Server SHALL preserve absolute references (marked with $): `=$A$1*B10` keeps $A$1 fixed while B10 adjusts
5. THE Server SHALL return the number of cells written
6. IF the start cell column is greater than or equal to the end column, THE Server SHALL return an error

### Requirement 12: Clone Column Formulas

**User Story:** As an LLM Writer agent, I want to copy all formulas from one column to multiple target columns with relative references adjusting, so that after writing Year 1 formulas I can clone them across Years 2-5 in one call instead of 60+ individual write_formula calls.

#### Acceptance Criteria

1. WHEN a source column, target columns array, and row range are provided, THE Server SHALL copy each formula from the source column to each target column, adjusting relative column references
2. THE Server SHALL adjust relative references: a formula `=B10-B12` in column C becomes `=C10-C12` in column D, `=D10-D12` in column E, etc.
3. THE Server SHALL preserve absolute references ($A$1 stays fixed)
4. THE Server SHALL skip non-formula cells (only copy cells that contain formulas)
5. THE Server SHALL return the count of formulas cloned
6. IF the source column has no formulas in the specified row range, THE Server SHALL return a success response with zero formulas cloned

### Requirement 13: Improved write_cells Tool Description and Response

**User Story:** As an LLM Writer agent, I want the `write_cells` tool description to clearly state that formulas are auto-detected, and I want the response to tell me which cells were written as formulas vs values, so that I don't waste calls using `write_formula` and I get feedback on how my input was interpreted.

#### Acceptance Criteria

1. THE Server SHALL update the `write_cells` tool description to explicitly state: "Strings starting with '=' are written as formulas with relative reference support. Use this for batch writes instead of write_formula for individual cells."
2. THE Server SHALL update the `write_formula` tool description to state: "For single formulas needing array or dynamic behavior. For regular formulas, prefer write_cells which auto-detects formulas."
3. THE `write_cells` response SHALL include a `formula_cells` array listing which cells were interpreted as formulas (e.g., ["B6", "C10"])
4. THE `write_cells` response SHALL include a `value_cells` count of cells written as plain values

---

### TIER 3: MEDIUM PRIORITY — Data Operations

---

### Requirement 14: Sort Range

**User Story:** As an LLM agent, I want to sort a range by one or more columns, so that I can organize data in a single call.

#### Acceptance Criteria

1. WHEN a valid range and one or more Sort_Keys are provided, THE Server SHALL sort the rows within the range according to the specified columns and directions
2. WHEN multiple Sort_Keys are provided, THE Server SHALL apply them in priority order (first key is primary sort)
3. WHERE `has_header` is true, THE Server SHALL exclude the first row from sorting
4. THE Server SHALL preserve cell formatting after sorting. NOTE: Formatting stays with cell positions, not with data — the tool description SHALL document this clearly so the LLM knows formatting doesn't move with data.
5. IF a Sort_Key references a column outside the range, THE Server SHALL return an error

### Requirement 15: Find and Replace

**User Story:** As an LLM agent, I want to find and replace values across a sheet or range, so that I can perform data cleanup in one call.

#### Acceptance Criteria

1. WHEN find and replace values are provided, THE Server SHALL replace all occurrences across the sheet (or within an optional range)
2. THE Server SHALL return the replacement count
3. WHERE `match_case` is false, THE Server SHALL match case-insensitively
4. THE Server SHALL match against displayed cell values, not formulas
5. IF no occurrences are found, THE Server SHALL return success with count zero

### Requirement 16: Fill Series

**User Story:** As an LLM agent, I want to auto-fill a pattern down a column or across a row from seed values, so that I can generate sequences without writing hundreds of cells.

#### Acceptance Criteria

1. WHEN a source range and fill count are provided, THE Server SHALL extend the pattern by the specified count in the given Fill_Direction
2. FOR "linear" Fill_Type with numeric seeds, THE Server SHALL detect the arithmetic step and continue
3. FOR "date" Fill_Type with date seeds, THE Server SHALL detect the date interval and continue
4. FOR "copy" Fill_Type, THE Server SHALL repeat seed values cyclically
5. THE Server SHALL default to Fill_Direction "down" and Fill_Type "linear"

### Requirement 17: Delete Rows Where

**User Story:** As an LLM agent, I want to delete rows matching a condition in one call, so that I can clean data without read-filter-delete loops.

#### Acceptance Criteria

1. WHEN a Condition is provided, THE Server SHALL evaluate rows and delete all matching ones
2. THE Server SHALL support operators: equals, not_equals, contains, greater_than, less_than, starts_with, ends_with, is_empty
3. THE Server SHALL delete from bottom to top to preserve indices
4. THE Server SHALL return the deleted row count
5. WHERE `has_header` is true, THE Server SHALL skip the first row

### Requirement 18: Copy Sheet

**User Story:** As an LLM agent, I want to duplicate a sheet with all data and formatting in one call.

#### Acceptance Criteria

1. WHEN source and new sheet names are provided, THE Server SHALL create a copy including values, formulas, formatting, and merged ranges
2. THE Server SHALL return the new sheet name
3. IF the source sheet doesn't exist, THE Server SHALL return a "not_found" error
4. IF the new name already exists, THE Server SHALL return an error indicating the conflict

### Requirement 19: Copy Range

**User Story:** As an LLM agent, I want to copy data and formatting from one range to another in one call.

#### Acceptance Criteria

1. WHEN source range and destination cell are provided, THE Server SHALL copy values, formulas, and formatting
2. THE Server SHALL support cross-sheet copy (source_sheet_name + destination_sheet_name)
3. THE Server SHALL preserve source range dimensions in the destination

### Requirement 20: Transpose Range

**User Story:** As an LLM agent, I want to flip rows and columns for a range in one call.

#### Acceptance Criteria

1. WHEN a source range is provided, THE Server SHALL read values and write them transposed at a destination cell
2. WHERE no destination is specified, THE Server SHALL write back to the source range origin
3. THE Server SHALL return the transposed dimensions

### Requirement 21: Remove Duplicates

**User Story:** As an LLM agent, I want to remove duplicate rows based on specified columns in one call.

#### Acceptance Criteria

1. WHEN a range and column identifiers are provided, THE Server SHALL remove duplicate rows keeping the first occurrence
2. WHERE no columns are specified, THE Server SHALL compare all columns
3. THE Server SHALL return the removed count
4. WHERE `has_header` is true, THE Server SHALL skip the first row

### Requirement 22: Split Column

**User Story:** As an LLM agent, I want to split a column by delimiter into multiple columns in one call.

#### Acceptance Criteria

1. WHEN a source column, row range, and Delimiter are provided, THE Server SHALL split each cell and write parts into consecutive columns to the right
2. THE Server SHALL default to comma delimiter
3. THE Server SHALL return the number of output columns created
4. WHERE `has_header` is true, THE Server SHALL skip the header row

---

### TIER 4: LOWER PRIORITY — Nice to Have

---

### Requirement 23: Auto-Detect Chart Series

**User Story:** As an LLM agent, I want to create a chart from a data range without manually specifying series definitions, so that the engine figures out series from the data layout.

#### Acceptance Criteria

1. WHEN `add_chart` is called with a `data_range` and `first_row_headers: true`, THE Server SHALL auto-detect series from the data layout (each column after the first becomes a series, first column becomes categories)
2. WHERE `first_col_labels` is true, THE Server SHALL use the first column as category labels
3. THE Server SHALL use the header row values as series names

### Requirement 24: Auto Style Sheet

**User Story:** As an LLM Styler agent, I want to say "make this sheet look professional" in one call, so that the server auto-detects headers, number columns, and applies sensible formatting.

#### Acceptance Criteria

1. WHEN called with a sheet name, THE Server SHALL analyze the sheet content and auto-apply formatting
2. THE Server SHALL detect header rows (text in row 1, numbers below) and apply bold
3. THE Server SHALL detect currency/percentage columns and apply appropriate number formats
4. THE Server SHALL autofit all columns
5. THE Server SHALL add borders between sections (where blank rows exist)

## Correctness Properties

### Property 1: Format Idempotency
Applying the same formatting operation twice to the same range SHALL produce the same result as applying it once.

### Property 2: Batch Equivalence
The result of `batch_format` with N operations SHALL be identical to calling `set_cell_format` N times with the same parameters in the same order.

### Property 3: Copy Format Fidelity
After `copy_format`, the formatting on each target range SHALL be identical to the formatting on the source range.

### Property 4: Sort Stability
Sorting a range that is already sorted by the same keys SHALL not change the row order.

### Property 5: Transpose Involution
Transposing a range twice SHALL restore the original data layout.

### Property 6: Delete Rows Correctness
After `delete_rows_where`, no remaining row SHALL match the specified condition.

### Property 7: Remove Duplicates Correctness
After `remove_duplicates`, no two remaining rows SHALL have identical values in the specified columns.

### Property 8: Clone Column Reference Adjustment
After `clone_column_formulas`, each cloned formula's relative column references SHALL be offset by exactly the column distance between source and target.

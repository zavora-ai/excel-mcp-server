# Implementation Plan: Excel MCP Server

## Overview

Build a Rust MCP server that exposes 25+ Excel manipulation tools to LLMs over stdio. The implementation proceeds bottom-up: shared types and utilities first, then the workbook store, engine adapters, tool implementations grouped by capability, and finally wiring everything into the server with the rmcp SDK.

## Tasks

- [x] 1. Initialize project and define core types
  - [x] 1.1 Create Cargo project and configure dependencies
    - Initialize a new Rust binary crate
    - Add dependencies to `Cargo.toml`: `rmcp` (v1.3.0, features: `server`, `transport-io`, `macros`), `rust_xlsxwriter`, `umya-spreadsheet` (v2.3.3), `calamine` (v0.26.1), `tokio` (features: `full`), `serde` (features: `derive`), `serde_json`, `schemars`, `anyhow`, `tracing`, `tracing-subscriber`, `base64`
    - Create the module structure: `src/main.rs`, `src/server.rs`, `src/store.rs`, `src/cell_ref.rs`, `src/error.rs`, `src/engines/mod.rs`, `src/tools/mod.rs`, `src/types/mod.rs` (files can be empty stubs initially)
    - _Requirements: 1.1, 1.2, 1.3_

  - [x] 1.2 Implement error types (`src/error.rs`)
    - Define `ExcelMcpError` enum with variants: `NotFound`, `InvalidInput`, `EngineUnsupported`, `CapacityExceeded`, `IoError`, `ParseError`, `Evicted`
    - Implement `std::fmt::Display` and `std::error::Error` for `ExcelMcpError`
    - Implement `From<std::io::Error>` and other relevant `From` impls
    - _Requirements: 23.3_

  - [x] 1.3 Implement structured response types (`src/types/responses.rs`)
    - Define `ToolResponse<T>`, `Status` enum (`Success`, `Error`), `ErrorData`, `ErrorCategory` enum
    - Implement builder helpers: `success()`, `success_no_data()`, `error()`
    - All builders return `String` (serialized JSON)
    - _Requirements: 23.1, 23.2, 23.3_

  - [x] 1.4 Implement enum types (`src/types/enums.rs`)
    - Define all fixed-value enums: `ChartType`, `SparklineType`, `HorizontalAlignment`, `VerticalAlignment`, `BorderStyle`, `LegendPosition`, `MatchMode`, `ComparisonOperator`, `IconSetStyle`, `AlertStyle`
    - Define tagged enums: `ConditionalFormatRule`, `ValidationRule`
    - Define supporting structs: `ConditionalFormatStyle`, `ValidationMessage`, `ValidationAlert`
    - All enums derive `Deserialize`, `JsonSchema` with `rename_all = "snake_case"`
    - _Requirements: 24.3_

  - [x] 1.5 Implement tool input structs (`src/types/inputs.rs`)
    - Define all input structs with `Deserialize`, `JsonSchema`, `deny_unknown_fields`
    - Add doc comments on every field for schema descriptions
    - Implement serde defaults for optional fields (`read_only`, `delimiter`, `autofilter`, chart dimensions, etc.)
    - Structs: `CreateWorkbookInput`, `OpenWorkbookInput`, `SaveWorkbookInput`, `CloseWorkbookInput`, `ReadSheetInput`, `ReadCellInput`, `CellWrite`, `WriteCellsInput`, `WriteRowInput`, `WriteColumnInput`, `SetCellFormatInput`, `MergeCellsInput`, `AddChartInput`, `AddImageInput`, `AddTableInput`, `AddConditionalFormatInput`, `AddDataValidationInput`, `SetColumnWidthInput`, `SetRowHeightInput`, `FreezePanesInput`, `AddSparklineInput`, `SearchCellsInput`, `SheetToCsvInput`, `AddSheetInput`, `RenameSheetInput`, `DeleteSheetInput`, `ListSheetsInput`, `GetSheetDimensionsInput`, `DescribeWorkbookInput`
    - _Requirements: 24.1, 24.2, 24.4_

  - [x] 1.6 Implement response data structs (`src/types/responses.rs` or `src/types/mod.rs`)
    - Define `WorkbookInfo`, `SheetSummary`, `ReadSheetData`, `CellData`, `WriteResult`, `SearchResult`, `SearchMatch`, `WorkbookDescription`, `SheetDescription`, `CsvExportData`
    - Define `ContinuationToken` struct for pagination (with `Serialize` + `Deserialize`)
    - _Requirements: 23.2, 5.2_

- [x] 2. Implement cell reference parser and workbook store
  - [x] 2.1 Implement A1 notation parser (`src/cell_ref.rs`)
    - Define `CellPos { row: u32, col: u16 }` and `CellRange { start, end }`
    - Implement `parse_cell_ref()`: split letters/digits, base-26 column conversion, 1-based to 0-based row
    - Implement `parse_range_ref()`: split on ":", parse both cell refs
    - Implement `col_letter_to_index()`, `index_to_col_letter()`, `cell_pos_to_a1()`
    - Validate against Excel limits: max column XFD (16383), max row 1048576
    - _Requirements: 7.5, 8.3_

  - [x]* 2.2 Write property tests for cell reference parser
    - **Property 1: Round-trip consistency — `parse_cell_ref(cell_pos_to_a1(pos)) == pos` for all valid positions**
    - **Validates: Requirements 7.5, 8.3**

  - [x]* 2.3 Write unit tests for cell reference parser
    - Test known conversions: "A1" → (0,0), "Z1" → (0,25), "AA1" → (0,26), "XFD1" → (0,16383)
    - Test invalid inputs: empty string, digits only, exceeding max column/row
    - Test range parsing: "A1:B2", "A1:XFD1048576"
    - _Requirements: 7.5, 8.3_

  - [x] 2.4 Implement WorkbookStore (`src/store.rs`)
    - Define `EngineTag` enum: `RustXlsxWriter`, `UmyaSpreadsheet`, `Calamine`
    - Define `WorkbookData` enum with one variant per engine
    - Define `CalamineWorkbook` wrapper struct
    - Define `WorkbookEntry` with `id`, `tag`, `data`, `last_access`, `sheet_index_map`
    - Implement `WorkbookStore` with `HashMap<String, WorkbookEntry>`, `max_capacity` (default 10), `ttl` (default 30 min)
    - Implement methods: `insert()` (generate UUID, check capacity), `get()`, `get_mut()`, `remove()`, `evict_expired()`, `open_ids()`, `is_full()`
    - Call `evict_expired()` lazily on every `get`/`get_mut`/`insert`
    - Update `last_access` on every access
    - _Requirements: 21.1, 21.2, 21.3, 21.4, 21.5_

  - [x]* 2.5 Write unit tests for WorkbookStore
    - Test capacity enforcement: inserting beyond max returns error
    - Test TTL eviction: entries older than TTL are removed
    - Test `remove()` frees entry
    - Test `open_ids()` returns all active IDs
    - _Requirements: 21.1, 21.3, 21.5_

- [x] 3. Checkpoint
  - Ensure all tests pass, ask the user if questions arise.

- [x] 4. Implement engine adapters
  - [x] 4.1 Implement rust_xlsxwriter adapter (`src/engines/rxw.rs`)
    - Implement `create_workbook()` → returns `(Workbook, HashMap<String, usize>)` with default "Sheet1"
    - Implement `write_cells()` with type detection: strings, numbers, booleans, formulas (prefix "="), dates (ISO 8601)
    - Implement `write_row()` and `write_column()` convenience functions
    - Implement `add_sheet()`, `rename_sheet()`, `delete_sheet()` with sheet_index_map maintenance
    - Implement `set_cell_format()` with all formatting options (bold, italic, underline, font size, colors, number format, alignment, borders)
    - Implement `merge_cells()`
    - Implement `add_chart()` for all 7 chart types with optional config
    - Implement `add_image()` with optional scaling
    - Implement `add_table()` with style, totals row, autofilter
    - Implement `add_conditional_format()` for all rule types
    - Implement `add_data_validation()` for all validation types
    - Implement `set_column_width()`, `set_row_height()`, `freeze_panes()`
    - Implement `add_sparkline()` for line, column, win_loss
    - Implement `save()` to write xlsx to disk
    - _Requirements: 2.1, 2.3, 7.1, 7.2, 7.3, 7.4, 8.1, 8.2, 9.1, 9.2, 10.1, 11.1, 11.2, 11.3, 12.1, 12.2, 12.3, 13.1, 13.2, 14.1, 14.2, 15.1, 15.2, 15.3, 16.1, 16.2, 16.3, 17.1, 17.2, 17.3, 18.1, 18.2, 4.1_

  - [x] 4.2 Implement umya-spreadsheet adapter (`src/engines/umya.rs`)
    - Implement `open_workbook()` with lazy_read support
    - Implement `read_sheet()` with optional range and pagination (100-row pages, continuation tokens)
    - Implement `read_cell()` returning value, type, and formula
    - Implement `write_cells()`, `write_row()`, `write_column()` with type detection
    - Implement `add_sheet()`, `rename_sheet()`, `delete_sheet()`
    - Implement `set_cell_format()` with supported formatting options
    - Implement `merge_cells()`
    - Implement `add_chart()` for supported chart types
    - Implement `add_image()` with optional scaling
    - Implement `freeze_panes()`, `set_column_width()`, `set_row_height()`
    - Implement `search_cells()` with exact and substring modes, 50-match cap
    - Implement `sheet_to_csv()` with configurable delimiter and 500-row truncation
    - Implement `save()` to write back to disk
    - Implement `list_sheets()`, `get_sheet_dimensions()`, `describe_workbook()` with sample rows
    - _Requirements: 3.1, 3.5, 5.1, 5.2, 5.3, 5.4, 6.1, 6.2, 6.3, 7.1, 7.2, 7.3, 7.4, 8.1, 8.2, 9.1, 9.2, 10.1, 11.1, 11.2, 11.3, 12.1, 17.1, 17.2, 17.3, 19.1, 19.2, 19.3, 19.4, 20.1, 20.2, 20.3, 4.1_

  - [x] 4.3 Implement calamine adapter (`src/engines/calamine.rs`)
    - Implement `open_workbook()` returning `CalamineWorkbook` with cached sheet names
    - Implement `read_sheet()` with optional range and pagination (100-row pages, continuation tokens)
    - Implement `read_cell()` returning value and type
    - Implement `list_sheets()`, `get_sheet_dimensions()`, `describe_workbook()` with sample rows
    - Implement `search_cells()` with exact and substring modes, 50-match cap
    - Implement `sheet_to_csv()` with configurable delimiter and 500-row truncation
    - _Requirements: 3.2, 5.1, 5.2, 5.3, 5.4, 6.1, 6.2, 6.3, 19.1, 19.2, 19.3, 19.4, 20.1, 20.2, 20.3_

  - [x] 4.4 Implement engine module re-exports (`src/engines/mod.rs`)
    - Re-export all three adapter modules
    - _Requirements: 22.1, 22.2_

- [x] 5. Checkpoint
  - Ensure all tests pass, ask the user if questions arise.

- [x] 6. Implement MCP tool methods
  - [x] 6.1 Implement workbook lifecycle tools (`src/tools/workbook.rs`)
    - `create_workbook`: create via rxw adapter, insert into store, return WorkbookInfo
    - `open_workbook`: route to umya (edit) or calamine (read-only) based on `read_only` flag, validate file exists and format, insert into store
    - `save_workbook`: lookup workbook, check engine supports save, delegate to adapter, handle calamine error case
    - `close_workbook`: remove from store, return confirmation
    - All methods: acquire store lock, check capacity, update timestamps, return structured JSON
    - _Requirements: 2.1, 2.2, 2.3, 2.4, 3.1, 3.2, 3.3, 3.4, 3.5, 4.1, 4.2, 4.3, 4.4, 21.5_

  - [x] 6.2 Implement sheet tools (`src/tools/sheets.rs`)
    - `list_sheets`: delegate to engine adapter based on tag
    - `get_sheet_dimensions`: delegate, handle rxw write-only case
    - `describe_workbook`: delegate, handle rxw write-only case (return names only)
    - `add_sheet`, `rename_sheet`, `delete_sheet`: check engine supports operation, validate name uniqueness, validate min-one-sheet constraint
    - _Requirements: 6.1, 6.2, 6.3, 6.4, 11.1, 11.2, 11.3, 11.4, 11.5_

  - [x] 6.3 Implement read tools (`src/tools/read.rs`)
    - `read_sheet`: validate sheet exists, check engine supports read (reject rxw), delegate with pagination
    - `read_cell`: validate cell ref, check engine supports read, delegate
    - `search_cells`: check engine supports search (reject rxw), delegate
    - `sheet_to_csv`: check engine supports csv export (reject rxw), delegate
    - _Requirements: 5.1, 5.2, 5.3, 5.4, 5.5, 19.1, 19.2, 19.3, 19.4, 20.1, 20.2, 20.3_

  - [x] 6.4 Implement write tools (`src/tools/write.rs`)
    - `write_cells`: validate all cell refs, check engine supports write (reject calamine), delegate, return WriteResult
    - `write_row`: validate start cell, check Excel column limit, delegate
    - `write_column`: validate start cell, check Excel row limit, delegate
    - _Requirements: 7.1, 7.2, 7.3, 7.4, 7.5, 8.1, 8.2, 8.3, 22.3_

  - [x] 6.5 Implement formatting tools (`src/tools/format.rs`)
    - `set_cell_format`: validate range, validate hex colors, check engine supports formatting, delegate
    - `merge_cells`: validate range, check engine supports merge, delegate
    - _Requirements: 9.1, 9.2, 9.3, 9.4, 10.1, 10.2_

  - [x] 6.6 Implement chart, image, table tools (`src/tools/charts.rs`, `src/tools/images.rs`, `src/tools/tables.rs`)
    - `add_chart`: validate range, check engine supports charts, delegate
    - `add_image`: validate image file exists and format, check engine supports images, delegate
    - `add_table`: validate range, check engine supports tables (rxw only), delegate
    - _Requirements: 12.1, 12.2, 12.3, 12.4, 13.1, 13.2, 13.3, 14.1, 14.2, 14.3_

  - [x] 6.7 Implement conditional formatting, data validation, sparklines (`src/tools/conditional.rs`, `src/tools/validation.rs`, `src/tools/sparklines.rs`)
    - `add_conditional_format`: check engine supports (rxw only), delegate
    - `add_data_validation`: check engine supports (rxw only), delegate
    - `add_sparkline`: check engine supports (rxw only), reject umya with informative error, delegate
    - _Requirements: 15.1, 15.2, 15.3, 16.1, 16.2, 16.3, 18.1, 18.2, 18.3_

  - [x] 6.8 Implement layout tools (`src/tools/layout.rs`)
    - `set_column_width`: validate column identifier, check engine supports, delegate
    - `set_row_height`: validate row number, check engine supports, delegate
    - `freeze_panes`: validate cell ref, check engine supports, delegate
    - _Requirements: 17.1, 17.2, 17.3, 17.4_

  - [x] 6.9 Implement tool module re-exports (`src/tools/mod.rs`)
    - Re-export all tool submodules
    - _Requirements: 1.3_

- [x] 7. Checkpoint
  - Ensure all tests pass, ask the user if questions arise.

- [x] 8. Wire server and main entry point
  - [x] 8.1 Implement ExcelMcpServer (`src/server.rs`)
    - Define `ExcelMcpServer` struct holding `Arc<RwLock<WorkbookStore>>`
    - Implement `ServerHandler` trait with `name()` returning "excel-mcp" and `instructions()` describing tool categories
    - Annotate with `#[tool(tool_box)]`
    - Implement all tool methods with `#[tool(description = "...")]` and `#[tool(aggr)]` for `Json<T>` input
    - Each method: deserialize input, validate, acquire store lock, check engine tag, delegate to tool module, update timestamp, return JSON string
    - _Requirements: 1.3, 22.1, 22.2, 22.3, 22.4, 23.1, 23.4, 23.5_

  - [x] 8.2 Implement main entry point (`src/main.rs`)
    - Initialize `tracing_subscriber` directing output to stderr
    - Create `WorkbookStore` with default config
    - Wrap in `Arc<RwLock<...>>`
    - Create `ExcelMcpServer` instance
    - Build and run stdio transport via `rmcp`
    - Handle initialization errors: log to stderr, exit with non-zero code
    - _Requirements: 1.1, 1.2, 1.4_

- [x] 9. Checkpoint
  - Ensure all tests pass, ask the user if questions arise.

- [x]* 10. Write integration tests for engine routing and error handling
  - [x]* 10.1 Write tests for engine capability routing
    - Test that write operations on calamine workbooks return `EngineUnsupported` error with suggestion
    - Test that read operations on rxw workbooks return informative error
    - Test that sparkline on umya workbook returns error suggesting rxw
    - Test that table/conditional-format/validation on umya workbook returns error
    - _Requirements: 22.1, 22.2, 22.3, 22.4_

  - [x]* 10.2 Write tests for structured error responses
    - Test that all error responses contain `status: "error"`, `category`, `description`, `suggestion`
    - Test not-found workbook error includes open IDs
    - Test invalid cell reference error identifies the bad reference
    - Test capacity exceeded error message
    - _Requirements: 23.1, 23.2, 23.3, 23.4, 23.5_

  - [x]* 10.3 Write tests for workbook lifecycle
    - Test create → write → save → close flow
    - Test open (edit) → read → write → save flow
    - Test open (read-only) → read → search → csv export flow
    - Test eviction after TTL expiry
    - _Requirements: 2.1, 3.1, 3.2, 4.1, 21.3, 21.4_

- [x] 11. Final checkpoint
  - Ensure all tests pass, ask the user if questions arise.

## Notes

- Tasks marked with `*` are optional and can be skipped for faster MVP
- Each task references specific requirements for traceability
- Checkpoints ensure incremental validation
- The rust_xlsxwriter engine is write-only — read operations on rxw workbooks must return informative errors
- Engine capability checks must happen before delegation to prevent runtime panics
- All tool methods return `Result<String, anyhow::Error>` where the String is always serialized `ToolResponse` JSON

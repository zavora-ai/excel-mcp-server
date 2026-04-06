# Requirements Document

## Introduction

This document specifies the requirements for an MCP (Model Context Protocol) server implemented in Rust that enables LLMs to create, read, edit, and analyze Excel files. The server uses a hybrid engine architecture: `rust_xlsxwriter` for creating new files, `umya-spreadsheet` for editing existing files, and `calamine` for fast read-only analysis. Workbooks are held in memory with handle-based access, and the server communicates over stdio using the `rmcp` SDK.

## Glossary

- **MCP_Server**: The Excel MCP server process that exposes tools over the Model Context Protocol via stdio transport
- **Workbook_Store**: The in-memory `HashMap` protected by `tokio::sync::RwLock` that maps workbook IDs to their engine-specific workbook representations
- **Workbook_ID**: A unique string identifier returned when a workbook is created or opened, used to reference that workbook in all subsequent tool calls
- **Engine_Tag**: A discriminator stored alongside each workbook in the Workbook_Store indicating which engine owns it (`rust_xlsxwriter` for new files, `umya-spreadsheet` for edited existing files, `calamine` for read-only files)
- **Cell_Reference**: A string in A1 notation (e.g., "B3") or row-column pair identifying a single cell
- **Range_Reference**: A string in A1:B2 notation identifying a rectangular region of cells
- **Tool**: An MCP tool exposed by the server, callable by an LLM client, annotated with `#[tool(description = "...")]`
- **LLM_Client**: The language model or agent that connects to the MCP_Server and invokes tools
- **Sheet_Name**: A string identifying a worksheet within a workbook
- **TTL**: Time-to-live duration after which an idle workbook is evicted from the Workbook_Store

## Requirements

### Requirement 1: Server Initialization and Transport

**User Story:** As an LLM_Client, I want to connect to the MCP_Server over stdio, so that I can invoke Excel tools through the standard MCP protocol.

#### Acceptance Criteria

1. WHEN the MCP_Server process starts, THE MCP_Server SHALL initialize the stdio transport and begin accepting MCP protocol messages on stdin/stdout
2. THE MCP_Server SHALL direct all logging and diagnostic output to stderr so that stdout remains exclusively for MCP protocol messages
3. THE MCP_Server SHALL report its name as "excel-mcp" and provide instructions describing available tool categories to the LLM_Client
4. IF the MCP_Server fails to initialize the transport, THEN THE MCP_Server SHALL log the error to stderr and exit with a non-zero exit code

### Requirement 2: Create New Workbook

**User Story:** As an LLM_Client, I want to create a new empty workbook in memory, so that I can build an Excel file from scratch using the rich `rust_xlsxwriter` engine.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `create_workbook` tool, THE MCP_Server SHALL create a new in-memory workbook using the `rust_xlsxwriter` engine and return a unique Workbook_ID
2. WHEN the LLM_Client invokes the `create_workbook` tool, THE MCP_Server SHALL tag the workbook with the `rust_xlsxwriter` Engine_Tag in the Workbook_Store
3. THE MCP_Server SHALL include a default worksheet named "Sheet1" in every newly created workbook
4. IF the Workbook_Store has reached its maximum capacity, THEN THE MCP_Server SHALL return an error message stating the limit and suggesting the LLM_Client save and close an existing workbook

### Requirement 3: Open Existing Workbook

**User Story:** As an LLM_Client, I want to open an existing Excel file for reading or editing, so that I can analyze or modify its contents.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `open_workbook` tool with a file path and `read_only` set to false, THE MCP_Server SHALL load the file using the `umya-spreadsheet` engine and return a Workbook_ID along with sheet names and dimensions
2. WHEN the LLM_Client invokes the `open_workbook` tool with a file path and `read_only` set to true, THE MCP_Server SHALL load the file using the `calamine` engine and return a Workbook_ID along with sheet names and dimensions
3. IF the specified file does not exist, THEN THE MCP_Server SHALL return an error message including the file path that was not found
4. IF the specified file is not a valid xlsx, xlsm, xls, or ods file, THEN THE MCP_Server SHALL return an error message stating the supported formats
5. WHEN opening a large file with `read_only` set to false, THE MCP_Server SHALL use `umya-spreadsheet`'s lazy_read mode to minimize memory usage

### Requirement 4: Save Workbook to Disk

**User Story:** As an LLM_Client, I want to save a workbook to disk, so that the Excel file is persisted and available for download or further use.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `save_workbook` tool with a Workbook_ID and file path, THE MCP_Server SHALL write the workbook to the specified path in xlsx format
2. IF the Workbook_ID does not exist in the Workbook_Store, THEN THE MCP_Server SHALL return an error message stating the workbook was not found and listing currently open Workbook_IDs
3. IF the file path is not writable, THEN THE MCP_Server SHALL return an error message including the path and the OS-level error description
4. WHEN saving a workbook opened with the `calamine` read-only engine, THE MCP_Server SHALL return an error message stating that read-only workbooks cannot be saved and suggesting the LLM_Client reopen the file in edit mode


### Requirement 5: Read Sheet Data

**User Story:** As an LLM_Client, I want to read data from a worksheet, so that I can analyze the contents of an Excel file.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `read_sheet` tool with a Workbook_ID, Sheet_Name, and optional Range_Reference, THE MCP_Server SHALL return the cell values as a structured JSON array of rows
2. WHEN the result set exceeds 100 rows, THE MCP_Server SHALL paginate the response by returning the first 100 rows and a continuation token for fetching subsequent pages
3. WHEN the LLM_Client invokes the `read_cell` tool with a Workbook_ID, Sheet_Name, and Cell_Reference, THE MCP_Server SHALL return the cell value, type, and formula (if present)
4. IF the specified Sheet_Name does not exist in the workbook, THEN THE MCP_Server SHALL return an error message listing the available sheet names
5. IF the LLM_Client invokes `read_sheet` or `read_cell` on a workbook created via `create_workbook` (rust_xlsxwriter engine), THEN THE MCP_Server SHALL return an error message stating that read operations are not available for newly created workbooks and suggesting the LLM_Client track its own writes or save and reopen the file in edit mode

### Requirement 6: List and Describe Sheets

**User Story:** As an LLM_Client, I want to list sheets and get dimensions, so that I can understand the structure of a workbook before operating on it.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `list_sheets` tool with a Workbook_ID, THE MCP_Server SHALL return an ordered list of sheet names
2. WHEN the LLM_Client invokes the `get_sheet_dimensions` tool with a Workbook_ID and Sheet_Name, THE MCP_Server SHALL return the used range as a Range_Reference, the row count, and the column count
3. WHEN the LLM_Client invokes the `describe_workbook` tool with a Workbook_ID, THE MCP_Server SHALL return a summary including sheet names, dimensions per sheet, and a sample of the first 5 rows of each sheet
4. WHEN the LLM_Client invokes `describe_workbook` or `get_sheet_dimensions` on a workbook created via `create_workbook` (rust_xlsxwriter engine), THE MCP_Server SHALL return sheet names and report that cell data and dimensions are not available for write-only workbooks

### Requirement 7: Write Cell Values in Batch

**User Story:** As an LLM_Client, I want to write multiple cell values in a single call, so that I can efficiently populate a spreadsheet without repeated round-trips.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `write_cells` tool with a Workbook_ID, Sheet_Name, and an array of cell writes (each specifying Cell_Reference and value), THE MCP_Server SHALL write all values to the specified cells
2. THE MCP_Server SHALL support writing values of type string, number, boolean, date (ISO 8601 format), and formula (prefixed with "=")
3. WHEN a value is prefixed with "=", THE MCP_Server SHALL write it as an Excel formula rather than a string literal
4. THE MCP_Server SHALL return a confirmation message stating the number of cells written and the range covered
5. IF any Cell_Reference is invalid, THEN THE MCP_Server SHALL return an error message identifying the invalid reference and its position in the input array

### Requirement 8: Write Row and Column Convenience Tools

**User Story:** As an LLM_Client, I want convenience tools for writing a row or column of data, so that I can quickly populate sequential data without specifying every cell reference.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `write_row` tool with a Workbook_ID, Sheet_Name, starting Cell_Reference, and an array of values, THE MCP_Server SHALL write the values to consecutive cells in the same row starting from the specified cell
2. WHEN the LLM_Client invokes the `write_column` tool with a Workbook_ID, Sheet_Name, starting Cell_Reference, and an array of values, THE MCP_Server SHALL write the values to consecutive cells in the same column starting from the specified cell
3. IF writing would exceed the maximum column index of 16383 or maximum row index of 1048575, THEN THE MCP_Server SHALL return an error message stating the Excel limit that would be exceeded

### Requirement 9: Cell Formatting

**User Story:** As an LLM_Client, I want to apply formatting to cells, so that the generated Excel files look professional and are easy to read.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `set_cell_format` tool with a Workbook_ID, Sheet_Name, Range_Reference, and formatting options, THE MCP_Server SHALL apply the specified formatting to all cells in the range
2. THE MCP_Server SHALL support formatting options including: bold, italic, underline, font size, font color (hex), background color (hex), number format string, horizontal alignment, vertical alignment, and border style
3. WHEN a formatting option is omitted, THE MCP_Server SHALL leave the existing formatting for that property unchanged
4. IF a hex color value is invalid, THEN THE MCP_Server SHALL return an error message showing the invalid value and the expected format "#RRGGBB"

### Requirement 10: Merge Cells

**User Story:** As an LLM_Client, I want to merge cells, so that I can create headers and labels that span multiple columns or rows.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `merge_cells` tool with a Workbook_ID, Sheet_Name, and Range_Reference, THE MCP_Server SHALL merge all cells in the specified range
2. IF the specified range overlaps with an already-merged range, THEN THE MCP_Server SHALL return an error message identifying the conflicting merged range

### Requirement 11: Sheet Management

**User Story:** As an LLM_Client, I want to add, rename, and delete worksheets, so that I can organize the workbook structure.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `add_sheet` tool with a Workbook_ID and a Sheet_Name, THE MCP_Server SHALL add a new empty worksheet with the specified name
2. WHEN the LLM_Client invokes the `rename_sheet` tool with a Workbook_ID, current Sheet_Name, and new Sheet_Name, THE MCP_Server SHALL rename the worksheet
3. WHEN the LLM_Client invokes the `delete_sheet` tool with a Workbook_ID and Sheet_Name, THE MCP_Server SHALL remove the worksheet from the workbook
4. IF a sheet with the specified name already exists when adding or renaming, THEN THE MCP_Server SHALL return an error message stating the name is already in use
5. IF the workbook contains only one sheet and the LLM_Client attempts to delete it, THEN THE MCP_Server SHALL return an error message stating that a workbook must contain at least one sheet


### Requirement 12: Charts

**User Story:** As an LLM_Client, I want to add charts to worksheets, so that I can create visual representations of data.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `add_chart` tool with a Workbook_ID, Sheet_Name, chart type, data Range_Reference, and optional configuration, THE MCP_Server SHALL create a chart embedded in the specified sheet
2. THE MCP_Server SHALL support chart types: bar, column, line, pie, scatter, area, and doughnut
3. THE MCP_Server SHALL accept optional chart configuration including: title, x-axis label, y-axis label, legend position, and chart dimensions
4. IF the data Range_Reference is empty or contains no numeric data, THEN THE MCP_Server SHALL return an error message stating that the chart data range must contain numeric values

### Requirement 13: Images

**User Story:** As an LLM_Client, I want to embed images into worksheets, so that I can include logos, diagrams, or visual content in the Excel file.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `add_image` tool with a Workbook_ID, Sheet_Name, Cell_Reference for position, and a file path to a PNG or JPEG image, THE MCP_Server SHALL embed the image at the specified cell position
2. THE MCP_Server SHALL accept optional width and height parameters to scale the image
3. IF the image file does not exist or is not a valid PNG or JPEG, THEN THE MCP_Server SHALL return an error message stating the supported image formats and the path that failed

### Requirement 14: Excel Tables

**User Story:** As an LLM_Client, I want to create structured Excel Tables, so that the data benefits from autofilter, banded rows, and structured references.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `add_table` tool with a Workbook_ID, Sheet_Name, Range_Reference, and column headers, THE MCP_Server SHALL create an Excel Table spanning the specified range
2. THE MCP_Server SHALL accept optional configuration including: table style name, totals row with aggregate functions per column, and autofilter toggle
3. IF the Range_Reference overlaps with an existing table, THEN THE MCP_Server SHALL return an error message identifying the conflicting table

### Requirement 15: Conditional Formatting

**User Story:** As an LLM_Client, I want to apply conditional formatting rules, so that cells are visually highlighted based on their values.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `add_conditional_format` tool with a Workbook_ID, Sheet_Name, Range_Reference, and a rule definition, THE MCP_Server SHALL apply the conditional formatting rule to the specified range
2. THE MCP_Server SHALL support rule types: cell value comparisons (greater than, less than, between, equal to), color scales (2-color and 3-color), data bars, and icon sets
3. THE MCP_Server SHALL accept formatting to apply when the condition is met, including font color, background color, and bold

### Requirement 16: Data Validation

**User Story:** As an LLM_Client, I want to add data validation rules to cells, so that users of the generated Excel file are guided to enter correct data.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `add_data_validation` tool with a Workbook_ID, Sheet_Name, Range_Reference, and validation rule, THE MCP_Server SHALL apply the data validation to the specified range
2. THE MCP_Server SHALL support validation types: dropdown list (from explicit values or a cell range), whole number range, decimal range, date range, text length range, and custom formula
3. THE MCP_Server SHALL accept optional input message (title and body) and error alert (style, title, and message) for the validation

### Requirement 17: Layout Controls

**User Story:** As an LLM_Client, I want to control column widths, row heights, and freeze panes, so that the spreadsheet layout is optimized for readability.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `set_column_width` tool with a Workbook_ID, Sheet_Name, column identifier, and width value, THE MCP_Server SHALL set the column width in character units
2. WHEN the LLM_Client invokes the `set_row_height` tool with a Workbook_ID, Sheet_Name, row number, and height value, THE MCP_Server SHALL set the row height in points
3. WHEN the LLM_Client invokes the `freeze_panes` tool with a Workbook_ID, Sheet_Name, and Cell_Reference, THE MCP_Server SHALL freeze all rows above and all columns to the left of the specified cell
4. IF the column identifier or row number is out of the valid Excel range, THEN THE MCP_Server SHALL return an error message stating the valid range

### Requirement 18: Sparklines

**User Story:** As an LLM_Client, I want to add sparklines to cells, so that I can show inline mini-charts alongside data.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `add_sparkline` tool with a Workbook_ID, Sheet_Name, target Cell_Reference, data Range_Reference, and sparkline type, THE MCP_Server SHALL create a sparkline in the target cell
2. THE MCP_Server SHALL support sparkline types: line, column, and win_loss
3. WHEN the workbook was opened with the `umya-spreadsheet` engine, THE MCP_Server SHALL return an error message stating that sparklines are only supported for newly created workbooks using the `rust_xlsxwriter` engine


### Requirement 19: Search Cells

**User Story:** As an LLM_Client, I want to search for cells matching a pattern or value, so that I can locate specific data within a workbook.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `search_cells` tool with a Workbook_ID, optional Sheet_Name, and a search query, THE MCP_Server SHALL return a list of matching cells with their sheet name, Cell_Reference, and value
2. THE MCP_Server SHALL support exact match and substring match modes
3. WHEN the search produces more than 50 matches, THE MCP_Server SHALL return the first 50 matches and a total count of all matches
4. WHEN no Sheet_Name is provided, THE MCP_Server SHALL search across all sheets in the workbook

### Requirement 20: CSV Export

**User Story:** As an LLM_Client, I want to export a sheet as a CSV string, so that I can use the data in other tools or formats.

#### Acceptance Criteria

1. WHEN the LLM_Client invokes the `sheet_to_csv` tool with a Workbook_ID and Sheet_Name, THE MCP_Server SHALL return the sheet data as a CSV-formatted string
2. THE MCP_Server SHALL use comma as the default delimiter and support an optional custom delimiter parameter
3. WHEN the sheet contains more than 500 rows, THE MCP_Server SHALL truncate the output at 500 rows and include a message stating the total row count and that the output was truncated

### Requirement 21: Memory Management and Workbook Lifecycle

**User Story:** As a system administrator, I want the MCP_Server to manage memory responsibly, so that long-running sessions do not exhaust system resources.

#### Acceptance Criteria

1. THE MCP_Server SHALL enforce a configurable maximum number of concurrent open workbooks (default: 10)
2. THE MCP_Server SHALL track the last-access timestamp for each workbook in the Workbook_Store
3. WHEN a workbook has not been accessed for longer than the configured TTL (default: 30 minutes), THE MCP_Server SHALL evict the workbook from the Workbook_Store and log a warning to stderr
4. WHEN the LLM_Client invokes any tool with a Workbook_ID that has been evicted, THE MCP_Server SHALL return an error message stating the workbook was evicted due to inactivity and suggesting the LLM_Client reopen the file
5. WHEN the LLM_Client invokes a `close_workbook` tool with a Workbook_ID, THE MCP_Server SHALL remove the workbook from the Workbook_Store and free associated memory

### Requirement 22: Engine Routing and Capability Boundaries

**User Story:** As an LLM_Client, I want clear feedback when a requested operation is not supported by the engine managing a workbook, so that I can choose the correct workflow.

#### Acceptance Criteria

1. THE MCP_Server SHALL route write operations to the `rust_xlsxwriter` engine for workbooks created via `create_workbook` and to the `umya-spreadsheet` engine for workbooks opened via `open_workbook` in edit mode
2. THE MCP_Server SHALL route read operations to the `calamine` engine for workbooks opened in read-only mode
3. IF the LLM_Client invokes a write operation on a read-only workbook, THEN THE MCP_Server SHALL return an error message stating the workbook is read-only and suggesting the LLM_Client reopen the file in edit mode
4. IF the LLM_Client invokes an operation not supported by the current engine for a workbook, THEN THE MCP_Server SHALL return an error message naming the unsupported operation and the engine, and suggest an alternative workflow

### Requirement 23: Error Handling and Structured Responses

**User Story:** As an LLM_Client, I want structured, descriptive error messages and responses, so that I can reason about failures and take corrective action.

#### Acceptance Criteria

1. THE MCP_Server SHALL return all tool responses as structured JSON strings containing a "status" field ("success" or "error") and a "message" field
2. WHEN a tool succeeds, THE MCP_Server SHALL include relevant result data in the response alongside the status and message
3. WHEN a tool fails, THE MCP_Server SHALL include the error category, a human-readable description, and a suggestion for resolution in the response
4. THE MCP_Server SHALL validate all input parameters against their JSON schemas before executing tool logic
5. IF a required parameter is missing or has an invalid type, THEN THE MCP_Server SHALL return an error message identifying the parameter name, the expected type, and the value received

### Requirement 24: Tool Parameter Schemas

**User Story:** As an LLM_Client, I want rich JSON schemas for every tool parameter, so that I can construct valid tool calls without guessing.

#### Acceptance Criteria

1. THE MCP_Server SHALL derive JSON schemas for all tool input parameters using `schemars::JsonSchema`
2. THE MCP_Server SHALL include a `description` annotation on every field of every tool input struct
3. THE MCP_Server SHALL use enum types with documented variants for parameters that accept a fixed set of values (e.g., chart types, alignment options, validation types)
4. THE MCP_Server SHALL provide default values for optional parameters using serde defaults and document those defaults in the schema descriptions

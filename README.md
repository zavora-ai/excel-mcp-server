# Excel MCP Server

A high-performance [Model Context Protocol](https://modelcontextprotocol.io/) (MCP) server that gives AI assistants full control over Excel spreadsheets. Built in Rust with [zavora-xlsx](https://github.com/zavora-ai/zavora-xlsx) for native xlsx read/write and [rmcp](https://github.com/modelcontextprotocol/rust-sdk) for the MCP protocol layer.

**43 tools** covering the complete Excel feature set — from basic cell writes to pivot tables, charts, conditional formatting, shapes, and document properties.

## Features

- **Create, open, edit, and save** xlsx files through natural language
- **Read and analyze** existing spreadsheets with pagination and search
- **Full formatting** — bold, italic, colors, borders, number formats, alignment
- **10 chart types** — bar, column, line, pie, scatter, area, doughnut, waterfall, funnel, treemap
- **Pivot tables** with row/column/value/filter fields and multiple aggregation modes
- **Conditional formatting** — cell value rules, color scales, data bars, icon sets
- **Data validation** — dropdowns, number ranges, date ranges, custom formulas
- **Tables, sparklines, images, shapes** — full Excel feature coverage
- **Two transports** — stdio for CLI clients, streamable HTTP for web clients
- **In-memory workbook store** with TTL eviction and capacity limits

## Quick Start

### Prerequisites

- [Rust](https://rustup.rs/) 1.75+
- [zavora-xlsx](https://github.com/zavora-ai/zavora-xlsx) cloned locally (path dependency)

### Build

```bash
cargo build --release
```

The binary is at `target/release/excel-mcp-server`.

### Run (stdio)

```bash
./target/release/excel-mcp-server
```

The server reads MCP JSON-RPC messages from stdin and writes responses to stdout. Logs go to stderr.

### Run (HTTP)

```bash
./target/release/excel-mcp-server http
```

Starts a streamable HTTP server on `127.0.0.1:8080`. Configure with `BIND_ADDRESS`:

```bash
BIND_ADDRESS=0.0.0.0:3000 ./target/release/excel-mcp-server http
```

## Client Configuration

### Claude Desktop

Add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "excel": {
      "command": "/path/to/excel-mcp-server"
    }
  }
}
```

### Kiro

Add to `.kiro/settings/mcp.json`:

```json
{
  "mcpServers": {
    "excel": {
      "command": "/path/to/excel-mcp-server",
      "autoApprove": []
    }
  }
}
```

### Cursor

Add to `.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "excel": {
      "command": "/path/to/excel-mcp-server"
    }
  }
}
```

### HTTP Client

Point any MCP client at `http://localhost:8080/mcp` when running in HTTP mode.

## Tools Reference

### Workbook Lifecycle (4 tools)

| Tool | Description |
|---|---|
| `create_workbook` | Create a new empty Excel workbook in memory |
| `open_workbook` | Open an existing Excel file for reading or editing |
| `save_workbook` | Save a workbook to disk as an xlsx file |
| `close_workbook` | Close a workbook and free its memory |

### Workbook Configuration (1 tool)

| Tool | Description |
|---|---|
| `configure_workbook` | Set calc mode (auto/manual), active sheet, document properties (title, author, company) |

### Sheet Management (7 tools)

| Tool | Description |
|---|---|
| `list_sheets` | List all sheet names in a workbook |
| `get_sheet_dimensions` | Get the dimensions of a sheet (used range, row count, column count) |
| `describe_workbook` | Describe a workbook's structure including sheet names, dimensions, and sample data |
| `add_sheet` | Add a new empty worksheet |
| `rename_sheet` | Rename an existing worksheet |
| `delete_sheet` | Delete a worksheet |
| `move_worksheet` | Move a worksheet to a different position |

### Reading Data (4 tools)

| Tool | Description |
|---|---|
| `read_sheet` | Read data from a worksheet with optional range and pagination |
| `read_cell` | Read a single cell's value, type, and formula |
| `search_cells` | Search for cells matching a value or pattern across sheets |
| `sheet_to_csv` | Export a sheet as a CSV-formatted string |

### Writing Data (4 tools)

| Tool | Description |
|---|---|
| `write_cells` | Write values to multiple cells (auto-detects numbers, booleans, dates, formulas) |
| `write_row` | Write a row of values starting from a cell, filling rightward |
| `write_column` | Write a column of values starting from a cell, filling downward |
| `write_rich_text` | Write rich text (mixed bold/italic/color runs) to a cell |

### Formulas (1 tool)

| Tool | Description |
|---|---|
| `write_formula` | Write regular, array (CSE), or dynamic (Excel 365 spill) formulas with optional cached result |

### Cell Operations (1 tool)

| Tool | Description |
|---|---|
| `manage_cell` | Write a formatted blank cell or clear content and formatting |

### Formatting (4 tools)

| Tool | Description |
|---|---|
| `set_cell_format` | Apply formatting (bold, italic, colors, borders, number format, alignment) to a range |
| `merge_cells` | Merge a range of cells into a single cell |
| `set_row_column_format` | Apply formatting to an entire row or column |
| `set_dimensions` | Set column width, row height, column range width, or default row height |

### Layout (5 tools)

| Tool | Description |
|---|---|
| `freeze_panes` | Freeze panes at a cell position for scrolling |
| `autofit_columns` | Auto-fit all column widths based on cell content |
| `set_selection` | Set the selected/active cell in a sheet |
| `set_visibility` | Hide or unhide a row or column |
| `set_sheet_settings` | Configure sheet display: hidden, zoom, gridlines, tab color, right-to-left |

### Charts (4 tools)

| Tool | Description |
|---|---|
| `add_chart` | Add a chart (bar/column/line/pie/scatter/area/doughnut) with multiple series, trendlines, markers, pivot source |
| `add_waterfall_chart` | Add a waterfall chart (Excel 2016+ ChartEx) with increase/decrease/total points |
| `add_funnel_chart` | Add a funnel chart (Excel 2016+ ChartEx) |
| `add_treemap_chart` | Add a treemap chart (Excel 2016+ ChartEx) with optional per-point colors |

### Tables & Data Features (4 tools)

| Tool | Description |
|---|---|
| `add_table` | Create an Excel Table with headers, autofilter, and style |
| `add_conditional_format` | Add conditional formatting: cell value, color scales, data bars, icon sets |
| `add_data_validation` | Add data validation: dropdowns, number/date ranges, custom formulas |
| `add_sparkline` | Add a sparkline (line, column, or win/loss) to a cell |

### Images & Shapes (2 tools)

| Tool | Description |
|---|---|
| `add_image` | Embed a PNG or JPEG image into a worksheet |
| `add_shape` | Add a drawing shape (rectangle, ellipse, arrow, callout, text box, etc.) with text, fill, and outline |

### Pivot Tables (1 tool)

| Tool | Description |
|---|---|
| `add_pivot_table` | Create a pivot table with row/column/value/filter fields, aggregation, and layout options |

### Page Setup & Print (1 tool)

| Tool | Description |
|---|---|
| `set_page_setup` | Configure landscape, paper size, margins, fit-to-page, print area, headers/footers, gridlines |

### Comments & Links (2 tools)

| Tool | Description |
|---|---|
| `manage_comments` | Add or read cell comments/notes |
| `add_link` | Add external URLs or internal sheet references |

### Named Ranges (1 tool)

| Tool | Description |
|---|---|
| `manage_defined_names` | Add or list defined names (named ranges) |

### Row/Column Manipulation (2 tools)

| Tool | Description |
|---|---|
| `modify_rows` | Insert or delete rows |
| `modify_columns` | Insert or delete columns |

### Grouping & Protection (2 tools)

| Tool | Description |
|---|---|
| `group` | Group rows or columns into expandable outlines |
| `protect` | Protect sheets, workbooks, or unprotect specific ranges |

### Autofilter & Errors (2 tools)

| Tool | Description |
|---|---|
| `manage_autofilter` | Set autofilter on a range with optional column filtering |
| `ignore_error` | Suppress Excel error indicators (green triangles) on a range |

### Document Properties (1 tool)

| Tool | Description |
|---|---|
| `set_doc_properties` | Set title, author, subject, description, keywords, category, company |

## Architecture

```
src/
├── main.rs          # Entry point: stdio or HTTP transport selection
├── server.rs        # MCP tool router — all 43 tools registered here
├── store.rs         # In-memory workbook store with TTL eviction
├── lib.rs           # Crate root
├── error.rs         # Error types
├── cell_ref.rs      # A1 notation parsing utilities
├── engines/
│   └── zavora.rs    # zavora-xlsx engine adapter
├── tools/
│   ├── workbook.rs  # create, open, save, close
│   ├── read.rs      # read_sheet, read_cell, search, CSV export
│   ├── write.rs     # write_cells, write_row, write_column
│   ├── format.rs    # set_cell_format, merge_cells
│   ├── charts.rs    # Legacy chart support
│   ├── tables.rs    # Excel Tables
│   ├── conditional.rs # Conditional formatting
│   ├── validation.rs  # Data validation
│   ├── sparklines.rs  # Sparklines
│   ├── images.rs    # Image embedding
│   ├── layout.rs    # Freeze panes
│   ├── sheets.rs    # Sheet management
│   └── expanded.rs  # All remaining tools (pivot tables, shapes, charts, etc.)
└── types/
    ├── inputs.rs    # Deserialized input structs for all tools
    ├── enums.rs     # Shared enums (chart types, formats, etc.)
    └── responses.rs # Structured JSON response builders
```

### Workbook Store

The server maintains an in-memory store of open workbooks:

- **Capacity**: 10 concurrent workbooks (configurable)
- **TTL**: 30-minute inactivity timeout — workbooks are automatically evicted
- **Thread-safe**: Protected by `Arc<RwLock<WorkbookStore>>`

Each `create_workbook` or `open_workbook` call returns a `workbook_id` handle. All subsequent operations reference this handle. Call `save_workbook` before the TTL expires to persist changes, and `close_workbook` to free memory.

### Response Format

All tools return structured JSON:

```json
{
  "status": "success",
  "message": "Descriptive message",
  "data": { ... }
}
```

On error:

```json
{
  "status": "error",
  "category": "not_found",
  "message": "Workbook not found",
  "suggestion": "Check the workbook_id"
}
```

## Example Workflow

Here's what a typical interaction looks like when an AI assistant uses this server:

1. **Create a workbook**
   ```
   → create_workbook {}
   ← { workbook_id: "abc-123", sheets: ["Sheet1"] }
   ```

2. **Write headers and data**
   ```
   → write_row { workbook_id: "abc-123", sheet_name: "Sheet1", start_cell: "A1", values: ["Product", "Q1", "Q2", "Q3", "Q4"] }
   → write_row { workbook_id: "abc-123", sheet_name: "Sheet1", start_cell: "A2", values: ["Widget", 150, 200, 180, 220] }
   ```

3. **Format headers**
   ```
   → set_cell_format { workbook_id: "abc-123", sheet_name: "Sheet1", range: "A1:E1", bold: true, background_color: "#4472C4", font_color: "#FFFFFF" }
   ```

4. **Add a chart**
   ```
   → add_chart { workbook_id: "abc-123", sheet_name: "Sheet1", chart_type: "column", series: [{ values: "Sheet1!$B$2:$E$2", categories: "Sheet1!$B$1:$E$1", name: "Widget" }], title: "Quarterly Sales", cell: "A5" }
   ```

5. **Save**
   ```
   → save_workbook { workbook_id: "abc-123", file_path: "./quarterly_report.xlsx" }
   ```

## Development

```bash
# Build
cargo build

# Run tests
cargo test

# Lint
cargo clippy

# Format
cargo fmt
```

## License

Apache 2.0 — see [LICENSE](LICENSE) for details.

Copyright 2025 Zavora Technologies Ltd.

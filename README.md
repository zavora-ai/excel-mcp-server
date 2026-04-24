# Excel MCP Server

A high-performance [Model Context Protocol](https://modelcontextprotocol.io/) (MCP) server that gives AI assistants full control over Excel spreadsheets. Built in Rust with [zavora-xlsx](https://github.com/zavora-ai/zavora-xlsx) for native xlsx read/write and [rmcp](https://github.com/modelcontextprotocol/rust-sdk) for the MCP protocol layer.

**74 tools** covering the complete Excel feature set — cell I/O, formatting, 14 chart types, pivot tables with calculated fields, slicers, timelines, form controls, conditional formatting, data validation, images, shapes, threaded comments, formula recalculation, encrypted save/open, and more.

## Install

```bash
cargo install excel-mcp-server
```

This compiles and installs the binary to `~/.cargo/bin/excel-mcp-server`.

> Requires [Rust](https://rustup.rs/) 1.85+. If you don't have Rust: `curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh`

## Client Configuration

### Claude Desktop

Add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "excel": {
      "command": "excel-mcp-server"
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
      "command": "excel-mcp-server",
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
      "command": "excel-mcp-server"
    }
  }
}
```

### HTTP Mode

For web-based MCP clients:

```bash
excel-mcp-server http
```

Starts a streamable HTTP server on `127.0.0.1:8080`. Point your client at `http://localhost:8080/mcp`.

```bash
BIND_ADDRESS=0.0.0.0:3000 excel-mcp-server http
```

## Features

- **Create, open, edit, and save** xlsx files through natural language
- **Read and analyze** spreadsheets with pagination, search, and CSV export
- **Full formatting** — bold, italic, colors, borders, number formats, alignment, merge
- **14 chart types** — bar, column, line, pie, scatter, area, doughnut, waterfall, funnel, treemap, sunburst, histogram, box & whisker, map
- **Chart enhancements** — data tables, 3D views, error bars, axis formatting/bounds/log scale, drop lines, high-low lines, gradients, bubble sizes, alt text, preset styles
- **Pivot tables** — calculated fields, date/range grouping, subtotals, grand totals, value formats, row stripes, layout options
- **Slicers and timelines** — interactive pivot table filters
- **Form controls** — buttons, checkboxes, dropdowns, spinners
- **Conditional formatting** — cell value rules, 2/3-color scales, data bars, icon sets
- **Data validation** — dropdowns, number/date ranges, text length, custom formulas
- **Tables, sparklines, images, shapes** — full Excel feature coverage
- **Threaded comments** — modern conversation-style comments with replies and timestamps
- **Formula recalculation** — 58-function engine with dependency graph, evaluated before every save
- **Multiple save formats** — xlsx, template (.xltx), encrypted (password-protected), parallel compression
- **Open encrypted** — password-protected workbook support
- **Named ranges** — full CRUD with workbook and sheet-level scoping
- **Sheet metadata** — read used range, hyperlinks, merged ranges, embedded charts
- **Custom XML parts** — add and read custom XML by namespace
- **Custom document properties** — text, number, integer, boolean, datetime
- **Two transports** — stdio for CLI clients, streamable HTTP for web clients
- **In-memory workbook store** — 10-workbook capacity with 30-minute TTL eviction

## Tools Reference (74 tools)

### Workbook Lifecycle (4)

| Tool | Description |
|---|---|
| `create_workbook` | Create a new empty workbook in memory |
| `open_workbook` | Open an existing xlsx file (edit or read-only mode) |
| `save_workbook` | Save to disk as xlsx (formulas recalculated automatically) |
| `close_workbook` | Close and free memory |

### Workbook Configuration (1)

| Tool | Description |
|---|---|
| `configure_workbook` | Set calc mode (auto/manual), active sheet, document properties |

### Sheet Management (7)

| Tool | Description |
|---|---|
| `list_sheets` | List all sheet names |
| `get_sheet_dimensions` | Get used range, row count, column count |
| `describe_workbook` | Sheet names, dimensions, and sample data |
| `add_sheet` | Add a new worksheet |
| `rename_sheet` | Rename a worksheet |
| `delete_sheet` | Delete a worksheet |
| `move_worksheet` | Reorder a worksheet |

### Reading (4)

| Tool | Description |
|---|---|
| `read_sheet` | Read data with optional range and pagination |
| `read_cell` | Read a cell's value, type, and formula |
| `search_cells` | Search by value or pattern across sheets |
| `sheet_to_csv` | Export as CSV string |

### Writing (5)

| Tool | Description |
|---|---|
| `write_cells` | Batch write to multiple cells (auto-detects types) |
| `write_row` | Write values rightward from a cell |
| `write_column` | Write values downward from a cell |
| `write_rich_text` | Write mixed bold/italic/color text runs |
| `write_json_rows` | Write JSON objects as rows with auto headers and type detection |

### Formulas (1)

| Tool | Description |
|---|---|
| `write_formula` | Regular, array (CSE), or dynamic (Excel 365 spill) formulas |

### Cell Operations (1)

| Tool | Description |
|---|---|
| `manage_cell` | Write formatted blank or clear content/formatting |

### Formatting (4)

| Tool | Description |
|---|---|
| `set_cell_format` | Bold, italic, colors, borders, number format, alignment on a range |
| `merge_cells` | Merge a range into one cell |
| `set_row_column_format` | Format an entire row or column |
| `set_dimensions` | Column width, row height, column range width, default row height |

### Layout (5)

| Tool | Description |
|---|---|
| `freeze_panes` | Freeze rows/columns for scrolling |
| `autofit_columns` | Auto-fit widths based on content |
| `set_selection` | Set the active cell |
| `set_visibility` | Hide/unhide rows or columns |
| `set_sheet_settings` | Hidden, zoom, gridlines, tab color, right-to-left |

### Charts (8)

| Tool | Description |
|---|---|
| `add_chart` | Bar/column/line/pie/scatter/area/doughnut with series, trendlines, markers, error bars, gradients, axis formatting, data tables, 3D views, alt text, preset styles |
| `add_waterfall_chart` | Waterfall (ChartEx) with increase/decrease/total points |
| `add_funnel_chart` | Funnel (ChartEx) |
| `add_treemap_chart` | Treemap (ChartEx) with per-point colors |
| `add_sunburst_chart` | Sunburst (ChartEx) for hierarchical data |
| `add_histogram_chart` | Histogram with bin control and Pareto overlay |
| `add_box_whisker_chart` | Box & whisker with outliers, mean markers, inner points |
| `add_map_chart` | Geographic map with country/region levels |

### Interactive Controls (3)

| Tool | Description |
|---|---|
| `add_slicer` | Interactive pivot table filter |
| `add_timeline` | Date-based pivot table filter |
| `add_form_control` | Button, checkbox, dropdown, or spinner |

### Tables & Data Features (4)

| Tool | Description |
|---|---|
| `add_table` | Excel Table with headers, autofilter, style |
| `add_conditional_format` | Cell value, color scales, data bars, icon sets |
| `add_data_validation` | Dropdowns, number/date ranges, custom formulas |
| `add_sparkline` | Line, column, or win/loss sparkline |

### Images & Shapes (2)

| Tool | Description |
|---|---|
| `add_image` | Embed PNG or JPEG |
| `add_shape` | Rectangle, ellipse, arrow, callout, text box with fill and outline |

### Pivot Tables (1)

| Tool | Description |
|---|---|
| `add_pivot_table` | Row/column/value/filter fields, aggregation, calculated fields, date/range grouping, subtotals, grand totals, value formats, layout |

### Page Setup (1)

| Tool | Description |
|---|---|
| `set_page_setup` | Landscape, paper size, margins, fit-to-page, print area, headers/footers, gridlines |

### Comments & Links (3)

| Tool | Description |
|---|---|
| `manage_comments` | Add or read legacy comments/notes |
| `add_threaded_comment` | Modern threaded comments with replies and timestamps |
| `add_link` | External URLs or internal sheet references |

### Named Ranges (2)

| Tool | Description |
|---|---|
| `manage_defined_names` | Add or list defined names |
| `manage_named_ranges` | Full CRUD: add, add_scoped, update, remove, list with scope |

### Row/Column Manipulation (2)

| Tool | Description |
|---|---|
| `modify_rows` | Insert or delete rows |
| `modify_columns` | Insert or delete columns |

### Protection (3)

| Tool | Description |
|---|---|
| `group` | Group rows/columns into outlines |
| `protect` | Protect sheet/workbook or unprotect ranges |
| `protect_sheet_advanced` | Granular protection (allow/deny insert, delete, format, sort per feature) |

### Autofilter & Errors (2)

| Tool | Description |
|---|---|
| `manage_autofilter` | Set autofilter with optional column filtering |
| `ignore_error` | Suppress green triangle error indicators |

### Document Properties (2)

| Tool | Description |
|---|---|
| `set_doc_properties` | Title, author, subject, description, keywords, category, company |
| `set_custom_property` | Custom properties (text, number, integer, bool, datetime) |

### Advanced Save/Open (2)

| Tool | Description |
|---|---|
| `save_workbook_advanced` | Save as template, encrypted, or parallel. Formulas recalculated. |
| `open_workbook_encrypted` | Open password-protected workbooks |

### Read Enhancements (3)

| Tool | Description |
|---|---|
| `read_cell_comment` | Read a single cell's comment |
| `read_cell_format` | Read a cell's formatting |
| `read_sheet_metadata` | Used range, hyperlinks, merged ranges, charts — individually or all |

### Workbook Features (4)

| Tool | Description |
|---|---|
| `add_chart_sheet` | Dedicated chart-only sheet |
| `manage_custom_xml` | Add or read custom XML parts by namespace |
| `add_connection` | Add external data connection |
| `set_sst_threshold` | Shared string table optimization threshold |

## Example Workflow

```
1. create_workbook {}
   → { workbook_id: "abc-123", sheets: ["Sheet1"] }

2. write_row { workbook_id: "abc-123", sheet_name: "Sheet1",
     start_cell: "A1", values: ["Product", "Q1", "Q2", "Q3", "Q4"] }

3. write_row { ..., start_cell: "A2", values: ["Widget", 150, 200, 180, 220] }

4. set_cell_format { ..., range: "A1:E1", bold: true,
     background_color: "#4472C4", font_color: "#FFFFFF" }

5. add_chart { ..., chart_type: "column",
     series: [{ values: "Sheet1!$B$2:$E$2",
                categories: "Sheet1!$B$1:$E$1", name: "Widget" }],
     title: "Quarterly Sales", cell: "A5",
     style: 26, show_data_table: true }

6. save_workbook { ..., file_path: "./quarterly_report.xlsx" }
```

## Architecture

```
src/
├── main.rs            Entry point — stdio or HTTP transport
├── server.rs          MCP tool router — 74 tools registered
├── store.rs           In-memory workbook store with TTL eviction
├── lib.rs             Crate root
├── error.rs           Error types
├── cell_ref.rs        A1 notation parsing
├── engines/
│   └── zavora.rs      zavora-xlsx engine adapter + recalculate
├── tools/
│   ├── workbook.rs    create, open, save, close
│   ├── read.rs        read_sheet, read_cell, search, CSV
│   ├── write.rs       write_cells, write_row, write_column
│   ├── format.rs      set_cell_format, merge_cells
│   ├── charts.rs      Basic chart support
│   ├── tables.rs      Excel Tables
│   ├── conditional.rs Conditional formatting
│   ├── validation.rs  Data validation
│   ├── sparklines.rs  Sparklines
│   ├── images.rs      Image embedding
│   ├── layout.rs      Freeze panes
│   ├── sheets.rs      Sheet management
│   └── expanded.rs    All advanced tools (charts, pivots, controls, etc.)
└── types/
    ├── inputs.rs      Input structs for all 74 tools
    ├── enums.rs       Chart types, formats, validation rules
    └── responses.rs   JSON response builders
```

### Workbook Store

- **Capacity**: 10 concurrent workbooks
- **TTL**: 30-minute inactivity timeout with automatic eviction
- **Thread-safe**: `Arc<RwLock<WorkbookStore>>`
- Every `create_workbook` or `open_workbook` returns a `workbook_id` handle
- All operations reference this handle
- `save_workbook` recalculates all formulas before writing

### Response Format

Success:
```json
{ "status": "success", "message": "...", "data": { ... } }
```

Error:
```json
{ "status": "error", "category": "not_found", "message": "...", "suggestion": "..." }
```

Error categories: `not_found`, `io_error`, `engine_unsupported`, `capacity_exceeded`

## Development

```bash
git clone https://github.com/zavora-ai/excel-mcp-server.git
cd excel-mcp-server
cargo build --release    # Binary at target/release/excel-mcp-server
cargo test               # Run all tests
cargo clippy             # Lint
cargo fmt                # Format
```

## License

Apache 2.0 — see [LICENSE](LICENSE).

Copyright 2025 Zavora Technologies Ltd.

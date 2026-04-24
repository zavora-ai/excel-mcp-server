# Changelog

All notable changes to this project will be documented in this file.

## [0.2.0] - 2025-07-25

### Critical Fix

- **Formula recalculation**: `save_workbook` now calls `Workbook::recalculate()` before writing to disk. All formula cells are evaluated with dependency graph ordering, circular reference detection, and volatile function support. Previously, formula cells saved with cached value 0.

### New Chart Types (4 tools)

- `add_sunburst_chart` — hierarchical sunburst chart (Excel 2016+ ChartEx)
- `add_histogram_chart` — histogram with optional bin_count, bin_width, and Pareto overlay
- `add_box_whisker_chart` — box & whisker with outliers, mean markers, inner points
- `add_map_chart` — geographic map chart with country/region levels

### Interactive Controls (3 tools)

- `add_slicer` — interactive filter for pivot tables
- `add_timeline` — date filter for pivot tables
- `add_form_control` — button, checkbox, dropdown, spinner form controls

### Advanced Save/Open (2 tools)

- `save_workbook_advanced` — save as template (.xltx), encrypted (password-protected), or parallel (fast compression)
- `open_workbook_encrypted` — open password-protected workbooks

### Named Ranges (1 tool)

- `manage_named_ranges` — full CRUD: add, add_scoped (sheet-level), update, remove, list with scope info

### Sheet Metadata (1 tool)

- `read_sheet_metadata` — read used_range, hyperlinks, merge_ranges, charts, or all at once

### Chart Sheet (1 tool)

- `add_chart_sheet` — dedicated chart-only sheet (no cells)

### Changed

- Upgraded zavora-xlsx dependency from 0.1.0 to 0.1.1 (adds `Workbook::recalculate()`)
- Tool count increased from 43 to 56

## [0.1.1] - 2025-07-25

### Changed

- Upgraded to Rust edition 2024 (requires Rust 1.85+)
- Updated repository URL to `zavora-ai/excel-mcp-server`
- Rewrote README to lead with `cargo install` instead of build-from-source instructions
- Updated minimum Rust version in docs from 1.75+ to 1.85+

### Added

- CHANGELOG.md

## [0.1.0] - 2025-07-25

Initial public release of Excel MCP Server.

### Core

- MCP server with **43 tools** for full Excel spreadsheet manipulation
- **Two transports**: stdio (for CLI clients) and streamable HTTP (for web clients)
- In-memory workbook store with 10-workbook capacity and 30-minute TTL eviction
- Thread-safe architecture using `Arc<RwLock<WorkbookStore>>`
- Built on [zavora-xlsx](https://github.com/zavora-ai/zavora-xlsx) for native xlsx read/write
- Built on [rmcp](https://github.com/modelcontextprotocol/rust-sdk) for MCP protocol compliance
- Rust edition 2024

### Workbook Lifecycle

- Create, open, save, and close xlsx workbooks
- Configure calc mode, active sheet, and document properties

### Sheet Management

- List, add, rename, delete, and reorder worksheets
- Get sheet dimensions and describe workbook structure

### Reading

- Read sheets with optional range selection and pagination
- Read individual cells with value, type, and formula info
- Search cells by value or pattern across sheets
- Export sheets as CSV strings

### Writing

- Write to multiple cells with auto-detection of numbers, booleans, dates, and formulas
- Write rows and columns from a starting cell
- Write rich text with mixed bold/italic/color formatting runs
- Write regular, array (CSE), and dynamic (Excel 365 spill) formulas

### Formatting

- Cell formatting: bold, italic, font/background colors, borders, number formats, alignment
- Merge cells
- Row and column formatting
- Set column widths, row heights, and default dimensions

### Layout

- Freeze panes
- Auto-fit column widths
- Set active cell selection
- Hide/unhide rows and columns
- Sheet display settings: zoom, gridlines, tab color, right-to-left, hidden

### Charts

- 10 chart types: bar, column, line, pie, scatter, area, doughnut, waterfall, funnel, treemap
- Multiple series per chart with trendlines and markers
- Waterfall, funnel, and treemap via Excel 2016+ ChartEx format
- Pivot table as chart data source

### Tables & Data Features

- Excel Tables with headers, autofilter, and styles
- Conditional formatting: cell value rules, color scales, data bars, icon sets
- Data validation: dropdowns, number/date ranges, custom formulas
- Sparklines: line, column, and win/loss

### Images & Shapes

- Embed PNG and JPEG images
- Drawing shapes: rectangle, ellipse, arrow, callout, text box with text, fill, and outline

### Pivot Tables

- Row, column, value, and filter fields
- Multiple aggregation modes
- Layout and style options

### Other Features

- Page setup and print configuration (landscape, paper size, margins, fit-to-page, headers/footers)
- Cell comments and notes
- Hyperlinks (external URLs and internal sheet references)
- Named ranges (defined names)
- Insert and delete rows and columns
- Row and column grouping (outlines)
- Sheet, workbook, and range protection
- Autofilter with column filtering
- Suppress Excel error indicators on ranges
- Document properties (title, author, subject, description, keywords, category, company)

### Publishing

- Published to [crates.io](https://crates.io/crates/excel-mcp-server) as `excel-mcp-server`
- Install with `cargo install excel-mcp-server`
- Apache 2.0 license

[0.2.0]: https://github.com/zavora-ai/excel-mcp-server/compare/v0.1.1...v0.2.0
[0.1.1]: https://github.com/zavora-ai/excel-mcp-server/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/zavora-ai/excel-mcp-server/releases/tag/v0.1.0

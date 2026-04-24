# Changelog

All notable changes to this project will be documented in this file.

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

[0.1.0]: https://github.com/zavora-ai/excel-mcp-server/releases/tag/v0.1.0

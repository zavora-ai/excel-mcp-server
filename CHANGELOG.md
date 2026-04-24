# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.0] - 2025-07-25

Major release: full feature parity with zavora-xlsx 0.1.1. Tool count increased from 43 to 74.

### Critical Fix

- **Formula recalculation**: `save_workbook` and `save_workbook_advanced` now call `Workbook::recalculate()` before writing to disk. All formula cells are evaluated using a dependency graph with topological ordering, circular reference detection, and volatile function support (58 built-in functions: SUM, IF, VLOOKUP, XLOOKUP, INDEX, MATCH, and more). Previously, formula cells saved with cached value 0.

### New Chart Types (4 new tools)

- `add_sunburst_chart` — hierarchical sunburst chart (Excel 2016+ ChartEx) with levels and values
- `add_histogram_chart` — histogram with bin_count, bin_width, and Pareto overlay option
- `add_box_whisker_chart` — box & whisker with outlier, mean marker, and inner point toggles
- `add_map_chart` — geographic map chart with country/region levels

### Chart Enhancements (extended `add_chart`)

The existing `add_chart` tool now supports:
- `show_data_table` — data table below chart
- `view_3d` — 3D perspective rotation (rot_x, rot_y, perspective)
- `style` — preset chart style number (1-48)
- `alt_text` — accessibility title and description
- `y_axis_min` / `y_axis_max` — axis bounds
- `y_axis_log_base` — logarithmic scale
- `x_axis_reverse` / `y_axis_reverse` — reverse axis direction
- `x_axis_format` / `y_axis_format` — number format on axes
- `drop_lines` / `high_low_lines` — line chart overlays
- `plot_area_fill` — background fill color

Per-series enhancements:
- `line_width` — line thickness in points
- `dash_style` — solid, dash, dot, dash_dot, long_dash, long_dash_dot
- `gradient` — gradient fill with color stops
- `bubble_sizes` — bubble chart size range
- `error_bars` — fixed value, percentage, standard deviation, or standard error

### Pivot Table Enhancements (extended `add_pivot_table`)

- `calculated_fields` — custom formula fields in the pivot table
- `date_groups` — group date fields by years, quarters, months, days, hours, minutes, seconds
- `range_groups` — group numeric fields into ranges (start, end, interval)
- `value_formats` — number format per value field
- `subtotals` — toggle subtotals per field
- `grand_total_rows` / `grand_total_cols` — toggle grand totals
- `show_row_headers` / `show_column_headers` — toggle headers
- `show_row_stripes` — banded rows

### Interactive Controls (3 new tools)

- `add_slicer` — interactive filter connected to a pivot table field
- `add_timeline` — date-based filter connected to a pivot table date field
- `add_form_control` — button, checkbox, dropdown, or spinner with cell link support

### Threaded Comments (1 new tool)

- `add_threaded_comment` — modern threaded comments with author, text, timestamp, and replies (replaces legacy notes for conversation-style commenting)

### Protection Enhancements (1 new tool)

- `protect_sheet_advanced` — granular sheet protection with per-feature allow/deny: insert_rows, delete_rows, insert_columns, delete_columns, format_cells, format_columns, format_rows, sort, insert_hyperlinks, select_locked_cells, select_unlocked_cells, pivot_tables

### Advanced Save/Open (2 new tools)

- `save_workbook_advanced` — save as template (.xltx), encrypted (password-protected), or with parallel compression. Formulas recalculated before save.
- `open_workbook_encrypted` — open password-protected Excel workbooks

### Named Ranges (1 new tool)

- `manage_named_ranges` — full CRUD operations: add (workbook-scoped), add_scoped (sheet-scoped), update, remove, list with scope information

### Read Enhancements (3 new tools)

- `read_sheet_metadata` — read used_range, hyperlinks, merged ranges, embedded charts, or all at once
- `read_cell_comment` — read a single cell's comment (author and text)
- `read_cell_format` — read a cell's formatting properties

### Document Properties (1 new tool)

- `set_custom_property` — set custom document properties with typed values (text, number, integer, bool, datetime)

### Workbook Features (4 new tools)

- `add_chart_sheet` — dedicated chart-only sheet (no cells, full-page chart)
- `manage_custom_xml` — add or read custom XML parts by namespace
- `add_connection` — add external data connections
- `set_sst_threshold` — tune shared string table threshold for write performance

### Data Import (1 new tool)

- `write_json_rows` — write JSON objects as spreadsheet rows with automatic header generation and type detection

### Changed

- Upgraded zavora-xlsx from 0.1.0 to 0.1.1 (adds `Workbook::recalculate()`)
- Upgraded Rust edition from 2021 to 2024 (requires Rust 1.85+)
- Updated repository URL to `zavora-ai/excel-mcp-server`
- Rewrote README with comprehensive tool reference and architecture docs
- Tool count: 43 → 74

## [0.1.1] - 2025-07-25

### Changed

- Upgraded to Rust edition 2024 (requires Rust 1.85+)
- Updated repository URL to `zavora-ai/excel-mcp-server`
- Rewrote README to lead with `cargo install` instead of build-from-source
- Updated minimum Rust version in docs from 1.75+ to 1.85+

### Added

- CHANGELOG.md

## [0.1.0] - 2025-07-25

Initial public release.

### Core

- MCP server with 43 tools for Excel spreadsheet manipulation
- Two transports: stdio and streamable HTTP
- In-memory workbook store with 10-workbook capacity and 30-minute TTL
- Thread-safe architecture using `Arc<RwLock<WorkbookStore>>`
- Built on zavora-xlsx for native xlsx read/write
- Built on rmcp for MCP protocol compliance

### Tools (43)

Workbook lifecycle (4), sheet management (7), reading (4), writing (4), formulas (1), cell operations (1), formatting (4), layout (5), charts — bar, column, line, pie, scatter, area, doughnut, waterfall, funnel, treemap (4), tables (1), conditional formatting (1), data validation (1), sparklines (1), images (1), shapes (1), pivot tables (1), page setup (1), comments (1), links (1), named ranges (1), row/column manipulation (2), grouping (1), protection (1), autofilter (1), error suppression (1), document properties (1).

[0.2.0]: https://github.com/zavora-ai/excel-mcp-server/compare/v0.1.1...v0.2.0
[0.1.1]: https://github.com/zavora-ai/excel-mcp-server/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/zavora-ai/excel-mcp-server/releases/tag/v0.1.0

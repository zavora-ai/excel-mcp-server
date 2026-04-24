# Excel MCP Server — Detailed Documentation

Complete reference for all 74 tools in the Excel MCP Server.

## Table of Contents

- [Concepts](#concepts)
- [Workbook Lifecycle](#workbook-lifecycle)
- [Reading Data](#reading-data)
- [Writing Data](#writing-data)
- [Formatting](#formatting)
- [Charts](#charts)
- [Pivot Tables](#pivot-tables)
- [Interactive Controls](#interactive-controls)
- [Tables and Data Features](#tables-and-data-features)
- [Images and Shapes](#images-and-shapes)
- [Comments and Links](#comments-and-links)
- [Named Ranges](#named-ranges)
- [Protection](#protection)
- [Page Setup and Print](#page-setup-and-print)
- [Document Properties](#document-properties)
- [Advanced Features](#advanced-features)
- [Error Handling](#error-handling)

## Concepts

### Workbook IDs

Every workbook operation requires a `workbook_id` — a UUID handle returned by `create_workbook` or `open_workbook`. This handle references an in-memory workbook in the server's store.

### Cell References

Cells use A1 notation: `"A1"`, `"B5"`, `"AA100"`. Ranges use colon notation: `"A1:E10"`. Sheet-qualified references: `"Sheet1!$A$1:$B$10"`.

### Formula Recalculation

The server includes a 58-function formula engine. When you call `save_workbook` or `save_workbook_advanced`, all formulas are automatically recalculated with dependency graph ordering before the file is written. Supported functions include SUM, AVERAGE, IF, VLOOKUP, XLOOKUP, INDEX, MATCH, COUNTIF, SUMIF, and many more.

### Workbook Store

- Maximum 10 concurrent workbooks
- 30-minute inactivity TTL — workbooks are evicted automatically
- Thread-safe with `Arc<RwLock<WorkbookStore>>`
- Always call `save_workbook` before the TTL expires to persist changes

## Workbook Lifecycle

### create_workbook

Create a new empty workbook with one sheet ("Sheet1").

```json
{ }
```

Returns `workbook_id`, engine name, and sheet list.

### open_workbook

Open an existing xlsx file.

```json
{
  "file_path": "/path/to/file.xlsx",
  "read_only": false
}
```

Set `read_only: true` for fast read-only access (cannot save).

### open_workbook_encrypted

Open a password-protected xlsx file.

```json
{
  "file_path": "/path/to/encrypted.xlsx",
  "password": "secret"
}
```

### save_workbook

Save to disk as xlsx. Formulas are recalculated automatically.

```json
{
  "workbook_id": "...",
  "file_path": "/path/to/output.xlsx"
}
```

### save_workbook_advanced

Save in different formats with recalculation.

```json
{
  "workbook_id": "...",
  "file_path": "/path/to/output.xltx",
  "format": "template",
  "password": null
}
```

Formats: `"xlsx"` (default), `"template"` (.xltx), `"encrypted"` (password-protected), `"parallel"` (fast compression for large files).

### close_workbook

Free memory for a workbook.

```json
{ "workbook_id": "..." }
```

### configure_workbook

Set calculation mode, active sheet, and basic properties.

```json
{
  "workbook_id": "...",
  "calc_mode": "auto",
  "active_sheet": 0,
  "title": "Report",
  "author": "AI Assistant"
}
```

## Reading Data

### read_sheet

Read data with optional range and pagination.

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "range": "A1:E10",
  "continuation_token": null
}
```

### read_cell

Read a single cell's value, type, and formula.

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "cell": "B5"
}
```

### read_cell_comment

Read a cell's comment (author and text).

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "cell": "A1"
}
```

### read_cell_format

Read a cell's formatting properties.

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "cell": "A1"
}
```

### read_sheet_metadata

Read structural metadata about a sheet.

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "info": "all"
}
```

`info` options: `"used_range"`, `"hyperlinks"`, `"merge_ranges"`, `"charts"`, `"all"`.

### search_cells

Search for values across sheets.

```json
{
  "workbook_id": "...",
  "query": "Revenue",
  "match_mode": "substring"
}
```

### sheet_to_csv

Export a sheet as a CSV string.

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "delimiter": ","
}
```

## Writing Data

### write_cells

Batch write to multiple cells. Types are auto-detected: strings starting with `=` become formulas, numbers stay numeric, ISO dates are parsed, booleans are preserved.

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "cells": [
    { "cell": "A1", "value": "Name" },
    { "cell": "B1", "value": 42 },
    { "cell": "C1", "value": "=SUM(B1:B10)" }
  ]
}
```

### write_row / write_column

Write a sequence of values starting from a cell.

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "start_cell": "A1",
  "values": ["Product", "Q1", "Q2", "Q3", "Q4"]
}
```

### write_rich_text

Write text with mixed formatting runs.

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "cell": "A1",
  "runs": [
    { "text": "Bold ", "bold": true },
    { "text": "and ", "italic": true },
    { "text": "red", "color": "#FF0000" }
  ]
}
```

### write_json_rows

Write JSON objects as spreadsheet rows. Keys become headers.

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "start_cell": "A1",
  "write_headers": true,
  "rows": [
    { "name": "Alice", "age": 30, "active": true },
    { "name": "Bob", "age": 25, "active": false }
  ]
}
```

### write_formula

Write formulas with type control.

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "cell": "A1",
  "formula": "SUM(B1:B10)",
  "formula_type": "regular",
  "cached_result": 150.0
}
```

`formula_type`: `"regular"` (default), `"array"` (CSE — cell should be a range like `"A1:C3"`), `"dynamic"` (Excel 365 spill).

## Formatting

### set_cell_format

Apply formatting to a range.

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "range": "A1:E1",
  "bold": true,
  "italic": false,
  "font_size": 12,
  "font_color": "#FFFFFF",
  "background_color": "#4472C4",
  "number_format": "#,##0.00",
  "horizontal_alignment": "center",
  "vertical_alignment": "center",
  "border_style": "thin"
}
```

Alignments: `left`, `center`, `right`, `fill`, `justify`. Borders: `thin`, `medium`, `thick`, `dashed`, `dotted`, `double`, `none`.

### merge_cells

```json
{ "workbook_id": "...", "sheet_name": "Sheet1", "range": "A1:E1" }
```

### set_dimensions

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "target": "column_width",
  "column": "A",
  "value": 20.0
}
```

Targets: `"column_width"`, `"row_height"`, `"column_range_width"` (needs `first_column` + `last_column`), `"default_row_height"`.

## Charts

### add_chart

Full-featured chart with 7 base types and extensive customization.

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "chart_type": "column",
  "series": [
    {
      "values": "Sheet1!$B$2:$B$10",
      "categories": "Sheet1!$A$2:$A$10",
      "name": "Revenue",
      "color": "#4472C4",
      "data_labels": true,
      "trendline": "linear",
      "marker": "circle",
      "line_width": 2.0,
      "dash_style": "solid",
      "error_bars": { "bar_type": "both", "value_type": "percentage", "value": 5.0 }
    }
  ],
  "cell": "E2",
  "title": "Revenue by Quarter",
  "x_axis_label": "Quarter",
  "y_axis_label": "Revenue ($)",
  "legend_position": "bottom",
  "width": 600,
  "height": 400,
  "show_data_table": true,
  "view_3d": { "rot_x": 15, "rot_y": 20, "perspective": 30 },
  "style": 26,
  "alt_text": { "title": "Revenue Chart", "description": "Quarterly revenue bar chart" },
  "y_axis_min": 0,
  "y_axis_max": 1000,
  "y_axis_format": "$#,##0",
  "plot_area_fill": "#F5F5F5"
}
```

Chart types: `bar`, `column`, `line`, `pie`, `scatter`, `area`, `doughnut`.

Trendlines: `linear`, `exponential`, `power`, `logarithmic`.

Markers: `circle`, `diamond`, `square`, `triangle`, `none`.

Dash styles: `solid`, `dash`, `dot`, `dash_dot`, `long_dash`, `long_dash_dot`.

Error bar types: `both`, `plus`, `minus`. Value types: `fixed`, `percentage`, `std_dev`, `std_error`.

### ChartEx Types

`add_waterfall_chart`, `add_funnel_chart`, `add_treemap_chart`, `add_sunburst_chart`, `add_histogram_chart`, `add_box_whisker_chart`, `add_map_chart` — see README for parameters.

### add_chart_sheet

Create a dedicated chart-only sheet.

```json
{
  "workbook_id": "...",
  "sheet_name": "Revenue Chart",
  "chart_type": "line",
  "series": [{ "values": "Data!$B$2:$B$10", "name": "Revenue" }],
  "title": "Annual Revenue"
}
```

## Pivot Tables

### add_pivot_table

```json
{
  "workbook_id": "...",
  "sheet_name": "PivotSheet",
  "cell": "A1",
  "name": "SalesPivot",
  "source_range": "'Data'!$A$1:$E$100",
  "row_fields": ["Region", "Product"],
  "column_fields": ["Quarter"],
  "value_fields": [
    { "field": "Revenue", "aggregation": "sum" },
    { "field": "Units", "aggregation": "count" }
  ],
  "filter_fields": ["Year"],
  "style": "PivotStyleMedium9",
  "layout": "tabular",
  "calculated_fields": [
    { "name": "AvgPrice", "formula": "Revenue/Units" }
  ],
  "date_groups": [
    { "field": "OrderDate", "levels": ["years", "quarters"] }
  ],
  "range_groups": [
    { "field": "Revenue", "start": 0, "end": 10000, "interval": 1000 }
  ],
  "value_formats": [
    { "field": "Revenue", "format": "$#,##0" }
  ],
  "subtotals": [
    { "field": "Region", "show": true }
  ],
  "grand_total_rows": true,
  "grand_total_cols": false,
  "show_row_headers": true,
  "show_row_stripes": true
}
```

Aggregations: `sum`, `count`, `average`, `max`, `min`, `product`, `count_nums`, `std_dev`, `var`.

Layouts: `compact`, `outline`, `tabular`.

Date group levels: `years`, `quarters`, `months`, `days`, `hours`, `minutes`, `seconds`.

## Interactive Controls

### add_slicer

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "pivot_table_name": "SalesPivot",
  "field_name": "Region",
  "cell": "H1",
  "width": 200,
  "height": 300
}
```

### add_timeline

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "pivot_table_name": "SalesPivot",
  "field_name": "OrderDate",
  "cell": "H15"
}
```

### add_form_control

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "cell": "A1",
  "control_type": "checkbox",
  "text": "Include tax",
  "cell_link": "$F$1"
}
```

Control types: `button`, `checkbox`, `dropdown`, `spinner`.

## Tables and Data Features

### add_table

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "range": "A1:E10",
  "columns": ["Name", "Q1", "Q2", "Q3", "Q4"],
  "style": "Table Style Medium 2",
  "autofilter": true,
  "totals_row": true
}
```

### add_conditional_format

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "range": "B2:B100",
  "rule": {
    "type": "color_scale_3",
    "min_color": "#FF0000",
    "mid_color": "#FFFF00",
    "max_color": "#00FF00"
  }
}
```

Rule types: `cell_value` (with operator, value, value2), `color_scale_2`, `color_scale_3`, `data_bar`, `icon_set`.

### add_data_validation

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "range": "C2:C100",
  "validation": {
    "type": "list",
    "values": ["Low", "Medium", "High"]
  },
  "input_message": { "title": "Priority", "body": "Select a priority level" },
  "error_alert": { "style": "stop", "title": "Invalid", "message": "Pick from the list" }
}
```

Validation types: `list`, `list_range`, `whole_number`, `decimal`, `date_range`, `text_length`, `custom_formula`.

## Images and Shapes

### add_image

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "cell": "A1",
  "image_path": "/path/to/logo.png",
  "width": 200,
  "height": 100
}
```

### add_shape

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "cell": "B2",
  "shape_type": "rounded_rectangle",
  "width": 200,
  "height": 100,
  "text": "Click here",
  "fill_color": "#4472C4",
  "outline_color": "#2F5496",
  "font_size": 14,
  "bold": true
}
```

Shape types: `rectangle`, `rounded_rectangle`, `ellipse`, `triangle`, `diamond`, `arrow`, `callout`, `text_box`.

## Comments and Links

### manage_comments

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "action": "add",
  "cell": "A1",
  "text": "Review this value",
  "author": "Analyst"
}
```

### add_threaded_comment

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "cell": "A1",
  "author": "Alice",
  "text": "Is this number correct?",
  "timestamp": "2025-07-25T10:00:00.000",
  "replies": [
    { "author": "Bob", "text": "Yes, verified against source." }
  ]
}
```

### add_link

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "cell": "A1",
  "link_type": "url",
  "target": "https://example.com",
  "display_text": "Example Site"
}
```

For internal links: `"link_type": "internal"`, `"target": "Sheet2!A1"`.

## Named Ranges

### manage_named_ranges

```json
{
  "workbook_id": "...",
  "action": "add_scoped",
  "name": "TaxRate",
  "formula": "Sheet1!$F$1",
  "sheet_index": 0
}
```

Actions: `add` (workbook scope), `add_scoped` (sheet scope), `update`, `remove` (needs `sheet_index` for scoped), `list`.

## Protection

### protect

Basic protection for sheets, workbooks, or unprotect ranges.

```json
{
  "workbook_id": "...",
  "target": "sheet",
  "sheet_name": "Sheet1",
  "password": "secret"
}
```

### protect_sheet_advanced

Granular protection with per-feature control.

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "password": "secret",
  "allow_format_cells": true,
  "allow_sort": true,
  "allow_insert_rows": false,
  "allow_delete_rows": false
}
```

All `allow_*` fields default to locked (not allowed). Set to `true` to permit that action.

## Page Setup and Print

### set_page_setup

```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "landscape": true,
  "paper_size": 1,
  "margins": { "top": 0.75, "bottom": 0.75, "left": 0.7, "right": 0.7 },
  "fit_to_pages": { "width": 1, "height": 0 },
  "print_area": "A1:E50",
  "header": "&LCompany&C&D&R&P of &N",
  "footer": "&CConfidential",
  "print_gridlines": true
}
```

## Document Properties

### set_doc_properties

```json
{
  "workbook_id": "...",
  "title": "Q3 Report",
  "author": "Finance Team",
  "company": "Acme Corp",
  "keywords": "quarterly, finance"
}
```

### set_custom_property

```json
{
  "workbook_id": "...",
  "name": "Department",
  "value": "Engineering",
  "value_type": "text"
}
```

Value types: `text`, `number`, `integer`, `bool`, `datetime`.

## Advanced Features

### manage_custom_xml

```json
{
  "workbook_id": "...",
  "action": "add",
  "namespace": "http://example.com/schema",
  "content": "<data><item>value</item></data>"
}
```

### add_connection

```json
{
  "workbook_id": "...",
  "connection_string": "Provider=SQLOLEDB;Data Source=server;...",
  "command": "SELECT * FROM sales"
}
```

### set_sst_threshold

Controls when the shared string table is used. Lower values use more memory but write faster.

```json
{
  "workbook_id": "...",
  "threshold": 100
}
```

## Error Handling

All tools return structured JSON responses.

### Error Categories

| Category | Description |
|---|---|
| `not_found` | Workbook ID or sheet name not found |
| `io_error` | File system or I/O error |
| `engine_unsupported` | Operation not supported (e.g., saving read-only workbook) |
| `capacity_exceeded` | Workbook store is full (max 10) |

### Error Response

```json
{
  "status": "error",
  "category": "not_found",
  "message": "Workbook 'abc-123' not found",
  "suggestion": "Check the workbook_id. Open workbooks: [\"def-456\"]"
}
```

The `suggestion` field provides actionable guidance. For `not_found` errors, it lists currently open workbook IDs.

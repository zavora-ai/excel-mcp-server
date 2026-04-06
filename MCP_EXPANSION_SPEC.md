# MCP Server — Tool Expansion Spec

Bring all 28 existing + 15 new tools to full zavora-xlsx parity.
Grouped into 4 batches for incremental implementation.

## Batch 1: Core Missing Tools (highest impact)

### 1. `set_page_setup` — Print-ready documents
```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "landscape": true,
  "paper_size": 1,
  "margins": { "top": 0.75, "bottom": 0.75, "left": 0.5, "right": 0.5 },
  "fit_to_pages": { "width": 1, "height": 0 },
  "print_scale": 100,
  "print_area": "A1:H50",
  "repeat_rows": { "first": 0, "last": 0 },
  "header": "&CMonthly Report",
  "footer": "&CPage &P of &N",
  "print_gridlines": false,
  "center_horizontally": true
}
```

### 2. `add_comment` — Cell annotations
```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "cell": "B5",
  "text": "Review this figure",
  "author": "Analyst"
}
```

### 3. `add_hyperlink` — Clickable links
```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "cell": "A1",
  "url": "https://example.com",
  "tooltip": "Click to visit"
}
```

### 4. `add_defined_name` / `list_defined_names` — Named ranges
```json
{ "workbook_id": "...", "name": "TaxRate", "formula": "Sheet1!$B$1" }
```

### 5. `set_sheet_settings` — Visibility, zoom, gridlines, tab color
```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "hidden": false,
  "zoom": 100,
  "hide_gridlines": true,
  "hide_headings": false,
  "tab_color": "#FF0000",
  "right_to_left": false
}
```

### 6. `set_active_sheet`
```json
{ "workbook_id": "...", "sheet_index": 0 }
```

## Batch 2: Data Manipulation Tools

### 7. `insert_rows` / `delete_rows`
```json
{ "workbook_id": "...", "sheet_name": "Sheet1", "at_row": 5, "count": 3 }
```

### 8. `insert_columns` / `delete_columns`
```json
{ "workbook_id": "...", "sheet_name": "Sheet1", "at_column": "C", "count": 2 }
```

### 9. `group_rows` / `group_columns` — Outline/grouping
```json
{ "workbook_id": "...", "sheet_name": "Sheet1", "start": 5, "end": 20, "level": 1 }
```

### 10. `protect_sheet` / `protect_workbook`
```json
{ "workbook_id": "...", "sheet_name": "Sheet1", "password": "secret" }
```

### 11. `autofit_columns`
```json
{ "workbook_id": "...", "sheet_name": "Sheet1" }
```

## Batch 3: Advanced Chart Features

### 12. Extend `add_chart` with full options
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
      "marker": "circle"
    }
  ],
  "title": "Sales Report",
  "x_axis_label": "Month",
  "y_axis_label": "Revenue ($)",
  "legend_position": "bottom",
  "width": 640,
  "height": 400,
  "cell": "E2",
  "pivot_source": { "pivot_table": "SalesPivot", "sheet": "Data" }
}
```

### 13. `add_pivot_table`
```json
{
  "workbook_id": "...",
  "sheet_name": "PivotSheet",
  "cell": "A1",
  "name": "SalesPivot",
  "source_range": "'Data'!$A$1:$E$100",
  "row_fields": ["Region", "Product"],
  "column_fields": ["Quarter"],
  "value_fields": [{ "field": "Revenue", "aggregation": "sum" }],
  "filter_fields": ["Category"],
  "style": "PivotStyleMedium9",
  "layout": "tabular"
}
```

## Batch 4: Read Enhancements + Extras

### 14. `read_comments`
```json
{ "workbook_id": "...", "sheet_name": "Sheet1" }
```

### 15. `write_rich_text`
```json
{
  "workbook_id": "...",
  "sheet_name": "Sheet1",
  "cell": "A1",
  "runs": [
    { "text": "Bold ", "bold": true },
    { "text": "and ", "color": "#FF0000" },
    { "text": "italic", "italic": true }
  ]
}
```

## Summary

| Batch | Tools | Count |
|-------|-------|-------|
| 1 | page_setup, comment, hyperlink, defined_names, sheet_settings, active_sheet | 7 |
| 2 | insert/delete rows/cols, group rows/cols, protect, autofit | 8 |
| 3 | enhanced charts (series/colors/labels/trendlines/pivot), pivot_table | 2 |
| 4 | read_comments, write_rich_text | 2 |
| **Total new** | | **19 tools** |
| **Existing** | | **28 tools** |
| **Grand total** | | **47 tools** |

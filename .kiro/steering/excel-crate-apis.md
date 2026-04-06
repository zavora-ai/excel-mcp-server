---
inclusion: auto
---

# Excel Crate API Reference

This steering file documents the APIs of the four Rust crates used in the Excel MCP Server project. Reference source code is available at `reference-repos/` in the workspace root.

## Dependency Versions (Cargo.toml)

```toml
rmcp = { version = "1.3.0", features = ["server", "transport-io", "macros"] }
rust_xlsxwriter = "0.94"
umya-spreadsheet = "2.3.3"
calamine = "0.34"
```

---

## 1. rmcp (v1.3.0) — MCP Server SDK

Reference: `reference-repos/rmcp/`

### Server Pattern

The rmcp SDK uses proc macros to define MCP tools. The pattern is:

```rust
use rmcp::{
    ServerHandler,
    handler::server::{router::tool::ToolRouter, wrapper::Parameters},
    model::{ServerCapabilities, ServerInfo},
    schemars, tool, tool_handler, tool_router,
    ServiceExt, transport::stdio,
};

#[derive(Debug, Clone)]
pub struct MyServer {
    tool_router: ToolRouter<Self>,
    // ... your state fields
}

#[tool_router]
impl MyServer {
    pub fn new() -> Self {
        Self {
            tool_router: Self::tool_router(),
            // ...
        }
    }

    #[tool(description = "Description of the tool")]
    fn my_tool(
        &self,
        Parameters(input): Parameters<MyInput>,
    ) -> Result<CallToolResult, McpError> {
        // implementation
        Ok(CallToolResult::success(vec![Content::text("result")]))
    }

    // Tools can also be async:
    #[tool(description = "Async tool")]
    async fn my_async_tool(&self) -> Result<CallToolResult, McpError> {
        Ok(CallToolResult::success(vec![Content::text("done")]))
    }

    // Tools can return plain String (auto-wrapped):
    #[tool(description = "Simple tool")]
    fn simple(&self, Parameters(input): Parameters<MyInput>) -> String {
        "result".to_string()
    }
}

#[tool_handler]
impl ServerHandler for MyServer {
    fn get_info(&self) -> ServerInfo {
        ServerInfo::new(ServerCapabilities::builder().enable_tools().build())
            .with_instructions("Server description".to_string())
    }
}
```

### Main Entry Point

```rust
#[tokio::main]
async fn main() -> anyhow::Result<()> {
    tracing_subscriber::fmt()
        .with_writer(std::io::stderr)
        .with_ansi(false)
        .init();

    let service = MyServer::new()
        .serve(stdio())
        .await
        .inspect_err(|e| tracing::error!("serving error: {:?}", e))?;

    service.waiting().await?;
    Ok(())
}
```

### Input Structs for Tools

Tool inputs use `Parameters<T>` wrapper. The inner type must derive `Deserialize` and `JsonSchema`:

```rust
#[derive(Debug, serde::Deserialize, schemars::JsonSchema)]
pub struct MyInput {
    /// Description shown in schema
    pub field: String,
    #[serde(default)]
    pub optional_field: Option<i32>,
}
```

### Key Imports

```rust
use rmcp::{
    ErrorData as McpError,
    ServerHandler,
    handler::server::{router::tool::ToolRouter, wrapper::Parameters},
    model::*,
    schemars,
    tool, tool_handler, tool_router,
    ServiceExt,
    transport::stdio,
};
```

---

## 2. rust_xlsxwriter (v0.94) — Write-Only Excel Engine

Reference: `reference-repos/rust_xlsxwriter/src/`

### CRITICAL: This engine is WRITE-ONLY. It cannot read cell values back.

### Workbook API (`workbook.rs`)

```rust
use rust_xlsxwriter::*;

// Create
let mut workbook = Workbook::new();

// Add worksheets
let worksheet = workbook.add_worksheet();           // returns &mut Worksheet
worksheet.set_name("MySheet")?;

// Access worksheets
let ws = workbook.worksheet_from_index(0)?;         // &mut Worksheet
let ws = workbook.worksheet_from_name("Sheet1")?;   // &mut Worksheet
let sheets = workbook.worksheets_mut();              // &mut Vec<Worksheet>

// Save
workbook.save("output.xlsx")?;
workbook.save_to_buffer()?;                          // Vec<u8>
```

### Worksheet — Writing Cells (`worksheet.rs`)

All write methods use 0-based (row: u32, col: u16) indices. Type aliases: `RowNum = u32`, `ColNum = u16`.

```rust
// Generic write (auto-detects type)
ws.write(row, col, "string")?;
ws.write(row, col, 42.5)?;
ws.write(row, col, true)?;

// Typed writes
ws.write_string(row, col, "text")?;
ws.write_number(row, col, 3.14)?;
ws.write_boolean(row, col, true)?;
ws.write_formula(row, col, "SUM(A1:A10)")?;
ws.write_blank(row, col, &Format::new())?;

// With format
ws.write_with_format(row, col, "text", &format)?;
ws.write_number_with_format(row, col, 3.14, &format)?;
ws.write_string_with_format(row, col, "text", &format)?;

// DateTime
use rust_xlsxwriter::ExcelDateTime;
let dt = ExcelDateTime::parse_from_str("2024-01-15")?;
ws.write_datetime_with_format(row, col, &dt, &date_format)?;

// Row/Column batch writes
ws.write_row(row, col, ["a", "b", "c"])?;
ws.write_column(row, col, [1, 2, 3])?;

// Formulas (no "=" prefix needed — just the formula text)
ws.write_formula(row, col, "SUM(A1:A10)")?;
```

### Format API (`format.rs`)

Format uses a builder pattern with `self`-consuming methods:

```rust
let format = Format::new()
    .set_bold()
    .set_italic()
    .set_underline(FormatUnderline::Single)
    .set_font_size(14.0)
    .set_font_color(Color::Red)
    .set_font_name("Arial")
    .set_background_color(Color::Yellow)
    .set_num_format("#,##0.00")
    .set_align(FormatAlign::Center)
    .set_align(FormatAlign::VerticalCenter)  // can chain both
    .set_border(FormatBorder::Thin)
    .set_border_color(Color::Black);

// Apply to a cell
ws.set_cell_format(row, col, &format)?;

// Apply to a range
ws.set_range_format(first_row, first_col, last_row, last_col, &format)?;
```

Key enums:
- `FormatAlign`: Left, Center, Right, Fill, Justify, Top, VerticalCenter, Bottom, VerticalJustify
- `FormatBorder`: None, Thin, Medium, Thick, Dashed, Dotted, Double, Hair, etc.
- `FormatUnderline`: None, Single, Double, SingleAccounting, DoubleAccounting
- `Color`: Named colors (Red, Blue, etc.) or `Color::from("hex_string")` or `Color::from(0xRRGGBB_u32)`

### Colors

```rust
// From hex string (with or without #)
Color::from("#FF0000")
Color::from("FF0000")

// From u32
Color::from(0xFF0000_u32)

// Named
Color::Red, Color::Blue, Color::Green, etc.
```

### Merge Cells

```rust
ws.merge_range(first_row, first_col, last_row, last_col, "text", &format)?;
```

### Charts (`chart.rs`)

```rust
use rust_xlsxwriter::{Chart, ChartType};

let mut chart = Chart::new(ChartType::Column);
// or: Chart::new_bar(), Chart::new_line(), Chart::new_pie(), etc.

// Add data series
chart.add_series()
    .set_values("Sheet1!$B$1:$B$5")
    .set_categories("Sheet1!$A$1:$A$5")
    .set_name("Sales");

// Configure
chart.title().set_name("My Chart");
chart.x_axis().set_name("Category");
chart.y_axis().set_name("Value");
chart.legend().set_position(ChartLegendPosition::Bottom);
chart.legend().set_hidden();  // to hide legend
chart.set_width(640);
chart.set_height(480);

// Insert into worksheet
ws.insert_chart(row, col, &chart)?;
```

ChartType enum: `Bar, Column, Line, Pie, Scatter, Area, Doughnut, Radar, Stock`

### Images (`image.rs`)

```rust
use rust_xlsxwriter::Image;

let mut image = Image::new("path/to/image.png")?;
image = image.set_width(200);
image = image.set_height(150);

ws.insert_image(row, col, &image)?;
```

### Tables (`table.rs`)

```rust
use rust_xlsxwriter::{Table, TableColumn, TableStyle};

let columns = vec![
    TableColumn::new().set_header("Name"),
    TableColumn::new().set_header("Value"),
];

let table = Table::new()
    .set_columns(&columns)
    .set_style(TableStyle::Medium2)
    .set_total_row(true)
    .set_autofilter(true);

ws.add_table(first_row, first_col, last_row, last_col, &table)?;
```

### Conditional Formatting (`conditional_format.rs`)

```rust
use rust_xlsxwriter::*;

// Cell value rule
let cf = ConditionalFormatCell::new()
    .set_rule(ConditionalFormatCellRule::GreaterThan(50.0))
    .set_format(Format::new().set_font_color(Color::Red));
ws.add_conditional_format(first_row, first_col, last_row, last_col, &cf)?;

// 2-color scale
let cf = ConditionalFormat2ColorScale::new()
    .set_minimum_color(Color::from("#FFFFFF"))
    .set_maximum_color(Color::from("#FF0000"));
ws.add_conditional_format(first_row, first_col, last_row, last_col, &cf)?;

// 3-color scale
let cf = ConditionalFormat3ColorScale::new()
    .set_minimum_color(Color::from("#FF0000"))
    .set_midpoint_color(Color::from("#FFFF00"))
    .set_maximum_color(Color::from("#00FF00"));

// Data bar
let cf = ConditionalFormatDataBar::new()
    .set_fill_color(Color::from("#4472C4"));

// Icon set
let cf = ConditionalFormatIconSet::new()
    .set_icon_type(ConditionalFormatIconType::ThreeArrows);
```

ConditionalFormatCellRule variants: `GreaterThan(f64)`, `LessThan(f64)`, `Between(f64, f64)`, `EqualTo(f64)`, `NotEqualTo(f64)`, `GreaterThanOrEqualTo(f64)`, `LessThanOrEqualTo(f64)`

### Data Validation (`data_validation.rs`)

```rust
use rust_xlsxwriter::*;

let mut dv = DataValidation::new();

// List validation
dv = dv.allow_list_strings(&["Yes", "No", "Maybe"])?;

// List from formula/range
dv = dv.allow_list_formula(Formula::new("Sheet2!$A$1:$A$10"));

// Number validation
dv = dv.allow_whole_number(DataValidationRule::Between(1, 100));
dv = dv.allow_decimal_number(DataValidationRule::GreaterThanOrEqualTo(0.0));

// Text length
dv = dv.allow_text_length(DataValidationRule::Between(1_u32, 50_u32));

// Date validation
let min = ExcelDateTime::parse_from_str("2024-01-01")?;
let max = ExcelDateTime::parse_from_str("2024-12-31")?;
dv = dv.allow_date(DataValidationRule::Between(min, max));

// Custom formula
dv = dv.allow_custom(Formula::new("AND(A1>0,A1<100)"));

// Messages
dv = dv.set_input_title("Enter value")?;
dv = dv.set_input_message("Please enter a valid value")?;
dv = dv.set_error_style(DataValidationErrorStyle::Stop);
dv = dv.set_error_title("Invalid")?;
dv = dv.set_error_message("Value out of range")?;

ws.add_data_validation(first_row, first_col, last_row, last_col, &dv)?;
```

### Sparklines (`sparkline.rs`)

```rust
use rust_xlsxwriter::{Sparkline, SparklineType};

let sparkline = Sparkline::new()
    .set_range("Sheet1!A1:F1")
    .set_type(SparklineType::Line);  // Line, Column, WinLose

ws.add_sparkline(row, col, &sparkline)?;
```

### Layout

```rust
// Column width (in character units)
ws.set_column_width(col, 20.0)?;

// Row height (in points)
ws.set_row_height(row, 30.0)?;

// Freeze panes (rows above and cols left of this cell are frozen)
ws.set_freeze_panes(row, col)?;
```

---

## 3. umya-spreadsheet (v3.0) — Read/Write Excel Engine

Reference: `reference-repos/umya-spreadsheet/src/`

### Opening Files

```rust
use umya_spreadsheet::*;

// Standard read
let path = std::path::Path::new("file.xlsx");
let mut book = reader::xlsx::read(path).unwrap();

// Lazy read (delays worksheet loading — better for large files)
let mut book = reader::xlsx::lazy_read(path).unwrap();

// New file
let mut book = new_file();  // creates workbook with "Sheet1"
```

### Saving Files

```rust
let path = std::path::Path::new("output.xlsx");
let _ = writer::xlsx::write(&book, path);
```

### Sheet Access

```rust
// By index (0-based)
let sheet = book.get_sheet(&0).unwrap();           // &Worksheet
let sheet = book.get_sheet_mut(&0).unwrap();       // &mut Worksheet

// By name
let sheet = book.get_sheet_by_name("Sheet1").unwrap();
let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();

// Sheet count
let count = book.get_sheet_count();

// All sheets
let sheets = book.get_sheet_collection();          // &[Worksheet]

// Sheet names
for sheet in book.get_sheet_collection() {
    println!("{}", sheet.get_name());
}
```

### Sheet Management

```rust
// Add new sheet
let sheet = book.new_sheet("Sheet2").unwrap();      // &mut Worksheet

// Rename
book.set_sheet_name(0, "NewName").unwrap();

// Remove
book.remove_sheet(0).unwrap();
book.remove_sheet_by_name("Sheet2").unwrap();
```

### Cell Access — Reading

Cells use 1-based (col, row) tuples or A1 string notation:

```rust
let sheet = book.get_sheet(&0).unwrap();

// By string address
let value = sheet.get_value("A1");                  // String
let value = sheet.get_value((1, 1));                // (col, row) 1-based

// Get cell object
let cell = sheet.get_cell("A1");                    // Option<&Cell>
let cell = sheet.get_cell((1, 1));

// Cell properties
if let Some(cell) = cell {
    let value = cell.get_value();                   // Cow<str>
    let data_type = cell.get_data_type();           // &str: "s", "n", "b", "e", etc.
    let is_formula = cell.is_formula();
    let formula = cell.get_formula();               // &str
    let raw = cell.get_raw_value();                 // &CellRawValue
}

// Formatted value
let formatted = sheet.get_formatted_value("A1");

// Dimensions
let (max_col, max_row) = sheet.get_highest_column_and_row();
let dim_string = sheet.calculate_worksheet_dimension();  // e.g. "A1:F100"
```

### Cell Access — Writing

```rust
let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();

// String value
sheet.get_cell_mut("A1").set_value("Hello");

// Number
sheet.get_cell_mut("B1").set_value_number(42.5);

// Boolean
sheet.get_cell_mut("C1").set_value_bool(true);

// Formula
sheet.get_cell_mut("D1").set_formula("SUM(A1:C1)");

// Using tuple coordinates (col, row) — both 1-based
sheet.get_cell_mut((1, 1)).set_value("Hello");
sheet.get_cell_mut((2, 1)).set_value_number(42.5);
```

**IMPORTANT**: umya-spreadsheet uses (column, row) order in tuples, both 1-based. Column 1 = A, Row 1 = 1.

### Styling

```rust
let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();

// Get style for a cell
let style = sheet.get_style_mut("A1");

// Font
style.get_font_mut().set_bold(true);
style.get_font_mut().set_italic(true);
style.get_font_mut().set_size(14.0);
style.get_font_mut().set_name("Arial");
style.get_font_mut().get_color_mut().set_argb("FF0000");

// Background color
style.set_background_color("FFFF00");

// Borders
style.get_borders_mut().get_bottom_mut().set_border_style(Border::BORDER_THIN);
style.get_borders_mut().get_top_mut().set_border_style(Border::BORDER_THIN);
style.get_borders_mut().get_left_mut().set_border_style(Border::BORDER_THIN);
style.get_borders_mut().get_right_mut().set_border_style(Border::BORDER_THIN);

// Number format
style.get_numbering_format_mut().set_format_code("#,##0.00");

// Alignment
use umya_spreadsheet::HorizontalAlignmentValues;
style.get_alignment_mut().set_horizontal(HorizontalAlignmentValues::Center);

// Apply style to a range
let style_obj = Style::default();
sheet.set_style_by_range("A1:D10", &style_obj);
```

### Merge Cells

```rust
sheet.add_merge_cells("A1:D1");
```

### Charts

```rust
use umya_spreadsheet::*;

let mut chart = Chart::default();
// umya charts are complex — set chart type, data references, etc.
// chart.set_chart_type(ChartType::...);
sheet.add_chart(chart);
```

### Images

```rust
use umya_spreadsheet::Image;

let mut image = Image::default();
// Configure image with file path, position, etc.
sheet.add_image(image);
```

### Column Width / Row Height

```rust
// Column width
sheet.get_column_dimension_mut("A").set_width(20.0);

// Row height (1-based row number)
sheet.get_row_dimension_mut(&1).set_height(30.0);
```

### Freeze Panes

```rust
// Via SheetView pane
let pane = Pane::default();
// Configure pane position
sheet.get_sheet_views_mut()...;
```

---

## 4. calamine (v0.34) — Read-Only Excel Engine

Reference: `reference-repos/calamine/src/`

### CRITICAL: This engine is READ-ONLY. No write operations.

### Opening Files

```rust
use calamine::{Reader, Xlsx, open_workbook, Data, DataType};
use std::io::BufReader;
use std::fs::File;

// Open xlsx
let mut workbook: Xlsx<BufReader<File>> = open_workbook("file.xlsx").unwrap();

// Sheet names
let names = workbook.sheet_names();  // Vec<String>
```

### Reading Sheet Data

```rust
use calamine::{Reader, Range, Data};

// Get a worksheet range
let range: Range<Data> = workbook.worksheet_range("Sheet1").unwrap();

// Dimensions
let (height, width) = range.get_size();  // (rows, cols)
let start = range.start();               // Option<(u32, u32)> — (row, col) 0-based
let end = range.end();                   // Option<(u32, u32)>

// Iterate rows
for row in range.rows() {
    for cell in row {
        // cell is &Data
    }
}

// Access specific cell (relative position)
let val = range.get((0, 0));  // Option<&Data> — (row, col) relative to range start

// Used cells iterator
for cell in range.used_cells() {
    let (row, col, value) = cell;  // (usize, usize, &Data)
}
```

### Data Enum

The `Data` enum represents cell values:

```rust
pub enum Data {
    Int(i64),
    Float(f64),
    String(String),
    Bool(bool),
    DateTime(ExcelDateTime),
    DateTimeIso(String),
    DurationIso(String),
    Error(CellErrorType),
    Empty,
}
```

Use the `DataType` trait methods:
```rust
use calamine::DataType;

let cell: &Data = ...;
cell.is_empty();
cell.is_int();
cell.is_float();
cell.is_bool();
cell.is_string();
cell.get_int()    -> Option<i64>
cell.get_float()  -> Option<f64>
cell.get_bool()   -> Option<bool>
cell.get_string() -> Option<&str>
cell.as_string()  -> Option<String>  // converts any type to string
cell.as_f64()     -> Option<f64>     // converts numeric types
```

### Sub-range

```rust
// Extract a sub-range
let sub = range.range((0, 0), (10, 5));  // Range<Data>
```

### Sheet Metadata

```rust
let metadata = workbook.metadata();
for sheet in &metadata.sheets {
    println!("Name: {}, Visible: {:?}", sheet.name, sheet.visible);
}
```

---

## Engine Capability Matrix

| Operation | rust_xlsxwriter | umya-spreadsheet | calamine |
|---|---|---|---|
| Create workbook | ✅ | ✅ | ❌ |
| Open existing | ❌ | ✅ (edit) | ✅ (read-only) |
| Read cells | ❌ | ✅ | ✅ |
| Write cells | ✅ | ✅ | ❌ |
| Cell formatting | ✅ | ✅ | ❌ |
| Merge cells | ✅ | ✅ | ❌ |
| Charts | ✅ (full) | ✅ (basic) | ❌ |
| Images | ✅ | ✅ | ❌ |
| Tables | ✅ | ❌ | ❌ |
| Conditional formatting | ✅ | ❌ | ❌ |
| Data validation | ✅ | ❌ | ❌ |
| Sparklines | ✅ | ❌ | ❌ |
| Freeze panes | ✅ | ✅ | ❌ |
| Save | ✅ | ✅ | ❌ |
| Search cells | ❌ | ✅ (manual) | ✅ (manual) |
| CSV export | ❌ | ✅ | ✅ (manual) |
| Sheet management | ✅ | ✅ | ❌ |

## Coordinate Systems

- **rust_xlsxwriter**: 0-based `(row: u32, col: u16)` — row 0 = Excel row 1, col 0 = column A
- **umya-spreadsheet**: 1-based `(col: u32, row: u32)` tuples OR A1 string notation — NOTE: tuple order is (col, row)
- **calamine**: 0-based `(row: u32, col: u32)` in Range positions — row 0 = first row of range

## Common Pitfalls

1. **umya tuple order is (col, row)** not (row, col) — easy to mix up
2. **rust_xlsxwriter Format consumes self** — `Format::new().set_bold()` returns a new Format, not &mut
3. **calamine Range is relative** — `range.get((0,0))` gets the first cell of the range, not A1
4. **rust_xlsxwriter formulas don't need "=" prefix** — just pass the formula text
5. **umya-spreadsheet lazy_read** delays worksheet loading — call `read_sheet_collection()` or access sheets to trigger loading
6. **calamine Data::DateTime** contains an `ExcelDateTime` with methods `as_datetime()` and `as_f64()`

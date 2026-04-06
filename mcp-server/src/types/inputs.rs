use schemars::JsonSchema;
use serde::Deserialize;

use super::enums::*;

// ── Serde default helpers ──────────────────────────────────────────

fn default_chart_width() -> u32 {
    480
}

fn default_chart_height() -> u32 {
    288
}

fn default_comma() -> String {
    ",".to_string()
}

fn default_true() -> bool {
    true
}

// ── Workbook lifecycle inputs ──────────────────────────────────────

/// Input for creating a new empty workbook
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct CreateWorkbookInput {}

/// Input for opening an existing Excel file
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct OpenWorkbookInput {
    /// Absolute or relative path to the Excel file (xlsx, xlsm, xls, ods)
    pub file_path: String,
    /// If true, opens in read-only mode using the fast calamine engine.
    /// If false, opens in edit mode using umya-spreadsheet. Default: false
    #[serde(default)]
    pub read_only: bool,
}

/// Input for saving a workbook to disk
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SaveWorkbookInput {
    /// The workbook handle returned by create_workbook or open_workbook
    pub workbook_id: String,
    /// Destination file path (must end in .xlsx)
    pub file_path: String,
}

/// Input for closing a workbook and freeing memory
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct CloseWorkbookInput {
    /// The workbook handle to close
    pub workbook_id: String,
}

// ── Read inputs ────────────────────────────────────────────────────

/// Input for reading sheet data with optional range and pagination
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ReadSheetInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the sheet to read
    pub sheet_name: String,
    /// Optional range in A1:B2 notation. If omitted, reads the entire used range
    pub range: Option<String>,
    /// Continuation token from a previous paginated response
    pub continuation_token: Option<String>,
}

/// Input for reading a single cell
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ReadCellInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the sheet containing the cell
    pub sheet_name: String,
    /// Cell reference in A1 notation (e.g., "C5")
    pub cell: String,
}

// ── Write inputs ───────────────────────────────────────────────────

/// A single cell write operation
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct CellWrite {
    /// Cell reference in A1 notation
    pub cell: String,
    /// Value to write. Strings starting with "=" are written as formulas.
    /// Numbers, booleans, and ISO 8601 dates are auto-detected.
    pub value: serde_json::Value,
}

/// Input for batch cell writing
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct WriteCellsInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the target sheet
    pub sheet_name: String,
    /// Array of cell writes to apply
    pub cells: Vec<CellWrite>,
}

/// Input for writing a row of values starting from a cell
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct WriteRowInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the target sheet
    pub sheet_name: String,
    /// Starting cell reference (values fill rightward)
    pub start_cell: String,
    /// Values to write in consecutive columns
    pub values: Vec<serde_json::Value>,
}

/// Input for writing a column of values starting from a cell
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct WriteColumnInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the target sheet
    pub sheet_name: String,
    /// Starting cell reference (values fill downward)
    pub start_cell: String,
    /// Values to write in consecutive rows
    pub values: Vec<serde_json::Value>,
}

// ── Formatting inputs ──────────────────────────────────────────────

/// Input for applying formatting to a range of cells
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetCellFormatInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the target sheet
    pub sheet_name: String,
    /// Range in A1:B2 notation
    pub range: String,
    /// Whether to apply bold font weight
    #[serde(default)]
    pub bold: Option<bool>,
    /// Whether to apply italic font style
    #[serde(default)]
    pub italic: Option<bool>,
    /// Whether to apply underline
    #[serde(default)]
    pub underline: Option<bool>,
    /// Font size in points
    #[serde(default)]
    pub font_size: Option<f64>,
    /// Hex color string for font, e.g., "#FF0000"
    #[serde(default)]
    pub font_color: Option<String>,
    /// Hex color string for cell background
    #[serde(default)]
    pub background_color: Option<String>,
    /// Excel number format string, e.g., "#,##0.00"
    #[serde(default)]
    pub number_format: Option<String>,
    /// Horizontal text alignment
    #[serde(default)]
    pub horizontal_alignment: Option<HorizontalAlignment>,
    /// Vertical text alignment
    #[serde(default)]
    pub vertical_alignment: Option<VerticalAlignment>,
    /// Border style for all edges
    #[serde(default)]
    pub border_style: Option<BorderStyle>,
}

/// Input for merging a range of cells
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct MergeCellsInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the target sheet
    pub sheet_name: String,
    /// Range to merge in A1:B2 notation
    pub range: String,
}

// ── Chart, image, table inputs ─────────────────────────────────────

/// Input for adding a chart to a worksheet
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddChartInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the target sheet
    pub sheet_name: String,
    /// Type of chart to create
    pub chart_type: ChartType,
    /// Data range in A1:B2 notation
    pub data_range: String,
    /// Optional chart title
    #[serde(default)]
    pub title: Option<String>,
    /// Optional X-axis label
    #[serde(default)]
    pub x_axis_label: Option<String>,
    /// Optional Y-axis label
    #[serde(default)]
    pub y_axis_label: Option<String>,
    /// Position of the chart legend
    #[serde(default)]
    pub legend_position: Option<LegendPosition>,
    /// Chart width in pixels. Default: 480
    #[serde(default = "default_chart_width")]
    pub width: u32,
    /// Chart height in pixels. Default: 288
    #[serde(default = "default_chart_height")]
    pub height: u32,
}

/// Input for embedding an image in a worksheet
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddImageInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the target sheet
    pub sheet_name: String,
    /// Cell where the image top-left corner is anchored
    pub cell: String,
    /// Path to a PNG or JPEG image file
    pub image_path: String,
    /// Optional width in pixels to scale the image
    #[serde(default)]
    pub width: Option<u32>,
    /// Optional height in pixels to scale the image
    #[serde(default)]
    pub height: Option<u32>,
}

/// Input for creating an Excel Table
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddTableInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the target sheet
    pub sheet_name: String,
    /// Range for the table in A1:B2 notation (includes header row)
    pub range: String,
    /// Column header names
    pub columns: Vec<String>,
    /// Optional Excel table style name (e.g., "Table Style Medium 2")
    #[serde(default)]
    pub style: Option<String>,
    /// Whether to show a totals row. Default: false
    #[serde(default)]
    pub totals_row: bool,
    /// Whether to enable autofilter. Default: true
    #[serde(default = "default_true")]
    pub autofilter: bool,
}

// ── Conditional formatting, validation, sparkline inputs ───────────

/// Input for adding a conditional formatting rule
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddConditionalFormatInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the target sheet
    pub sheet_name: String,
    /// Range to apply the rule to in A1:B2 notation
    pub range: String,
    /// The conditional formatting rule definition
    pub rule: ConditionalFormatRule,
    /// Formatting to apply when condition is met
    #[serde(default)]
    pub format: Option<ConditionalFormatStyle>,
}

/// Input for adding data validation to a range
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddDataValidationInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the target sheet
    pub sheet_name: String,
    /// Range to apply validation to in A1:B2 notation
    pub range: String,
    /// The validation rule definition
    pub validation: ValidationRule,
    /// Optional input message shown when the cell is selected
    #[serde(default)]
    pub input_message: Option<ValidationMessage>,
    /// Optional error alert shown when invalid data is entered
    #[serde(default)]
    pub error_alert: Option<ValidationAlert>,
}

// ── Layout inputs ──────────────────────────────────────────────────

/// Input for setting the width of a column
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetColumnWidthInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the target sheet
    pub sheet_name: String,
    /// Column identifier in letter notation (e.g., "A", "BC")
    pub column: String,
    /// Width in character units
    pub width: f64,
}

/// Input for setting the height of a row
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetRowHeightInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the target sheet
    pub sheet_name: String,
    /// 1-based row number
    pub row: u32,
    /// Height in points
    pub height: f64,
}

/// Input for freezing panes at a cell position
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct FreezePanesInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the target sheet
    pub sheet_name: String,
    /// Cell reference — rows above and columns left of this cell are frozen
    pub cell: String,
}

/// Input for adding a sparkline to a cell
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddSparklineInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the target sheet
    pub sheet_name: String,
    /// Cell where the sparkline is placed
    pub target_cell: String,
    /// Data range for the sparkline values
    pub data_range: String,
    /// Type of sparkline to create
    pub sparkline_type: SparklineType,
}

// ── Search and export inputs ───────────────────────────────────────

/// Input for searching cell values across sheets
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SearchCellsInput {
    /// The workbook handle
    pub workbook_id: String,
    /// If omitted, searches all sheets
    #[serde(default)]
    pub sheet_name: Option<String>,
    /// The value or substring to search for
    pub query: String,
    /// Search mode: exact or substring. Default: substring
    #[serde(default)]
    pub match_mode: MatchMode,
}

/// Input for exporting a sheet as CSV
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SheetToCsvInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the sheet to export
    pub sheet_name: String,
    /// Delimiter character. Default: ","
    #[serde(default = "default_comma")]
    pub delimiter: String,
}

// ── Sheet management inputs ────────────────────────────────────────

/// Input for adding a new worksheet
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddSheetInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name for the new sheet
    pub sheet_name: String,
}

/// Input for renaming an existing worksheet
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct RenameSheetInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Current name of the sheet to rename
    pub current_name: String,
    /// New name for the sheet
    pub new_name: String,
}

/// Input for deleting a worksheet
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct DeleteSheetInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the sheet to delete
    pub sheet_name: String,
}

/// Input for listing all sheets in a workbook
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ListSheetsInput {
    /// The workbook handle
    pub workbook_id: String,
}

/// Input for getting the dimensions of a sheet
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct GetSheetDimensionsInput {
    /// The workbook handle
    pub workbook_id: String,
    /// Name of the sheet to measure
    pub sheet_name: String,
}

/// Input for describing a workbook's structure and sample data
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct DescribeWorkbookInput {
    /// The workbook handle
    pub workbook_id: String,
}

// ── New tools: Batch 1–4 ───────────────────────────────────────────

/// Margins for page setup (inches)
#[derive(Deserialize, JsonSchema)]
pub struct PageMargins {
    #[serde(default)] pub top: Option<f64>,
    #[serde(default)] pub bottom: Option<f64>,
    #[serde(default)] pub left: Option<f64>,
    #[serde(default)] pub right: Option<f64>,
}

/// Fit-to-page settings
#[derive(Deserialize, JsonSchema)]
pub struct FitToPages {
    pub width: u16,
    pub height: u16,
}

/// Repeat rows for printing
#[derive(Deserialize, JsonSchema)]
pub struct RepeatRows {
    /// First row (0-based)
    pub first: u32,
    /// Last row (0-based)
    pub last: u32,
}

/// Input for configuring page setup, headers, footers, print options
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetPageSetupInput {
    pub workbook_id: String,
    pub sheet_name: String,
    #[serde(default)] pub landscape: Option<bool>,
    #[serde(default)] pub paper_size: Option<u8>,
    #[serde(default)] pub margins: Option<PageMargins>,
    #[serde(default)] pub fit_to_pages: Option<FitToPages>,
    #[serde(default)] pub print_scale: Option<u16>,
    /// Print area in A1:B2 notation
    #[serde(default)] pub print_area: Option<String>,
    #[serde(default)] pub repeat_rows: Option<RepeatRows>,
    /// Excel header string (supports &L, &C, &R, &P, &N, &D codes)
    #[serde(default)] pub header: Option<String>,
    /// Excel footer string
    #[serde(default)] pub footer: Option<String>,
    #[serde(default)] pub print_gridlines: Option<bool>,
    #[serde(default)] pub center_horizontally: Option<bool>,
    #[serde(default)] pub center_vertically: Option<bool>,
}

/// Input for adding a comment/note to a cell
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddCommentInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub cell: String,
    pub text: String,
    #[serde(default)] pub author: Option<String>,
}

/// Input for adding a hyperlink to a cell
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddHyperlinkInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub cell: String,
    pub url: String,
    #[serde(default)] pub tooltip: Option<String>,
    /// Display text (if different from URL)
    #[serde(default)] pub display_text: Option<String>,
}

/// Input for adding a defined name (named range)
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddDefinedNameInput {
    pub workbook_id: String,
    pub name: String,
    /// Formula or range reference, e.g. "Sheet1!$A$1:$B$10"
    pub formula: String,
}

/// Input for listing defined names
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ListDefinedNamesInput {
    pub workbook_id: String,
}

/// Input for sheet display settings
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetSheetSettingsInput {
    pub workbook_id: String,
    pub sheet_name: String,
    #[serde(default)] pub hidden: Option<bool>,
    #[serde(default)] pub very_hidden: Option<bool>,
    #[serde(default)] pub zoom: Option<u16>,
    #[serde(default)] pub hide_gridlines: Option<bool>,
    #[serde(default)] pub hide_headings: Option<bool>,
    #[serde(default)] pub tab_color: Option<String>,
    #[serde(default)] pub right_to_left: Option<bool>,
}

/// Input for setting the active (visible) sheet
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetActiveSheetInput {
    pub workbook_id: String,
    /// 0-based sheet index
    pub sheet_index: usize,
}

/// Input for inserting/deleting rows
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct InsertDeleteRowsInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// 1-based row number where insertion/deletion starts
    pub at_row: u32,
    /// Number of rows to insert or delete
    pub count: u32,
}

/// Input for inserting/deleting columns
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct InsertDeleteColumnsInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// Column letter where insertion/deletion starts (e.g. "C")
    pub at_column: String,
    pub count: u16,
}

/// Input for grouping rows or columns (outline)
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct GroupRowsInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// First row (1-based)
    pub start: u32,
    /// Last row (1-based)
    pub end: u32,
    /// Outline level (1-7). Default: 1
    #[serde(default = "default_level")]
    pub level: u8,
}

/// Input for grouping columns
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct GroupColumnsInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// First column letter
    pub start: String,
    /// Last column letter
    pub end: String,
    #[serde(default = "default_level")]
    pub level: u8,
}

fn default_level() -> u8 { 1 }

/// Input for protecting a sheet
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ProtectSheetInput {
    pub workbook_id: String,
    pub sheet_name: String,
    #[serde(default)] pub password: Option<String>,
}

/// Input for protecting a workbook
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ProtectWorkbookInput {
    pub workbook_id: String,
    #[serde(default)] pub password: Option<String>,
}

/// Input for autofitting column widths
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AutofitColumnsInput {
    pub workbook_id: String,
    pub sheet_name: String,
}

/// A single chart series definition
#[derive(Deserialize, JsonSchema)]
pub struct ChartSeriesInput {
    /// Data range for values (e.g. "Sheet1!$B$2:$B$10")
    pub values: String,
    /// Data range for categories/labels
    #[serde(default)] pub categories: Option<String>,
    /// Series name
    #[serde(default)] pub name: Option<String>,
    /// Hex color for the series
    #[serde(default)] pub color: Option<String>,
    /// Show data labels on this series
    #[serde(default)] pub data_labels: Option<bool>,
    /// Trendline type: linear, exponential, polynomial, power, logarithmic, moving_average
    #[serde(default)] pub trendline: Option<String>,
    /// Marker type: circle, diamond, square, triangle, none
    #[serde(default)] pub marker: Option<String>,
    /// Use secondary Y axis
    #[serde(default)] pub secondary_axis: Option<bool>,
}

/// Pivot chart source
#[derive(Deserialize, JsonSchema)]
pub struct PivotChartSourceInput {
    pub pivot_table: String,
    pub sheet: String,
}

/// Enhanced chart input with full series control
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddChartEnhancedInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub chart_type: ChartType,
    /// Individual series definitions (preferred over data_range)
    #[serde(default)] pub series: Vec<ChartSeriesInput>,
    /// Simple data range (used if series is empty)
    #[serde(default)] pub data_range: Option<String>,
    /// Cell where chart top-left is placed (e.g. "E2"). Default: "A1"
    #[serde(default)] pub cell: Option<String>,
    #[serde(default)] pub title: Option<String>,
    #[serde(default)] pub x_axis_label: Option<String>,
    #[serde(default)] pub y_axis_label: Option<String>,
    #[serde(default)] pub legend_position: Option<LegendPosition>,
    #[serde(default = "default_chart_width")] pub width: u32,
    #[serde(default = "default_chart_height")] pub height: u32,
    /// Link chart to a pivot table
    #[serde(default)] pub pivot_source: Option<PivotChartSourceInput>,
}

/// Pivot table value field
#[derive(Deserialize, JsonSchema)]
pub struct PivotValueFieldInput {
    pub field: String,
    /// Aggregation: sum, count, average, max, min, product, count_nums, std_dev, var
    #[serde(default = "default_sum")]
    pub aggregation: String,
}

fn default_sum() -> String { "sum".to_string() }

/// Input for creating a pivot table
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddPivotTableInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// Cell where pivot table starts (e.g. "A1")
    #[serde(default)] pub cell: Option<String>,
    pub name: String,
    /// Source range including sheet (e.g. "'Data'!$A$1:$E$100")
    pub source_range: String,
    #[serde(default)] pub row_fields: Vec<String>,
    #[serde(default)] pub column_fields: Vec<String>,
    pub value_fields: Vec<PivotValueFieldInput>,
    #[serde(default)] pub filter_fields: Vec<String>,
    #[serde(default)] pub style: Option<String>,
    /// Layout: compact, outline, tabular. Default: compact
    #[serde(default)] pub layout: Option<String>,
}

/// Input for reading comments from a sheet
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ReadCommentsInput {
    pub workbook_id: String,
    pub sheet_name: String,
}

/// A rich text run
#[derive(Deserialize, JsonSchema)]
pub struct RichTextRunInput {
    pub text: String,
    #[serde(default)] pub bold: Option<bool>,
    #[serde(default)] pub italic: Option<bool>,
    #[serde(default)] pub color: Option<String>,
    #[serde(default)] pub font_size: Option<f64>,
}

/// Input for writing rich text to a cell
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct WriteRichTextInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub cell: String,
    pub runs: Vec<RichTextRunInput>,
}

// ── Batch 5–8: Remaining 22 tools ──────────────────────────────────

/// Input for setting column/row format
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetColumnFormatInput {
    pub workbook_id: String, pub sheet_name: String,
    pub column: String,
    #[serde(default)] pub bold: Option<bool>, #[serde(default)] pub italic: Option<bool>,
    #[serde(default)] pub font_size: Option<f64>, #[serde(default)] pub font_color: Option<String>,
    #[serde(default)] pub background_color: Option<String>, #[serde(default)] pub number_format: Option<String>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetRowFormatInput {
    pub workbook_id: String, pub sheet_name: String,
    /// 1-based row number
    pub row: u32,
    #[serde(default)] pub bold: Option<bool>, #[serde(default)] pub italic: Option<bool>,
    #[serde(default)] pub font_size: Option<f64>, #[serde(default)] pub font_color: Option<String>,
    #[serde(default)] pub background_color: Option<String>, #[serde(default)] pub number_format: Option<String>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetColumnHiddenInput {
    pub workbook_id: String, pub sheet_name: String, pub column: String, #[serde(default)] pub hidden: bool,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetRowHiddenInput {
    pub workbook_id: String, pub sheet_name: String, pub row: u32, #[serde(default)] pub hidden: bool,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetColumnRangeWidthInput {
    pub workbook_id: String, pub sheet_name: String,
    pub first_column: String, pub last_column: String, pub width: f64,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetDefaultRowHeightInput {
    pub workbook_id: String, pub sheet_name: String, pub height: f64,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetSelectionInput {
    pub workbook_id: String, pub sheet_name: String, pub cell: String,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetAutofilterInput {
    pub workbook_id: String, pub sheet_name: String,
    /// Range in A1:B2 notation
    pub range: String,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct FilterColumnInput {
    pub workbook_id: String, pub sheet_name: String,
    pub column: String, pub values: Vec<String>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct IgnoreErrorInput {
    pub workbook_id: String, pub sheet_name: String,
    /// Error type: "number_stored_as_text", "formula_range", etc.
    pub error_type: String,
    /// Range in A1:B2 notation
    pub range: String,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetPageBreaksInput {
    pub workbook_id: String, pub sheet_name: String,
    #[serde(default)] pub row_breaks: Vec<u32>,
    #[serde(default)] pub col_breaks: Vec<u16>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct UnprotectRangeInput {
    pub workbook_id: String, pub sheet_name: String,
    pub range: String, pub title: String,
    #[serde(default)] pub password: Option<String>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct WriteFormulaInput {
    pub workbook_id: String, pub sheet_name: String, pub cell: String,
    pub formula: String,
    /// Optional cached numeric result
    #[serde(default)] pub cached_result: Option<f64>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct WriteArrayFormulaInput {
    pub workbook_id: String, pub sheet_name: String,
    /// Range the array formula spans (e.g. "A1:C3")
    pub range: String,
    pub formula: String,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct WriteDynamicFormulaInput {
    pub workbook_id: String, pub sheet_name: String, pub cell: String,
    pub formula: String,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct WriteBlankInput {
    pub workbook_id: String, pub sheet_name: String, pub cell: String,
    #[serde(default)] pub bold: Option<bool>, #[serde(default)] pub background_color: Option<String>,
    #[serde(default)] pub number_format: Option<String>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ClearCellInput {
    pub workbook_id: String, pub sheet_name: String, pub cell: String,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetCalcModeInput {
    pub workbook_id: String,
    /// "auto", "manual", or "auto_no_table"
    pub mode: String,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetPropertiesInput {
    pub workbook_id: String,
    #[serde(default)] pub title: Option<String>,
    #[serde(default)] pub author: Option<String>,
    #[serde(default)] pub subject: Option<String>,
    #[serde(default)] pub company: Option<String>,
    #[serde(default)] pub description: Option<String>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct MoveWorksheetInput {
    pub workbook_id: String, pub sheet_name: String,
    /// 0-based target position
    pub to_index: usize,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct WriteInternalLinkInput {
    pub workbook_id: String, pub sheet_name: String, pub cell: String,
    /// Internal location (e.g. "Sheet2!A1")
    pub location: String,
    pub display_text: String,
}

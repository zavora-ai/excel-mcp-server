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
    #[serde(default)]
    pub top: Option<f64>,
    #[serde(default)]
    pub bottom: Option<f64>,
    #[serde(default)]
    pub left: Option<f64>,
    #[serde(default)]
    pub right: Option<f64>,
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
    #[serde(default)]
    pub landscape: Option<bool>,
    #[serde(default)]
    pub paper_size: Option<u8>,
    #[serde(default)]
    pub margins: Option<PageMargins>,
    #[serde(default)]
    pub fit_to_pages: Option<FitToPages>,
    #[serde(default)]
    pub print_scale: Option<u16>,
    /// Print area in A1:B2 notation
    #[serde(default)]
    pub print_area: Option<String>,
    #[serde(default)]
    pub repeat_rows: Option<RepeatRows>,
    /// Excel header string (supports &L, &C, &R, &P, &N, &D codes)
    #[serde(default)]
    pub header: Option<String>,
    /// Excel footer string
    #[serde(default)]
    pub footer: Option<String>,
    #[serde(default)]
    pub print_gridlines: Option<bool>,
    #[serde(default)]
    pub center_horizontally: Option<bool>,
    #[serde(default)]
    pub center_vertically: Option<bool>,
}

/// Input for adding a comment/note to a cell
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddCommentInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub cell: String,
    pub text: String,
    #[serde(default)]
    pub author: Option<String>,
}

/// Input for adding a hyperlink to a cell
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddHyperlinkInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub cell: String,
    pub url: String,
    #[serde(default)]
    pub tooltip: Option<String>,
    /// Display text (if different from URL)
    #[serde(default)]
    pub display_text: Option<String>,
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
    #[serde(default)]
    pub hidden: Option<bool>,
    #[serde(default)]
    pub very_hidden: Option<bool>,
    #[serde(default)]
    pub zoom: Option<u16>,
    #[serde(default)]
    pub hide_gridlines: Option<bool>,
    #[serde(default)]
    pub hide_headings: Option<bool>,
    #[serde(default)]
    pub tab_color: Option<String>,
    #[serde(default)]
    pub right_to_left: Option<bool>,
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

fn default_level() -> u8 {
    1
}

/// Input for protecting a sheet
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ProtectSheetInput {
    pub workbook_id: String,
    pub sheet_name: String,
    #[serde(default)]
    pub password: Option<String>,
}

/// Input for protecting a workbook
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ProtectWorkbookInput {
    pub workbook_id: String,
    #[serde(default)]
    pub password: Option<String>,
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
    #[serde(default)]
    pub categories: Option<String>,
    /// Series name
    #[serde(default)]
    pub name: Option<String>,
    /// Hex color for the series
    #[serde(default)]
    pub color: Option<String>,
    /// Show data labels on this series
    #[serde(default)]
    pub data_labels: Option<bool>,
    /// Trendline type: linear, exponential, polynomial, power, logarithmic, moving_average
    #[serde(default)]
    pub trendline: Option<String>,
    /// Marker type: circle, diamond, square, triangle, none
    #[serde(default)]
    pub marker: Option<String>,
    /// Use secondary Y axis
    #[serde(default)]
    pub secondary_axis: Option<bool>,
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
    #[serde(default)]
    pub series: Vec<ChartSeriesInput>,
    /// Simple data range (used if series is empty)
    #[serde(default)]
    pub data_range: Option<String>,
    /// Cell where chart top-left is placed (e.g. "E2"). Default: "A1"
    #[serde(default)]
    pub cell: Option<String>,
    #[serde(default)]
    pub title: Option<String>,
    #[serde(default)]
    pub x_axis_label: Option<String>,
    #[serde(default)]
    pub y_axis_label: Option<String>,
    #[serde(default)]
    pub legend_position: Option<LegendPosition>,
    #[serde(default = "default_chart_width")]
    pub width: u32,
    #[serde(default = "default_chart_height")]
    pub height: u32,
    /// Link chart to a pivot table
    #[serde(default)]
    pub pivot_source: Option<PivotChartSourceInput>,
}

/// Pivot table value field
#[derive(Deserialize, JsonSchema)]
pub struct PivotValueFieldInput {
    pub field: String,
    /// Aggregation: sum, count, average, max, min, product, count_nums, std_dev, var
    #[serde(default = "default_sum")]
    pub aggregation: String,
}

fn default_sum() -> String {
    "sum".to_string()
}

/// Input for creating a pivot table
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddPivotTableInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// Cell where pivot table starts (e.g. "A1")
    #[serde(default)]
    pub cell: Option<String>,
    pub name: String,
    /// Source range including sheet (e.g. "'Data'!$A$1:$E$100")
    pub source_range: String,
    #[serde(default)]
    pub row_fields: Vec<String>,
    #[serde(default)]
    pub column_fields: Vec<String>,
    pub value_fields: Vec<PivotValueFieldInput>,
    #[serde(default)]
    pub filter_fields: Vec<String>,
    #[serde(default)]
    pub style: Option<String>,
    /// Layout: compact, outline, tabular. Default: compact
    #[serde(default)]
    pub layout: Option<String>,
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
    #[serde(default)]
    pub bold: Option<bool>,
    #[serde(default)]
    pub italic: Option<bool>,
    #[serde(default)]
    pub color: Option<String>,
    #[serde(default)]
    pub font_size: Option<f64>,
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
    pub workbook_id: String,
    pub sheet_name: String,
    pub column: String,
    #[serde(default)]
    pub bold: Option<bool>,
    #[serde(default)]
    pub italic: Option<bool>,
    #[serde(default)]
    pub font_size: Option<f64>,
    #[serde(default)]
    pub font_color: Option<String>,
    #[serde(default)]
    pub background_color: Option<String>,
    #[serde(default)]
    pub number_format: Option<String>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetRowFormatInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// 1-based row number
    pub row: u32,
    #[serde(default)]
    pub bold: Option<bool>,
    #[serde(default)]
    pub italic: Option<bool>,
    #[serde(default)]
    pub font_size: Option<f64>,
    #[serde(default)]
    pub font_color: Option<String>,
    #[serde(default)]
    pub background_color: Option<String>,
    #[serde(default)]
    pub number_format: Option<String>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetColumnHiddenInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub column: String,
    #[serde(default)]
    pub hidden: bool,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetRowHiddenInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub row: u32,
    #[serde(default)]
    pub hidden: bool,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetColumnRangeWidthInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub first_column: String,
    pub last_column: String,
    pub width: f64,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetDefaultRowHeightInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub height: f64,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetSelectionInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub cell: String,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetAutofilterInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// Range in A1:B2 notation
    pub range: String,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct FilterColumnInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub column: String,
    pub values: Vec<String>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct IgnoreErrorInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// Error type: "number_stored_as_text", "formula_range", etc.
    pub error_type: String,
    /// Range in A1:B2 notation
    pub range: String,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetPageBreaksInput {
    pub workbook_id: String,
    pub sheet_name: String,
    #[serde(default)]
    pub row_breaks: Vec<u32>,
    #[serde(default)]
    pub col_breaks: Vec<u16>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct UnprotectRangeInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub range: String,
    pub title: String,
    #[serde(default)]
    pub password: Option<String>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct WriteFormulaInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub cell: String,
    pub formula: String,
    /// Optional cached numeric result
    #[serde(default)]
    pub cached_result: Option<f64>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct WriteArrayFormulaInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// Range the array formula spans (e.g. "A1:C3")
    pub range: String,
    pub formula: String,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct WriteDynamicFormulaInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub cell: String,
    pub formula: String,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct WriteBlankInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub cell: String,
    #[serde(default)]
    pub bold: Option<bool>,
    #[serde(default)]
    pub background_color: Option<String>,
    #[serde(default)]
    pub number_format: Option<String>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ClearCellInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub cell: String,
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
    #[serde(default)]
    pub title: Option<String>,
    #[serde(default)]
    pub author: Option<String>,
    #[serde(default)]
    pub subject: Option<String>,
    #[serde(default)]
    pub company: Option<String>,
    #[serde(default)]
    pub description: Option<String>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct MoveWorksheetInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// 0-based target position
    pub to_index: usize,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct WriteInternalLinkInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub cell: String,
    /// Internal location (e.g. "Sheet2!A1")
    pub location: String,
    pub display_text: String,
}

// ══════════════════════════════════════════════════════════════════
// Consolidated input types (replacing multiple separate inputs)
// ══════════════════════════════════════════════════════════════════

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ConfigureWorkbookInput {
    pub workbook_id: String,
    /// "auto", "manual", or "auto_no_table"
    #[serde(default)]
    pub calc_mode: Option<String>,
    /// 0-based sheet index to make active
    #[serde(default)]
    pub active_sheet: Option<usize>,
    #[serde(default)]
    pub title: Option<String>,
    #[serde(default)]
    pub author: Option<String>,
    #[serde(default)]
    pub subject: Option<String>,
    #[serde(default)]
    pub company: Option<String>,
    #[serde(default)]
    pub description: Option<String>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ModifyRowsInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// "insert" or "delete"
    pub action: String,
    /// 1-based row number
    pub at_row: u32,
    pub count: u32,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ModifyColumnsInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// "insert" or "delete"
    pub action: String,
    pub at_column: String,
    pub count: u16,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct WriteFormulaConsolidatedInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// Cell for regular/dynamic, or range for array (e.g. "A1" or "A1:C3")
    pub cell: String,
    pub formula: String,
    /// "regular" (default), "array", or "dynamic"
    #[serde(default)]
    pub formula_type: Option<String>,
    #[serde(default)]
    pub cached_result: Option<f64>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ManageCellInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub cell: String,
    /// "blank" or "clear"
    pub action: String,
    #[serde(default)]
    pub background_color: Option<String>,
    #[serde(default)]
    pub number_format: Option<String>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ManageCommentsInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// "add" or "read"
    pub action: String,
    #[serde(default)]
    pub cell: Option<String>,
    #[serde(default)]
    pub text: Option<String>,
    #[serde(default)]
    pub author: Option<String>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ManageDefinedNamesInput {
    pub workbook_id: String,
    /// "add" or "list"
    pub action: String,
    #[serde(default)]
    pub name: Option<String>,
    #[serde(default)]
    pub formula: Option<String>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddLinkInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub cell: String,
    /// "url" or "internal"
    #[serde(default = "default_url")]
    pub link_type: String,
    /// URL for external, or "Sheet2!A1" for internal
    pub target: String,
    #[serde(default)]
    pub display_text: Option<String>,
    #[serde(default)]
    pub tooltip: Option<String>,
}

fn default_url() -> String {
    "url".to_string()
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ProtectInput {
    pub workbook_id: String,
    /// "sheet", "workbook", or "unprotect_range"
    pub target: String,
    #[serde(default)]
    pub sheet_name: Option<String>,
    #[serde(default)]
    pub password: Option<String>,
    /// For unprotect_range: the range to allow editing
    #[serde(default)]
    pub range: Option<String>,
    /// For unprotect_range: title for the range
    #[serde(default)]
    pub range_title: Option<String>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetDimensionsInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// "column_width", "row_height", "column_range_width", or "default_row_height"
    pub target: String,
    #[serde(default)]
    pub column: Option<String>,
    #[serde(default)]
    pub first_column: Option<String>,
    #[serde(default)]
    pub last_column: Option<String>,
    /// 1-based row number
    #[serde(default)]
    pub row: Option<u32>,
    pub value: f64,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetVisibilityInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// "row" or "column"
    pub target: String,
    /// Column letter (for column) or 1-based row number as string (for row)
    pub identifier: String,
    pub hidden: bool,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetRowColumnFormatInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// "row" or "column"
    pub target: String,
    /// Column letter or 1-based row number as string
    pub identifier: String,
    #[serde(default)]
    pub bold: Option<bool>,
    #[serde(default)]
    pub italic: Option<bool>,
    #[serde(default)]
    pub font_size: Option<f64>,
    #[serde(default)]
    pub font_color: Option<String>,
    #[serde(default)]
    pub background_color: Option<String>,
    #[serde(default)]
    pub number_format: Option<String>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct GroupInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// "rows" or "columns"
    pub target: String,
    /// For rows: 1-based row numbers. For columns: column letters.
    pub start: String,
    pub end: String,
    #[serde(default = "default_level")]
    pub level: u8,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ManageAutofilterInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// Range in A1:B2 notation (sets autofilter)
    pub range: String,
    /// Optional: column letter to filter
    #[serde(default)]
    pub filter_column: Option<String>,
    /// Optional: filter values for the column
    #[serde(default)]
    pub filter_values: Option<Vec<String>>,
}

// ── Batch 9: Feature-parity tools ──────────────────────────────────

/// A single data point for a waterfall chart
#[derive(Deserialize, JsonSchema)]
pub struct WaterfallPoint {
    pub category: String,
    pub value: f64,
    /// "increase", "decrease", or "total"
    pub point_type: WaterfallPointKind,
}

/// Input for adding a waterfall chart (Excel 2016+ ChartEx format)
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddWaterfallChartInput {
    pub workbook_id: String,
    pub sheet_name: String,
    #[serde(default)]
    pub title: Option<String>,
    #[serde(default)]
    pub series_name: Option<String>,
    /// Data points for the waterfall chart
    pub points: Vec<WaterfallPoint>,
    /// Chart width in pixels. Default: 480
    #[serde(default = "default_chart_width")]
    pub width: u32,
    /// Chart height in pixels. Default: 288
    #[serde(default = "default_chart_height")]
    pub height: u32,
    /// Anchor cell for the chart. Default: "A1"
    #[serde(default)]
    pub cell: Option<String>,
}

/// A single data point for a funnel chart
#[derive(Deserialize, JsonSchema)]
pub struct FunnelPoint {
    pub category: String,
    pub value: f64,
}

/// Input for adding a funnel chart (Excel 2016+ ChartEx format)
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddFunnelChartInput {
    pub workbook_id: String,
    pub sheet_name: String,
    #[serde(default)]
    pub title: Option<String>,
    #[serde(default)]
    pub series_name: Option<String>,
    /// Data points for the funnel chart
    pub points: Vec<FunnelPoint>,
    /// Chart width in pixels. Default: 480
    #[serde(default = "default_chart_width")]
    pub width: u32,
    /// Chart height in pixels. Default: 288
    #[serde(default = "default_chart_height")]
    pub height: u32,
    /// Anchor cell for the chart. Default: "A1"
    #[serde(default)]
    pub cell: Option<String>,
}

/// A single data point for a treemap chart
#[derive(Deserialize, JsonSchema)]
pub struct TreemapPoint {
    pub category: String,
    pub value: f64,
    /// Optional hex color (e.g. "#FF0000")
    #[serde(default)]
    pub color: Option<String>,
}

/// Input for adding a treemap chart (Excel 2016+ ChartEx format)
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddTreemapChartInput {
    pub workbook_id: String,
    pub sheet_name: String,
    #[serde(default)]
    pub title: Option<String>,
    #[serde(default)]
    pub series_name: Option<String>,
    /// Data points for the treemap chart
    pub points: Vec<TreemapPoint>,
    /// Chart width in pixels. Default: 480
    #[serde(default = "default_chart_width")]
    pub width: u32,
    /// Chart height in pixels. Default: 288
    #[serde(default = "default_chart_height")]
    pub height: u32,
    /// Anchor cell for the chart. Default: "A1"
    #[serde(default)]
    pub cell: Option<String>,
}

/// Input for adding a drawing shape to a worksheet
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddShapeInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// Anchor cell (e.g. "B2")
    pub cell: String,
    /// Shape type: rectangle, rounded_rectangle, ellipse, triangle, diamond, arrow, callout, text_box
    pub shape_type: ShapeKind,
    /// Width in pixels
    pub width: u32,
    /// Height in pixels
    pub height: u32,
    /// Optional text body for the shape
    #[serde(default)]
    pub text: Option<String>,
    /// Fill color as hex (e.g. "#FF0000")
    #[serde(default)]
    pub fill_color: Option<String>,
    /// Outline color as hex (e.g. "#000000")
    #[serde(default)]
    pub outline_color: Option<String>,
    /// Outline width in points
    #[serde(default)]
    pub outline_width: Option<f64>,
    /// Font size for the text body
    #[serde(default)]
    pub font_size: Option<f64>,
    /// Whether the text should be bold
    #[serde(default)]
    pub bold: Option<bool>,
}

/// Input for setting document properties (core + app)
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SetDocPropertiesInput {
    pub workbook_id: String,
    #[serde(default)]
    pub title: Option<String>,
    #[serde(default)]
    pub author: Option<String>,
    #[serde(default)]
    pub subject: Option<String>,
    #[serde(default)]
    pub description: Option<String>,
    #[serde(default)]
    pub keywords: Option<String>,
    #[serde(default)]
    pub category: Option<String>,
    #[serde(default)]
    pub company: Option<String>,
}

// ══════════════════════════════════════════════════════════════════
// v0.2.0: New tools — charts, pivots, controls, protection, save formats
// ══════════════════════════════════════════════════════════════════

/// Input for adding a sunburst chart (Excel 2016+ ChartEx)
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddSunburstChartInput {
    pub workbook_id: String,
    pub sheet_name: String,
    #[serde(default)]
    pub title: Option<String>,
    #[serde(default)]
    pub series_name: Option<String>,
    pub points: Vec<TreemapPoint>,
    #[serde(default = "default_chart_width")]
    pub width: u32,
    #[serde(default = "default_chart_height")]
    pub height: u32,
    #[serde(default)]
    pub cell: Option<String>,
}

/// Input for adding a histogram chart (Excel 2016+ ChartEx)
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddHistogramChartInput {
    pub workbook_id: String,
    pub sheet_name: String,
    #[serde(default)]
    pub title: Option<String>,
    #[serde(default)]
    pub series_name: Option<String>,
    pub points: Vec<FunnelPoint>,
    #[serde(default)]
    pub bin_count: Option<u32>,
    #[serde(default)]
    pub bin_width: Option<f64>,
    /// Show Pareto line overlay
    #[serde(default)]
    pub pareto: Option<bool>,
    #[serde(default = "default_chart_width")]
    pub width: u32,
    #[serde(default = "default_chart_height")]
    pub height: u32,
    #[serde(default)]
    pub cell: Option<String>,
}

/// Input for adding a box & whisker chart (Excel 2016+ ChartEx)
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddBoxWhiskerChartInput {
    pub workbook_id: String,
    pub sheet_name: String,
    #[serde(default)]
    pub title: Option<String>,
    #[serde(default)]
    pub series_name: Option<String>,
    pub points: Vec<FunnelPoint>,
    /// Show outlier points
    #[serde(default)]
    pub show_outliers: Option<bool>,
    /// Show mean markers
    #[serde(default)]
    pub show_mean: Option<bool>,
    /// Show inner data points
    #[serde(default)]
    pub show_inner_points: Option<bool>,
    #[serde(default = "default_chart_width")]
    pub width: u32,
    #[serde(default = "default_chart_height")]
    pub height: u32,
    #[serde(default)]
    pub cell: Option<String>,
}

/// Input for adding a map chart (Excel 2016+ ChartEx)
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddMapChartInput {
    pub workbook_id: String,
    pub sheet_name: String,
    #[serde(default)]
    pub title: Option<String>,
    #[serde(default)]
    pub series_name: Option<String>,
    pub points: Vec<FunnelPoint>,
    /// Map level: "world", "continent", "country", "state"
    #[serde(default)]
    pub map_level: Option<String>,
    #[serde(default = "default_chart_width")]
    pub width: u32,
    #[serde(default = "default_chart_height")]
    pub height: u32,
    #[serde(default)]
    pub cell: Option<String>,
}

/// Input for adding a slicer to a pivot table
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddSlicerInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// Name of the pivot table to connect to
    pub pivot_table_name: String,
    /// Field name to filter on
    pub field_name: String,
    /// Anchor cell for the slicer
    #[serde(default)]
    pub cell: Option<String>,
    #[serde(default)]
    pub width: Option<u32>,
    #[serde(default)]
    pub height: Option<u32>,
}

/// Input for adding a timeline to a pivot table
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddTimelineInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// Name of the pivot table to connect to
    pub pivot_table_name: String,
    /// Date field name
    pub field_name: String,
    /// Anchor cell for the timeline
    #[serde(default)]
    pub cell: Option<String>,
    #[serde(default)]
    pub width: Option<u32>,
    #[serde(default)]
    pub height: Option<u32>,
}

/// Input for adding a form control (button, checkbox, spinner, dropdown)
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddFormControlInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// Anchor cell
    pub cell: String,
    /// Control type: "button", "checkbox", "spinner", "dropdown", "radio", "scroll_bar", "group_box", "label"
    pub control_type: String,
    /// Display text/label
    #[serde(default)]
    pub text: Option<String>,
    /// Cell link for the control value
    #[serde(default)]
    pub cell_link: Option<String>,
    /// Input range (for dropdown/list controls)
    #[serde(default)]
    pub input_range: Option<String>,
    #[serde(default)]
    pub width: Option<u32>,
    #[serde(default)]
    pub height: Option<u32>,
}

/// Input for saving in different formats
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct SaveWorkbookAdvancedInput {
    pub workbook_id: String,
    pub file_path: String,
    /// Save format: "xlsx" (default), "xlsm", "template", "template_macro", "encrypted", "parallel"
    #[serde(default = "default_xlsx")]
    pub format: String,
    /// Password for encrypted save
    #[serde(default)]
    pub password: Option<String>,
}

fn default_xlsx() -> String {
    "xlsx".to_string()
}

/// Input for opening a password-protected workbook
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct OpenWorkbookEncryptedInput {
    pub file_path: String,
    pub password: String,
}

/// Input for managing named ranges with full CRUD and scoping
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ManageNamedRangesInput {
    pub workbook_id: String,
    /// "add", "update", "remove", "list", "add_scoped"
    pub action: String,
    #[serde(default)]
    pub name: Option<String>,
    #[serde(default)]
    pub formula: Option<String>,
    /// Sheet index for scoped names
    #[serde(default)]
    pub sheet_index: Option<usize>,
}

/// Input for reading worksheet metadata (used_range, hyperlinks, merge_ranges, charts)
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct ReadSheetMetadataInput {
    pub workbook_id: String,
    pub sheet_name: String,
    /// What to read: "used_range", "hyperlinks", "merge_ranges", "charts", "all"
    #[serde(default = "default_all")]
    pub info: String,
}

fn default_all() -> String {
    "all".to_string()
}

/// Input for adding a chart sheet (dedicated chart-only sheet)
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub struct AddChartSheetInput {
    pub workbook_id: String,
    pub sheet_name: String,
    pub chart_type: ChartType,
    #[serde(default)]
    pub series: Vec<ChartSeriesInput>,
    #[serde(default)]
    pub data_range: Option<String>,
    #[serde(default)]
    pub title: Option<String>,
    #[serde(default)]
    pub x_axis_label: Option<String>,
    #[serde(default)]
    pub y_axis_label: Option<String>,
    #[serde(default)]
    pub legend_position: Option<LegendPosition>,
}

/// Chart enhancement options (extend existing add_chart)
#[derive(Deserialize, JsonSchema)]
pub struct ChartEnhancements {
    /// Show data table below chart
    #[serde(default)]
    pub show_data_table: Option<bool>,
    /// 3D perspective rotation
    #[serde(default)]
    pub view_3d: Option<View3DInput>,
    /// Preset chart style number (1-48)
    #[serde(default)]
    pub style: Option<u8>,
    /// Accessibility alt text (title, description)
    #[serde(default)]
    pub alt_text: Option<AltTextInput>,
    /// Y-axis minimum value
    #[serde(default)]
    pub y_axis_min: Option<f64>,
    /// Y-axis maximum value
    #[serde(default)]
    pub y_axis_max: Option<f64>,
    /// Y-axis logarithmic base (e.g. 10)
    #[serde(default)]
    pub y_axis_log_base: Option<f64>,
    /// Reverse X axis
    #[serde(default)]
    pub x_axis_reverse: Option<bool>,
    /// Reverse Y axis
    #[serde(default)]
    pub y_axis_reverse: Option<bool>,
    /// X-axis number format
    #[serde(default)]
    pub x_axis_format: Option<String>,
    /// Y-axis number format
    #[serde(default)]
    pub y_axis_format: Option<String>,
    /// Show drop lines
    #[serde(default)]
    pub drop_lines: Option<bool>,
    /// Show high-low lines
    #[serde(default)]
    pub high_low_lines: Option<bool>,
    /// Plot area background fill color (hex)
    #[serde(default)]
    pub plot_area_fill: Option<String>,
}

#[derive(Deserialize, JsonSchema)]
pub struct View3DInput {
    #[serde(default)]
    pub rot_x: Option<i16>,
    #[serde(default)]
    pub rot_y: Option<i16>,
    #[serde(default)]
    pub perspective: Option<u8>,
}

#[derive(Deserialize, JsonSchema)]
pub struct AltTextInput {
    pub title: String,
    pub description: String,
}

/// Enhanced chart series with additional options
#[derive(Deserialize, JsonSchema)]
pub struct ChartSeriesEnhanced {
    pub values: String,
    #[serde(default)]
    pub categories: Option<String>,
    #[serde(default)]
    pub name: Option<String>,
    #[serde(default)]
    pub color: Option<String>,
    #[serde(default)]
    pub data_labels: Option<bool>,
    #[serde(default)]
    pub trendline: Option<String>,
    #[serde(default)]
    pub marker: Option<String>,
    #[serde(default)]
    pub secondary_axis: Option<bool>,
    /// Line width in points
    #[serde(default)]
    pub line_width: Option<f64>,
    /// Dash style: "solid", "dash", "dot", "dash_dot"
    #[serde(default)]
    pub dash_style: Option<String>,
    /// Gradient stops: array of [color_hex, position_0_to_1]
    #[serde(default)]
    pub gradient: Option<Vec<GradientStopInput>>,
    /// Bubble sizes range (for bubble charts)
    #[serde(default)]
    pub bubble_sizes: Option<String>,
}

#[derive(Deserialize, JsonSchema)]
pub struct GradientStopInput {
    pub color: String,
    pub position: f64,
}

/// Sheet protection with granular options
#[derive(Deserialize, JsonSchema)]
pub struct SheetProtectionOptionsInput {
    #[serde(default)]
    pub password: Option<String>,
    #[serde(default)]
    pub insert_rows: Option<bool>,
    #[serde(default)]
    pub delete_rows: Option<bool>,
    #[serde(default)]
    pub insert_columns: Option<bool>,
    #[serde(default)]
    pub delete_columns: Option<bool>,
    #[serde(default)]
    pub format_cells: Option<bool>,
    #[serde(default)]
    pub format_columns: Option<bool>,
    #[serde(default)]
    pub format_rows: Option<bool>,
    #[serde(default)]
    pub sort: Option<bool>,
    #[serde(default)]
    pub insert_hyperlinks: Option<bool>,
    #[serde(default)]
    pub select_locked_cells: Option<bool>,
    #[serde(default)]
    pub select_unlocked_cells: Option<bool>,
    #[serde(default)]
    pub pivot_tables: Option<bool>,
}

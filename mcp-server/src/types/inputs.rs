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

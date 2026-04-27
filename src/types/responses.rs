use serde::{Deserialize, Serialize};

/// Status indicator for all tool responses.
#[derive(Debug, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum Status {
    Success,
    Error,
}

/// Structured response wrapper returned by every tool.
#[derive(Debug, Serialize)]
pub struct ToolResponse<T: Serialize> {
    pub status: Status,
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data: Option<T>,
}

/// Error payload included in error responses.
#[derive(Debug, Serialize)]
pub struct ErrorData {
    pub category: ErrorCategory,
    pub description: String,
    pub suggestion: String,
}

/// Categorises the kind of error that occurred.
#[derive(Debug, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum ErrorCategory {
    NotFound,
    InvalidInput,
    EngineUnsupported,
    CapacityExceeded,
    IoError,
    ParseError,
    Evicted,
}

/// Build a success response with data, serialised to JSON.
pub fn success<T: Serialize>(message: &str, data: T) -> String {
    serde_json::to_string(&ToolResponse {
        status: Status::Success,
        message: message.to_string(),
        data: Some(data),
    })
    .expect("serialization of ToolResponse should never fail")
}

/// Build a success response without data, serialised to JSON.
pub fn success_no_data(message: &str) -> String {
    serde_json::to_string(&ToolResponse::<()> {
        status: Status::Success,
        message: message.to_string(),
        data: None,
    })
    .expect("serialization of ToolResponse should never fail")
}

/// Build an error response, serialised to JSON.
pub fn error(category: ErrorCategory, description: &str, suggestion: &str) -> String {
    serde_json::to_string(&ToolResponse {
        status: Status::Error,
        message: description.to_string(),
        data: Some(ErrorData {
            category,
            description: description.to_string(),
            suggestion: suggestion.to_string(),
        }),
    })
    .expect("serialization of ToolResponse should never fail")
}

/// Data returned when a workbook is created or opened.
#[derive(Debug, Serialize)]
pub struct WorkbookInfo {
    pub workbook_id: String,
    pub engine: String,
    pub sheets: Vec<SheetSummary>,
}

/// Summary of a single sheet within a workbook.
#[derive(Debug, Serialize)]
pub struct SheetSummary {
    pub name: String,
    pub dimensions: Option<String>,
    pub row_count: Option<u32>,
    pub col_count: Option<u16>,
}

/// Paginated read result for sheet data.
#[derive(Debug, Serialize)]
pub struct ReadSheetData {
    pub rows: Vec<Vec<serde_json::Value>>,
    pub total_rows: u32,
    pub page_rows: u32,
    pub continuation_token: Option<String>,
}

/// Single cell read result.
#[derive(Debug, Serialize)]
pub struct CellData {
    pub cell: String,
    pub value: serde_json::Value,
    pub value_type: String,
    pub formula: Option<String>,
}

/// Write confirmation returned after cell writes.
#[derive(Debug, Serialize)]
pub struct WriteResult {
    pub cells_written: usize,
    pub range_covered: String,
}

/// Search result containing matching cells.
#[derive(Debug, Serialize)]
pub struct SearchResult {
    pub matches: Vec<SearchMatch>,
    pub total_matches: usize,
    pub truncated: bool,
}

/// A single cell match from a search operation.
#[derive(Debug, Serialize)]
pub struct SearchMatch {
    pub sheet: String,
    pub cell: String,
    pub value: serde_json::Value,
}

/// Full workbook description with sample data per sheet.
#[derive(Debug, Serialize)]
pub struct WorkbookDescription {
    pub workbook_id: String,
    pub engine: String,
    pub sheets: Vec<SheetDescription>,
}

/// Sheet description including dimensions and sample rows.
#[derive(Debug, Serialize)]
pub struct SheetDescription {
    pub name: String,
    pub dimensions: Option<String>,
    pub row_count: Option<u32>,
    pub col_count: Option<u16>,
    pub sample_rows: Vec<Vec<serde_json::Value>>,
}

/// CSV export result.
#[derive(Debug, Serialize)]
pub struct CsvExportData {
    pub csv: String,
    pub total_rows: u32,
    pub truncated: bool,
}

/// Continuation token for paginated read_sheet responses.
#[derive(Debug, Serialize, Deserialize)]
pub struct ContinuationToken {
    pub sheet: String,
    pub offset: u32,
    pub range: Option<String>,
}

/// Result from batch_format.
#[derive(Debug, Serialize)]
pub struct BatchFormatResult {
    pub operations_applied: usize,
    pub failures: Vec<BatchFormatFailure>,
}

/// A single failure within a batch_format operation.
#[derive(Debug, Serialize)]
pub struct BatchFormatFailure {
    pub operation_index: usize,
    pub range: String,
    pub error: String,
}

/// Result from copy_format.
#[derive(Debug, Serialize)]
pub struct CopyFormatResult {
    pub targets_formatted: usize,
    pub note: Option<String>,
}

/// Result from describe_formatting.
#[derive(Debug, Serialize)]
pub struct DescribeFormattingResult {
    pub format_groups: Vec<FormatGroup>,
}

/// A group of cells sharing identical formatting.
#[derive(Debug, Serialize)]
pub struct FormatGroup {
    pub ranges: Vec<String>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub font_size: Option<f64>,
    pub font_color: Option<String>,
    pub background_color: Option<String>,
    pub number_format: Option<String>,
    pub horizontal_alignment: Option<String>,
    pub vertical_alignment: Option<String>,
    pub border_style: Option<String>,
}

// ── Tier 2: Writer Agent responses ─────────────────────────────────

#[derive(Debug, Serialize)]
pub struct WriteGridResult {
    pub rows_written: usize,
    pub columns_written: usize,
    pub cells_written: usize,
    pub failures: Vec<String>,
}

#[derive(Debug, Serialize)]
pub struct WriteRowRangeResult {
    pub cells_written: usize,
}

#[derive(Debug, Serialize)]
pub struct CloneFormulasResult {
    pub formulas_cloned: usize,
    pub columns_filled: usize,
}

// ── Tier 3: Data Operations responses ──────────────────────────────

#[derive(Debug, Serialize)]
pub struct SortResult {
    pub rows_sorted: usize,
}

#[derive(Debug, Serialize)]
pub struct FindReplaceResult {
    pub replacements: usize,
}

#[derive(Debug, Serialize)]
pub struct FillSeriesResult {
    pub cells_filled: usize,
}

#[derive(Debug, Serialize)]
pub struct DeleteRowsResult {
    pub rows_deleted: usize,
}

#[derive(Debug, Serialize)]
pub struct TransposeResult {
    pub original_rows: usize,
    pub original_columns: usize,
    pub transposed_rows: usize,
    pub transposed_columns: usize,
}

#[derive(Debug, Serialize)]
pub struct RemoveDuplicatesResult {
    pub rows_removed: usize,
    pub rows_remaining: usize,
}

#[derive(Debug, Serialize)]
pub struct SplitColumnResult {
    pub rows_split: usize,
    pub output_columns: usize,
}

//! ExcelMcpServer struct and ServerHandler implementation.
//!
//! Wires all 28 MCP tools to the tool modules via the rmcp SDK.

use std::sync::Arc;
use tokio::sync::RwLock;

use rmcp::{
    ServerHandler, tool, tool_handler, tool_router,
    handler::server::{router::tool::ToolRouter, wrapper::Parameters},
    model::{ServerCapabilities, ServerInfo},
};

use crate::store::WorkbookStore;
use crate::tools;
use crate::types::inputs::*;
use crate::types::responses::{ErrorCategory, error};

/// The MCP server that exposes Excel manipulation tools to LLM clients.
#[derive(Debug, Clone)]
pub struct ExcelMcpServer {
    tool_router: ToolRouter<Self>,
    store: Arc<RwLock<WorkbookStore>>,
}

impl ExcelMcpServer {
    /// Create a new server instance with the given shared workbook store.
    pub fn new(store: Arc<RwLock<WorkbookStore>>) -> Self {
        Self {
            tool_router: Self::tool_router(),
            store,
        }
    }
}

/// Helper to convert unexpected errors into a JSON error response string.
fn unexpected_error(e: anyhow::Error) -> String {
    error(
        ErrorCategory::IoError,
        &format!("Unexpected error: {}", e),
        "Please try again or report this issue.",
    )
}

#[tool_router]
impl ExcelMcpServer {
    // ── Workbook lifecycle ──────────────────────────────────────

    #[tool(description = "Create a new empty Excel workbook in memory")]
    async fn create_workbook(&self, _params: Parameters<CreateWorkbookInput>) -> String {
        let mut store = self.store.write().await;
        match tools::workbook::create_workbook(&mut store) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Open an existing Excel file for reading or editing")]
    async fn open_workbook(&self, Parameters(input): Parameters<OpenWorkbookInput>) -> String {
        let mut store = self.store.write().await;
        match tools::workbook::open_workbook(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Save a workbook to disk as an xlsx file")]
    async fn save_workbook(&self, Parameters(input): Parameters<SaveWorkbookInput>) -> String {
        let mut store = self.store.write().await;
        match tools::workbook::save_workbook(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Close a workbook and free its memory")]
    async fn close_workbook(&self, Parameters(input): Parameters<CloseWorkbookInput>) -> String {
        let mut store = self.store.write().await;
        match tools::workbook::close_workbook(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    // ── Sheet management ────────────────────────────────────────

    #[tool(description = "List all sheet names in a workbook")]
    async fn list_sheets(&self, Parameters(input): Parameters<ListSheetsInput>) -> String {
        let mut store = self.store.write().await;
        match tools::sheets::list_sheets(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Get the dimensions of a sheet (used range, row count, column count)")]
    async fn get_sheet_dimensions(&self, Parameters(input): Parameters<GetSheetDimensionsInput>) -> String {
        let mut store = self.store.write().await;
        match tools::sheets::get_sheet_dimensions(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Describe a workbook's structure including sheet names, dimensions, and sample data")]
    async fn describe_workbook(&self, Parameters(input): Parameters<DescribeWorkbookInput>) -> String {
        let mut store = self.store.write().await;
        match tools::sheets::describe_workbook(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Add a new empty worksheet to a workbook")]
    async fn add_sheet(&self, Parameters(input): Parameters<AddSheetInput>) -> String {
        let mut store = self.store.write().await;
        match tools::sheets::add_sheet(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Rename an existing worksheet")]
    async fn rename_sheet(&self, Parameters(input): Parameters<RenameSheetInput>) -> String {
        let mut store = self.store.write().await;
        match tools::sheets::rename_sheet(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Delete a worksheet from a workbook")]
    async fn delete_sheet(&self, Parameters(input): Parameters<DeleteSheetInput>) -> String {
        let mut store = self.store.write().await;
        match tools::sheets::delete_sheet(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    // ── Read tools ──────────────────────────────────────────────

    #[tool(description = "Read data from a worksheet with optional range and pagination")]
    async fn read_sheet(&self, Parameters(input): Parameters<ReadSheetInput>) -> String {
        let mut store = self.store.write().await;
        match tools::read::read_sheet(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Read a single cell's value, type, and formula")]
    async fn read_cell(&self, Parameters(input): Parameters<ReadCellInput>) -> String {
        let mut store = self.store.write().await;
        match tools::read::read_cell(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Search for cells matching a value or pattern across sheets")]
    async fn search_cells(&self, Parameters(input): Parameters<SearchCellsInput>) -> String {
        let mut store = self.store.write().await;
        match tools::read::search_cells(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Export a sheet as a CSV-formatted string")]
    async fn sheet_to_csv(&self, Parameters(input): Parameters<SheetToCsvInput>) -> String {
        let mut store = self.store.write().await;
        match tools::read::sheet_to_csv(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    // ── Write tools ─────────────────────────────────────────────

    #[tool(description = "Write values to multiple cells in a single batch operation")]
    async fn write_cells(&self, Parameters(input): Parameters<WriteCellsInput>) -> String {
        let mut store = self.store.write().await;
        match tools::write::write_cells(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Write a row of values starting from a cell, filling rightward")]
    async fn write_row(&self, Parameters(input): Parameters<WriteRowInput>) -> String {
        let mut store = self.store.write().await;
        match tools::write::write_row(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Write a column of values starting from a cell, filling downward")]
    async fn write_column(&self, Parameters(input): Parameters<WriteColumnInput>) -> String {
        let mut store = self.store.write().await;
        match tools::write::write_column(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    // ── Formatting tools ────────────────────────────────────────

    #[tool(description = "Apply formatting (bold, italic, colors, borders, etc.) to a range of cells")]
    async fn set_cell_format(&self, Parameters(input): Parameters<SetCellFormatInput>) -> String {
        let mut store = self.store.write().await;
        match tools::format::set_cell_format(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Merge a range of cells into a single cell")]
    async fn merge_cells(&self, Parameters(input): Parameters<MergeCellsInput>) -> String {
        let mut store = self.store.write().await;
        match tools::format::merge_cells(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    // ── Charts, images, tables ──────────────────────────────────

    #[tool(description = "Add a chart (bar, column, line, pie, scatter, area, doughnut) to a worksheet")]
    async fn add_chart(&self, Parameters(input): Parameters<AddChartInput>) -> String {
        let mut store = self.store.write().await;
        match tools::charts::add_chart(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Embed a PNG or JPEG image into a worksheet at a specified cell")]
    async fn add_image(&self, Parameters(input): Parameters<AddImageInput>) -> String {
        let mut store = self.store.write().await;
        match tools::images::add_image(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Create an Excel Table with headers, autofilter, and optional totals row")]
    async fn add_table(&self, Parameters(input): Parameters<AddTableInput>) -> String {
        let mut store = self.store.write().await;
        match tools::tables::add_table(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    // ── Conditional formatting, validation, sparklines ──────────

    #[tool(description = "Add a conditional formatting rule (cell value, color scale, data bar, icon set) to a range")]
    async fn add_conditional_format(&self, Parameters(input): Parameters<AddConditionalFormatInput>) -> String {
        let mut store = self.store.write().await;
        match tools::conditional::add_conditional_format(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Add data validation (dropdown list, number range, date range, custom formula) to a range")]
    async fn add_data_validation(&self, Parameters(input): Parameters<AddDataValidationInput>) -> String {
        let mut store = self.store.write().await;
        match tools::validation::add_data_validation(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Add a sparkline (line, column, or win/loss) to a cell")]
    async fn add_sparkline(&self, Parameters(input): Parameters<AddSparklineInput>) -> String {
        let mut store = self.store.write().await;
        match tools::sparklines::add_sparkline(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    // ── Layout tools ────────────────────────────────────────────

    #[tool(description = "Set the width of a column in character units")]
    async fn set_column_width(&self, Parameters(input): Parameters<SetColumnWidthInput>) -> String {
        let mut store = self.store.write().await;
        match tools::layout::set_column_width(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Set the height of a row in points")]
    async fn set_row_height(&self, Parameters(input): Parameters<SetRowHeightInput>) -> String {
        let mut store = self.store.write().await;
        match tools::layout::set_row_height(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Freeze panes at a cell position so rows above and columns left remain visible while scrolling")]
    async fn freeze_panes(&self, Parameters(input): Parameters<FreezePanesInput>) -> String {
        let mut store = self.store.write().await;
        match tools::layout::freeze_panes(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Page setup & print ──────────────────────────────────────

    #[tool(description = "Configure page setup: orientation, margins, headers/footers, print area, fit-to-page, repeat rows")]
    async fn set_page_setup(&self, Parameters(input): Parameters<SetPageSetupInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::set_page_setup(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Comments & hyperlinks ───────────────────────────────────

    #[tool(description = "Add a comment/note to a cell")]
    async fn add_comment(&self, Parameters(input): Parameters<AddCommentInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::add_comment(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Read all comments from a sheet")]
    async fn read_comments(&self, Parameters(input): Parameters<ReadCommentsInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::read_comments(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Add a clickable hyperlink to a cell")]
    async fn add_hyperlink(&self, Parameters(input): Parameters<AddHyperlinkInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::add_hyperlink(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Named ranges ────────────────────────────────────────────

    #[tool(description = "Add a defined name (named range) to the workbook")]
    async fn add_defined_name(&self, Parameters(input): Parameters<AddDefinedNameInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::add_defined_name(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "List all defined names in a workbook")]
    async fn list_defined_names(&self, Parameters(input): Parameters<ListDefinedNamesInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::list_defined_names(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Sheet settings ──────────────────────────────────────────

    #[tool(description = "Configure sheet display: visibility, zoom, gridlines, tab color")]
    async fn set_sheet_settings(&self, Parameters(input): Parameters<SetSheetSettingsInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::set_sheet_settings(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Set which sheet is active (visible) when the workbook opens")]
    async fn set_active_sheet(&self, Parameters(input): Parameters<SetActiveSheetInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::set_active_sheet(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Row/column manipulation ─────────────────────────────────

    #[tool(description = "Insert empty rows at a position, shifting existing rows down")]
    async fn insert_rows(&self, Parameters(input): Parameters<InsertDeleteRowsInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::insert_rows(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Delete rows at a position, shifting remaining rows up")]
    async fn delete_rows(&self, Parameters(input): Parameters<InsertDeleteRowsInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::delete_rows(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Insert empty columns at a position, shifting existing columns right")]
    async fn insert_columns(&self, Parameters(input): Parameters<InsertDeleteColumnsInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::insert_columns(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Delete columns at a position, shifting remaining columns left")]
    async fn delete_columns(&self, Parameters(input): Parameters<InsertDeleteColumnsInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::delete_columns(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Grouping (outline) ──────────────────────────────────────

    #[tool(description = "Group rows into an expandable/collapsible outline")]
    async fn group_rows(&self, Parameters(input): Parameters<GroupRowsInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::group_rows(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Group columns into an expandable/collapsible outline")]
    async fn group_columns(&self, Parameters(input): Parameters<GroupColumnsInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::group_columns(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Protection ──────────────────────────────────────────────

    #[tool(description = "Protect a sheet with optional password")]
    async fn protect_sheet(&self, Parameters(input): Parameters<ProtectSheetInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::protect_sheet(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Protect workbook structure with optional password")]
    async fn protect_workbook(&self, Parameters(input): Parameters<ProtectWorkbookInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::protect_workbook(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Autofit ─────────────────────────────────────────────────

    #[tool(description = "Auto-fit all column widths based on cell content")]
    async fn autofit_columns(&self, Parameters(input): Parameters<AutofitColumnsInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::autofit_columns(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Enhanced charts ─────────────────────────────────────────

    #[tool(description = "Add a chart with full control: multiple series, colors, data labels, trendlines, markers, pivot source")]
    async fn add_chart_enhanced(&self, Parameters(input): Parameters<AddChartEnhancedInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::add_chart_enhanced(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Pivot tables ────────────────────────────────────────────

    #[tool(description = "Create a pivot table with row/column/value/filter fields")]
    async fn add_pivot_table(&self, Parameters(input): Parameters<AddPivotTableInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::add_pivot_table(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Rich text ───────────────────────────────────────────────

    #[tool(description = "Write rich text (mixed bold/italic/color) to a cell")]
    async fn write_rich_text(&self, Parameters(input): Parameters<WriteRichTextInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::write_rich_text(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Column/row format & visibility ──────────────────────────

    #[tool(description = "Apply formatting to an entire column")]
    async fn set_column_format(&self, Parameters(input): Parameters<SetColumnFormatInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::set_column_format(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Apply formatting to an entire row")]
    async fn set_row_format(&self, Parameters(input): Parameters<SetRowFormatInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::set_row_format(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Hide or unhide a column")]
    async fn set_column_hidden(&self, Parameters(input): Parameters<SetColumnHiddenInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::set_column_hidden(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Hide or unhide a row")]
    async fn set_row_hidden(&self, Parameters(input): Parameters<SetRowHiddenInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::set_row_hidden(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Set width for a range of columns at once")]
    async fn set_column_range_width(&self, Parameters(input): Parameters<SetColumnRangeWidthInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::set_column_range_width(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Set the default row height for all rows")]
    async fn set_default_row_height(&self, Parameters(input): Parameters<SetDefaultRowHeightInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::set_default_row_height(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── View settings ───────────────────────────────────────────

    #[tool(description = "Set the selected/active cell in a sheet")]
    async fn set_selection(&self, Parameters(input): Parameters<SetSelectionInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::set_selection(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Autofilter ──────────────────────────────────────────────

    #[tool(description = "Enable autofilter dropdown on a range")]
    async fn set_autofilter(&self, Parameters(input): Parameters<SetAutofilterInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::set_autofilter(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Set filter criteria on a specific column")]
    async fn filter_column(&self, Parameters(input): Parameters<FilterColumnInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::filter_column(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Error suppression ───────────────────────────────────────

    #[tool(description = "Suppress Excel error indicators (green triangles) on a range")]
    async fn ignore_error(&self, Parameters(input): Parameters<IgnoreErrorInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::ignore_error(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Page breaks ─────────────────────────────────────────────

    #[tool(description = "Set manual page breaks for printing")]
    async fn set_page_breaks(&self, Parameters(input): Parameters<SetPageBreaksInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::set_page_breaks(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Unprotect range ─────────────────────────────────────────

    #[tool(description = "Allow editing a specific range on a protected sheet")]
    async fn unprotect_range(&self, Parameters(input): Parameters<UnprotectRangeInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::unprotect_range(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Formula tools ───────────────────────────────────────────

    #[tool(description = "Write a formula to a cell with optional cached result")]
    async fn write_formula(&self, Parameters(input): Parameters<WriteFormulaInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::write_formula(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Write a CSE array formula spanning a range (Ctrl+Shift+Enter)")]
    async fn write_array_formula(&self, Parameters(input): Parameters<WriteArrayFormulaInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::write_array_formula(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Write a dynamic array formula (Excel 365 spill formula)")]
    async fn write_dynamic_formula(&self, Parameters(input): Parameters<WriteDynamicFormulaInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::write_dynamic_formula(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Cell operations ─────────────────────────────────────────

    #[tool(description = "Write a blank cell with formatting (background color, number format)")]
    async fn write_blank(&self, Parameters(input): Parameters<WriteBlankInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::write_blank(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Clear a cell's content and formatting")]
    async fn clear_cell(&self, Parameters(input): Parameters<ClearCellInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::clear_cell(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    // ── Workbook settings ───────────────────────────────────────

    #[tool(description = "Set calculation mode: auto, manual, or auto_no_table")]
    async fn set_calc_mode(&self, Parameters(input): Parameters<SetCalcModeInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::set_calc_mode(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Set document properties: title, author, subject, company, description")]
    async fn set_properties(&self, Parameters(input): Parameters<SetPropertiesInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::set_properties(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Move a worksheet to a different position in the workbook")]
    async fn move_worksheet(&self, Parameters(input): Parameters<MoveWorksheetInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::move_worksheet(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }

    #[tool(description = "Write an internal link to another sheet/cell within the workbook")]
    async fn write_internal_link(&self, Parameters(input): Parameters<WriteInternalLinkInput>) -> String {
        let mut store = self.store.write().await;
        match tools::expanded::write_internal_link(&mut store, input) { Ok(j) => j, Err(e) => unexpected_error(e) }
    }
}

#[tool_handler]
impl ServerHandler for ExcelMcpServer {
    fn get_info(&self) -> ServerInfo {
        ServerInfo::new(ServerCapabilities::builder().enable_tools().build())
            .with_instructions(
                "Excel file manipulation tools. Create, read, edit, format, and analyze Excel workbooks. \
                 Use create_workbook for new files, open_workbook for existing files. \
                 Tools cover: workbook lifecycle, sheet management, cell reading/writing, \
                 formatting, charts, images, tables, conditional formatting, data validation, \
                 sparklines, layout controls, search, and CSV export."
                    .to_string(),
            )
    }
}

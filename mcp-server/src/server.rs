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
        match tools::layout::set_column_width(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Set the height of a row in points")]
    async fn set_row_height(&self, Parameters(input): Parameters<SetRowHeightInput>) -> String {
        let mut store = self.store.write().await;
        match tools::layout::set_row_height(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Freeze panes at a cell position so rows above and columns left remain visible while scrolling")]
    async fn freeze_panes(&self, Parameters(input): Parameters<FreezePanesInput>) -> String {
        let mut store = self.store.write().await;
        match tools::layout::freeze_panes(&mut store, input) {
            Ok(json) => json,
            Err(e) => unexpected_error(e),
        }
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

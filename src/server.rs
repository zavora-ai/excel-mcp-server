//! ExcelMcpServer — 43 consolidated MCP tools backed by zavora-xlsx.

use std::sync::Arc;
use tokio::sync::RwLock;

use rmcp::{
    handler::server::{router::tool::ToolRouter, wrapper::Parameters},
    model::{ServerCapabilities, ServerInfo},
    tool, tool_handler, tool_router, ServerHandler,
};

use crate::store::WorkbookStore;
use crate::tools;
use crate::types::inputs::*;

#[derive(Debug, Clone)]
pub struct ExcelMcpServer {
    tool_router: ToolRouter<Self>,
    store: Arc<RwLock<WorkbookStore>>,
}

impl ExcelMcpServer {
    pub fn new(store: Arc<RwLock<WorkbookStore>>) -> Self {
        Self {
            tool_router: Self::tool_router(),
            store,
        }
    }
}

fn unexpected_error(e: anyhow::Error) -> String {
    crate::types::responses::error(
        crate::types::responses::ErrorCategory::IoError,
        &format!("Unexpected error: {}", e),
        "Please try again.",
    )
}

macro_rules! tool_fn {
    ($store:expr, $module:path, $input:expr) => {{
        let mut store = $store.write().await;
        match $module(&mut store, $input) {
            Ok(j) => j,
            Err(e) => unexpected_error(e),
        }
    }};
}

#[tool_router]
impl ExcelMcpServer {
    // ── Workbook lifecycle (4) ──

    #[tool(description = "Create a new empty Excel workbook in memory")]
    async fn create_workbook(&self, _p: Parameters<CreateWorkbookInput>) -> String {
        let mut store = self.store.write().await;
        match tools::workbook::create_workbook(&mut store) {
            Ok(j) => j,
            Err(e) => unexpected_error(e),
        }
    }

    #[tool(description = "Open an existing Excel file for reading or editing")]
    async fn open_workbook(&self, Parameters(i): Parameters<OpenWorkbookInput>) -> String {
        tool_fn!(self.store, tools::workbook::open_workbook, i)
    }

    #[tool(description = "Save a workbook to disk as an xlsx file")]
    async fn save_workbook(&self, Parameters(i): Parameters<SaveWorkbookInput>) -> String {
        tool_fn!(self.store, tools::workbook::save_workbook, i)
    }

    #[tool(description = "Close a workbook and free its memory")]
    async fn close_workbook(&self, Parameters(i): Parameters<CloseWorkbookInput>) -> String {
        tool_fn!(self.store, tools::workbook::close_workbook, i)
    }

    // ── Workbook configuration (1) ──

    #[tool(
        description = "Configure workbook: set calc mode (auto/manual), active sheet, document properties (title, author, company)"
    )]
    async fn configure_workbook(
        &self,
        Parameters(i): Parameters<ConfigureWorkbookInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::configure_workbook, i)
    }

    // ── Sheet management (7) ──

    #[tool(description = "List all sheet names in a workbook")]
    async fn list_sheets(&self, Parameters(i): Parameters<ListSheetsInput>) -> String {
        tool_fn!(self.store, tools::sheets::list_sheets, i)
    }

    #[tool(description = "Get the dimensions of a sheet (used range, row count, column count)")]
    async fn get_sheet_dimensions(
        &self,
        Parameters(i): Parameters<GetSheetDimensionsInput>,
    ) -> String {
        tool_fn!(self.store, tools::sheets::get_sheet_dimensions, i)
    }

    #[tool(
        description = "Describe a workbook's structure including sheet names, dimensions, and sample data"
    )]
    async fn describe_workbook(&self, Parameters(i): Parameters<DescribeWorkbookInput>) -> String {
        tool_fn!(self.store, tools::sheets::describe_workbook, i)
    }

    #[tool(description = "Add a new empty worksheet to a workbook")]
    async fn add_sheet(&self, Parameters(i): Parameters<AddSheetInput>) -> String {
        tool_fn!(self.store, tools::sheets::add_sheet, i)
    }

    #[tool(description = "Rename an existing worksheet")]
    async fn rename_sheet(&self, Parameters(i): Parameters<RenameSheetInput>) -> String {
        tool_fn!(self.store, tools::sheets::rename_sheet, i)
    }

    #[tool(description = "Delete a worksheet from a workbook")]
    async fn delete_sheet(&self, Parameters(i): Parameters<DeleteSheetInput>) -> String {
        tool_fn!(self.store, tools::sheets::delete_sheet, i)
    }

    #[tool(description = "Move a worksheet to a different position in the workbook")]
    async fn move_worksheet(&self, Parameters(i): Parameters<MoveWorksheetInput>) -> String {
        tool_fn!(self.store, tools::expanded::move_worksheet, i)
    }

    // ── Read (4) ──

    #[tool(description = "Read data from a worksheet with optional range and pagination")]
    async fn read_sheet(&self, Parameters(i): Parameters<ReadSheetInput>) -> String {
        tool_fn!(self.store, tools::read::read_sheet, i)
    }

    #[tool(description = "Read a single cell's value, type, and formula")]
    async fn read_cell(&self, Parameters(i): Parameters<ReadCellInput>) -> String {
        tool_fn!(self.store, tools::read::read_cell, i)
    }

    #[tool(description = "Search for cells matching a value or pattern across sheets")]
    async fn search_cells(&self, Parameters(i): Parameters<SearchCellsInput>) -> String {
        tool_fn!(self.store, tools::read::search_cells, i)
    }

    #[tool(description = "Export a sheet as a CSV-formatted string")]
    async fn sheet_to_csv(&self, Parameters(i): Parameters<SheetToCsvInput>) -> String {
        tool_fn!(self.store, tools::read::sheet_to_csv, i)
    }

    // ── Write (4) ──

    #[tool(
        description = "Write values to multiple cells. Strings starting with '=' are formulas. Numbers, booleans, ISO dates auto-detected."
    )]
    async fn write_cells(&self, Parameters(i): Parameters<WriteCellsInput>) -> String {
        tool_fn!(self.store, tools::write::write_cells, i)
    }

    #[tool(description = "Write a row of values starting from a cell, filling rightward")]
    async fn write_row(&self, Parameters(i): Parameters<WriteRowInput>) -> String {
        tool_fn!(self.store, tools::write::write_row, i)
    }

    #[tool(description = "Write a column of values starting from a cell, filling downward")]
    async fn write_column(&self, Parameters(i): Parameters<WriteColumnInput>) -> String {
        tool_fn!(self.store, tools::write::write_column, i)
    }

    #[tool(description = "Write rich text (mixed bold/italic/color runs) to a cell")]
    async fn write_rich_text(&self, Parameters(i): Parameters<WriteRichTextInput>) -> String {
        tool_fn!(self.store, tools::expanded::write_rich_text, i)
    }

    // ── Formulas (1) ──

    #[tool(
        description = "Write a formula. Set formula_type to 'array' for CSE array formulas (cell = range), 'dynamic' for Excel 365 spill formulas, or omit for regular. Optional cached_result."
    )]
    async fn write_formula(
        &self,
        Parameters(i): Parameters<WriteFormulaConsolidatedInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::write_formula_consolidated, i)
    }

    // ── Cell operations (1) ──

    #[tool(
        description = "Manage a cell: action='blank' writes a formatted blank cell, action='clear' removes content and formatting"
    )]
    async fn manage_cell(&self, Parameters(i): Parameters<ManageCellInput>) -> String {
        tool_fn!(self.store, tools::expanded::manage_cell, i)
    }

    // ── Formatting (2) ──

    #[tool(
        description = "Apply formatting (bold, italic, colors, borders, number format, alignment) to a range of cells"
    )]
    async fn set_cell_format(&self, Parameters(i): Parameters<SetCellFormatInput>) -> String {
        tool_fn!(self.store, tools::format::set_cell_format, i)
    }

    #[tool(description = "Merge a range of cells into a single cell")]
    async fn merge_cells(&self, Parameters(i): Parameters<MergeCellsInput>) -> String {
        tool_fn!(self.store, tools::format::merge_cells, i)
    }

    // ── Row/column format (1) ──

    #[tool(
        description = "Apply formatting to an entire row or column. Set target='row' with identifier='5' or target='column' with identifier='B'."
    )]
    async fn set_row_column_format(
        &self,
        Parameters(i): Parameters<SetRowColumnFormatInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::set_row_column_format, i)
    }

    // ── Dimensions (1) ──

    #[tool(
        description = "Set dimensions: target='column_width' (column+value), 'row_height' (row+value), 'column_range_width' (first_column+last_column+value), 'default_row_height' (value)"
    )]
    async fn set_dimensions(&self, Parameters(i): Parameters<SetDimensionsInput>) -> String {
        tool_fn!(self.store, tools::expanded::set_dimensions, i)
    }

    // ── Layout (4) ──

    #[tool(
        description = "Freeze panes at a cell position so rows above and columns left remain visible while scrolling"
    )]
    async fn freeze_panes(&self, Parameters(i): Parameters<FreezePanesInput>) -> String {
        tool_fn!(self.store, tools::layout::freeze_panes, i)
    }

    #[tool(description = "Auto-fit all column widths based on cell content")]
    async fn autofit_columns(&self, Parameters(i): Parameters<AutofitColumnsInput>) -> String {
        tool_fn!(self.store, tools::expanded::autofit_columns, i)
    }

    #[tool(description = "Set the selected/active cell in a sheet")]
    async fn set_selection(&self, Parameters(i): Parameters<SetSelectionInput>) -> String {
        tool_fn!(self.store, tools::expanded::set_selection, i)
    }

    #[tool(
        description = "Hide or unhide a row or column. Set target='row' or 'column', identifier='5' or 'B'."
    )]
    async fn set_visibility(&self, Parameters(i): Parameters<SetVisibilityInput>) -> String {
        tool_fn!(self.store, tools::expanded::set_visibility, i)
    }

    // ── Sheet settings (1) ──

    #[tool(
        description = "Configure sheet display: hidden, very_hidden, zoom, hide_gridlines, hide_headings, tab_color, right_to_left"
    )]
    async fn set_sheet_settings(&self, Parameters(i): Parameters<SetSheetSettingsInput>) -> String {
        tool_fn!(self.store, tools::expanded::set_sheet_settings, i)
    }

    // ── Charts (1) ──

    #[tool(
        description = "Add a chart with full control: type (bar/column/line/pie/scatter/area/doughnut), multiple series with colors/labels/trendlines/markers, pivot source, position cell"
    )]
    async fn add_chart(&self, Parameters(i): Parameters<AddChartEnhancedInput>) -> String {
        tool_fn!(self.store, tools::expanded::add_chart_enhanced, i)
    }

    // ── Tables, CF, DV, Sparklines (4) ──

    #[tool(description = "Create an Excel Table with headers, autofilter, and style")]
    async fn add_table(&self, Parameters(i): Parameters<AddTableInput>) -> String {
        tool_fn!(self.store, tools::tables::add_table, i)
    }

    #[tool(
        description = "Add conditional formatting: cell_value, color_scale_2, color_scale_3, data_bar, or icon_set"
    )]
    async fn add_conditional_format(
        &self,
        Parameters(i): Parameters<AddConditionalFormatInput>,
    ) -> String {
        tool_fn!(self.store, tools::conditional::add_conditional_format, i)
    }

    #[tool(
        description = "Add data validation: list, list_range, whole_number, decimal, date_range, text_length, or custom_formula"
    )]
    async fn add_data_validation(
        &self,
        Parameters(i): Parameters<AddDataValidationInput>,
    ) -> String {
        tool_fn!(self.store, tools::validation::add_data_validation, i)
    }

    #[tool(description = "Add a sparkline (line, column, or win/loss) to a cell")]
    async fn add_sparkline(&self, Parameters(i): Parameters<AddSparklineInput>) -> String {
        tool_fn!(self.store, tools::sparklines::add_sparkline, i)
    }

    // ── Images (1) ──

    #[tool(description = "Embed a PNG or JPEG image into a worksheet at a specified cell")]
    async fn add_image(&self, Parameters(i): Parameters<AddImageInput>) -> String {
        tool_fn!(self.store, tools::images::add_image, i)
    }

    // ── Pivot tables (1) ──

    #[tool(
        description = "Create a pivot table with row/column/value/filter fields, aggregation (sum/count/average/max/min), layout (compact/outline/tabular)"
    )]
    async fn add_pivot_table(&self, Parameters(i): Parameters<AddPivotTableInput>) -> String {
        tool_fn!(self.store, tools::expanded::add_pivot_table, i)
    }

    // ── Page setup (1) ──

    #[tool(
        description = "Configure page setup: landscape, paper_size, margins, fit_to_pages, print_scale, print_area, repeat_rows, header, footer, print_gridlines, center, page_breaks"
    )]
    async fn set_page_setup(&self, Parameters(i): Parameters<SetPageSetupInput>) -> String {
        tool_fn!(self.store, tools::expanded::set_page_setup, i)
    }

    // ── Comments & links (2) ──

    #[tool(
        description = "Manage comments: action='add' (cell, text, author) or action='read' to list all comments"
    )]
    async fn manage_comments(&self, Parameters(i): Parameters<ManageCommentsInput>) -> String {
        tool_fn!(self.store, tools::expanded::manage_comments, i)
    }

    #[tool(
        description = "Add a link: link_type='url' for external URLs, link_type='internal' for sheet references (e.g. 'Sheet2!A1')"
    )]
    async fn add_link(&self, Parameters(i): Parameters<AddLinkInput>) -> String {
        tool_fn!(self.store, tools::expanded::add_link, i)
    }

    // ── Named ranges (1) ──

    #[tool(
        description = "Manage defined names: action='add' (name, formula) or action='list' to show all"
    )]
    async fn manage_defined_names(
        &self,
        Parameters(i): Parameters<ManageDefinedNamesInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::manage_defined_names, i)
    }

    // ── Row/column manipulation (2) ──

    #[tool(
        description = "Insert or delete rows: action='insert' or 'delete', at_row (1-based), count"
    )]
    async fn modify_rows(&self, Parameters(i): Parameters<ModifyRowsInput>) -> String {
        tool_fn!(self.store, tools::expanded::modify_rows, i)
    }

    #[tool(
        description = "Insert or delete columns: action='insert' or 'delete', at_column (letter), count"
    )]
    async fn modify_columns(&self, Parameters(i): Parameters<ModifyColumnsInput>) -> String {
        tool_fn!(self.store, tools::expanded::modify_columns, i)
    }

    // ── Grouping (1) ──

    #[tool(
        description = "Group rows or columns into expandable outlines. target='rows' (start/end as numbers) or 'columns' (start/end as letters)"
    )]
    async fn group(&self, Parameters(i): Parameters<GroupInput>) -> String {
        tool_fn!(self.store, tools::expanded::group_consolidated, i)
    }

    // ── Protection (1) ──

    #[tool(
        description = "Protect: target='sheet' (sheet_name, password), 'workbook' (password), or 'unprotect_range' (sheet_name, range, range_title, password)"
    )]
    async fn protect(&self, Parameters(i): Parameters<ProtectInput>) -> String {
        tool_fn!(self.store, tools::expanded::protect_consolidated, i)
    }

    // ── Autofilter (1) ──

    #[tool(
        description = "Set autofilter on a range. Optionally filter a specific column with filter_column and filter_values."
    )]
    async fn manage_autofilter(&self, Parameters(i): Parameters<ManageAutofilterInput>) -> String {
        tool_fn!(self.store, tools::expanded::manage_autofilter, i)
    }

    // ── Error suppression (1) ──

    #[tool(
        description = "Suppress Excel error indicators (green triangles) on a range for a specific error type"
    )]
    async fn ignore_error(&self, Parameters(i): Parameters<IgnoreErrorInput>) -> String {
        tool_fn!(self.store, tools::expanded::ignore_error, i)
    }

    // ── ChartEx charts (3) ──

    #[tool(
        description = "Add a waterfall chart (Excel 2016+ ChartEx). Points have category, value, and point_type (increase/decrease/total)."
    )]
    async fn add_waterfall_chart(
        &self,
        Parameters(i): Parameters<AddWaterfallChartInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::add_waterfall_chart, i)
    }

    #[tool(
        description = "Add a funnel chart (Excel 2016+ ChartEx). Points have category and value, rendered as a narrowing funnel."
    )]
    async fn add_funnel_chart(&self, Parameters(i): Parameters<AddFunnelChartInput>) -> String {
        tool_fn!(self.store, tools::expanded::add_funnel_chart, i)
    }

    #[tool(
        description = "Add a treemap chart (Excel 2016+ ChartEx). Points have category, value, and optional color."
    )]
    async fn add_treemap_chart(&self, Parameters(i): Parameters<AddTreemapChartInput>) -> String {
        tool_fn!(self.store, tools::expanded::add_treemap_chart, i)
    }

    // ── Shapes (1) ──

    #[tool(
        description = "Add a drawing shape (rectangle, rounded_rectangle, ellipse, triangle, diamond, arrow, callout, text_box) with optional text, fill, outline, and font settings."
    )]
    async fn add_shape(&self, Parameters(i): Parameters<AddShapeInput>) -> String {
        tool_fn!(self.store, tools::expanded::add_shape, i)
    }

    // ── Document properties (1) ──

    #[tool(
        description = "Set document properties: title, author, subject, description, keywords, category, company"
    )]
    async fn set_doc_properties(&self, Parameters(i): Parameters<SetDocPropertiesInput>) -> String {
        tool_fn!(self.store, tools::expanded::set_doc_properties, i)
    }

    // ── New chart types (4) ──

    #[tool(
        description = "Add a sunburst chart (Excel 2016+ ChartEx). Hierarchical data with category labels and values."
    )]
    async fn add_sunburst_chart(
        &self,
        Parameters(i): Parameters<AddSunburstChartInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::add_sunburst_chart, i)
    }

    #[tool(
        description = "Add a histogram chart (Excel 2016+ ChartEx). Raw data values with optional bin_count, bin_width, and pareto overlay."
    )]
    async fn add_histogram_chart(
        &self,
        Parameters(i): Parameters<AddHistogramChartInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::add_histogram_chart, i)
    }

    #[tool(
        description = "Add a box & whisker chart (Excel 2016+ ChartEx). Data sets with outliers, mean markers, and inner points options."
    )]
    async fn add_box_whisker_chart(
        &self,
        Parameters(i): Parameters<AddBoxWhiskerChartInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::add_box_whisker_chart, i)
    }

    #[tool(
        description = "Add a map chart (Excel 2016+ ChartEx). Geographic data with location names and values. map_level: 'country' or 'region'."
    )]
    async fn add_map_chart(&self, Parameters(i): Parameters<AddMapChartInput>) -> String {
        tool_fn!(self.store, tools::expanded::add_map_chart, i)
    }

    // ── Interactive controls (3) ──

    #[tool(
        description = "Add a slicer — interactive filter for a pivot table. Specify pivot_table_name and field_name."
    )]
    async fn add_slicer(&self, Parameters(i): Parameters<AddSlicerInput>) -> String {
        tool_fn!(self.store, tools::expanded::add_slicer, i)
    }

    #[tool(
        description = "Add a timeline — date filter for a pivot table. Specify pivot_table_name and date field_name."
    )]
    async fn add_timeline(&self, Parameters(i): Parameters<AddTimelineInput>) -> String {
        tool_fn!(self.store, tools::expanded::add_timeline, i)
    }

    #[tool(
        description = "Add a form control: control_type='button', 'checkbox', 'dropdown', or 'spinner'. Checkbox supports cell_link, dropdown takes comma-separated input_range."
    )]
    async fn add_form_control(&self, Parameters(i): Parameters<AddFormControlInput>) -> String {
        tool_fn!(self.store, tools::expanded::add_form_control, i)
    }

    // ── Advanced save/open (2) ──

    #[tool(
        description = "Save workbook in different formats: 'xlsx' (default), 'template' (.xltx), 'encrypted' (password-protected), 'parallel' (fast compression). Formulas are recalculated before save."
    )]
    async fn save_workbook_advanced(
        &self,
        Parameters(i): Parameters<SaveWorkbookAdvancedInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::save_workbook_advanced, i)
    }

    #[tool(description = "Open a password-protected (encrypted) Excel workbook")]
    async fn open_workbook_encrypted(
        &self,
        Parameters(i): Parameters<OpenWorkbookEncryptedInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::open_workbook_encrypted, i)
    }

    // ── Named ranges CRUD (1) ──

    #[tool(
        description = "Manage named ranges: action='add' (name, formula), 'add_scoped' (name, formula, sheet_index), 'update' (name, formula), 'remove' (name, sheet_index), 'list'"
    )]
    async fn manage_named_ranges(
        &self,
        Parameters(i): Parameters<ManageNamedRangesInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::manage_named_ranges, i)
    }

    // ── Sheet metadata (1) ──

    #[tool(
        description = "Read sheet metadata: info='used_range', 'hyperlinks', 'merge_ranges', 'charts', or 'all'"
    )]
    async fn read_sheet_metadata(
        &self,
        Parameters(i): Parameters<ReadSheetMetadataInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::read_sheet_metadata, i)
    }

    // ── Chart sheet (1) ──

    #[tool(
        description = "Add a dedicated chart-only sheet (no cells, just a chart). Specify chart_type, series or data_range, title."
    )]
    async fn add_chart_sheet(&self, Parameters(i): Parameters<AddChartSheetInput>) -> String {
        tool_fn!(self.store, tools::expanded::add_chart_sheet, i)
    }

    // ── Threaded comments (1) ──

    #[tool(
        description = "Add a modern threaded comment with author, text, optional timestamp, and replies. Unlike legacy comments, these support conversation threads."
    )]
    async fn add_threaded_comment(
        &self,
        Parameters(i): Parameters<AddThreadedCommentInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::add_threaded_comment, i)
    }

    // ── Granular protection (1) ──

    #[tool(
        description = "Protect a sheet with granular options: allow_insert_rows, allow_delete_rows, allow_format_cells, allow_sort, etc. Optional password."
    )]
    async fn protect_sheet_advanced(
        &self,
        Parameters(i): Parameters<ProtectSheetAdvancedInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::protect_sheet_advanced, i)
    }

    // ── Custom properties (1) ──

    #[tool(
        description = "Set a custom document property. value_type: 'text' (default), 'number', 'integer', 'bool', 'datetime'."
    )]
    async fn set_custom_property(
        &self,
        Parameters(i): Parameters<SetCustomPropertyInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::set_custom_property, i)
    }

    // ── Read enhancements (2) ──

    #[tool(description = "Read a single cell's comment (author and text). Returns null if no comment exists.")]
    async fn read_cell_comment(
        &self,
        Parameters(i): Parameters<ReadCellCommentInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::read_cell_comment, i)
    }

    #[tool(description = "Read a cell's format (bold, italic, colors, number format, etc.)")]
    async fn read_cell_format(
        &self,
        Parameters(i): Parameters<ReadCellFormatInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::read_cell_format, i)
    }

    // ── Custom XML (1) ──

    #[tool(
        description = "Manage custom XML parts: action='add' (namespace, content) or action='read' (namespace)"
    )]
    async fn manage_custom_xml(
        &self,
        Parameters(i): Parameters<ManageCustomXmlInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::manage_custom_xml, i)
    }

    // ── External connections (1) ──

    #[tool(description = "Add an external data connection (connection_string, command)")]
    async fn add_connection(&self, Parameters(i): Parameters<AddConnectionInput>) -> String {
        tool_fn!(self.store, tools::expanded::add_connection, i)
    }

    // ── SST optimization (1) ──

    #[tool(description = "Set the shared string table threshold for optimization. Lower values use more memory but faster writes.")]
    async fn set_sst_threshold(
        &self,
        Parameters(i): Parameters<SetSstThresholdInput>,
    ) -> String {
        tool_fn!(self.store, tools::expanded::set_sst_threshold, i)
    }

    // ── JSON rows (1) ──

    #[tool(
        description = "Write JSON objects as rows. Each object's keys become column headers (if write_headers=true), values become cells. Auto-detects types."
    )]
    async fn write_json_rows(&self, Parameters(i): Parameters<WriteJsonRowsInput>) -> String {
        tool_fn!(self.store, tools::expanded::write_json_rows, i)
    }
}

#[tool_handler]
impl ServerHandler for ExcelMcpServer {
    fn get_info(&self) -> ServerInfo {
        ServerInfo::new(ServerCapabilities::builder().enable_tools().build()).with_instructions(
            "Excel file manipulation server powered by zavora-xlsx. 66 tools covering: \
                 workbook lifecycle, sheet management, cell reading/writing, formatting, \
                 charts (11 types + pivot charts + waterfall/funnel/treemap/sunburst/histogram/box-whisker/map \
                 with data tables, 3D views, error bars, axis formatting, drop/high-low lines, gradients), \
                 images, shapes, tables, conditional formatting, data validation, sparklines, \
                 pivot tables (calculated fields, date/range grouping, subtotals, grand totals, value formats), \
                 slicers, timelines, form controls, threaded comments, page setup, hyperlinks, \
                 named ranges (CRUD with scoping), row/column manipulation, grouping, \
                 protection (basic + granular with per-feature allow/deny), autofilter, \
                 formulas (regular/array/dynamic with recalculation), rich text, \
                 document properties, custom properties, sheet metadata, chart sheets, \
                 encrypted open/save, template save, parallel save, and CSV export."
                .to_string(),
        )
    }
}

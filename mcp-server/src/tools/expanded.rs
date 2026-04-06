//! New tools: page setup, comments, hyperlinks, defined names, sheet settings,
//! active sheet, insert/delete rows/cols, grouping, protection, autofit,
//! enhanced charts, pivot tables, read comments, rich text.

use super::common::workbook_not_found;
use crate::store::WorkbookStore;
use crate::types::enums::{ChartType, LegendPosition};
use crate::types::inputs::*;
use crate::types::responses::*;

fn find_sheet(wb: &zavora_xlsx::Workbook, name: &str) -> Option<usize> { wb.sheet_names().iter().position(|n| *n == name) }
fn sheet_err(name: &str) -> String { error(ErrorCategory::NotFound, &format!("Sheet '{}' not found", name), "Check sheet name.") }

// ── Batch 1 ──

pub fn set_page_setup(store: &mut WorkbookStore, input: SetPageSetupInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    if input.landscape == Some(true) { ws.set_landscape(); } else if input.landscape == Some(false) { ws.set_portrait(); }
    if let Some(ps) = input.paper_size { ws.set_paper_size(ps); }
    if let Some(ref m) = input.margins { ws.set_margins(m.top.unwrap_or(0.75), m.bottom.unwrap_or(0.75), m.left.unwrap_or(0.7), m.right.unwrap_or(0.7)); }
    if let Some(ref f) = input.fit_to_pages { ws.set_fit_to_page(f.width, f.height); }
    if let Some(s) = input.print_scale { ws.set_print_scale(s); }
    if let Some(ref pa) = input.print_area {
        let (r1, c1, r2, c2) = zavora_xlsx::utility::parse_range_ref(pa).map_err(|e| anyhow::anyhow!("{e}"))?;
        ws.set_print_area(r1, c1, r2, c2);
    }
    if let Some(ref rr) = input.repeat_rows { ws.set_repeat_rows(rr.first, rr.last); }
    if let Some(ref h) = input.header { ws.set_header(h); }
    if let Some(ref f) = input.footer { ws.set_footer(f); }
    if let Some(true) = input.print_gridlines { ws.set_print_settings(&zavora_xlsx::PrintSettings::new().print_gridlines(true)); }
    if input.center_horizontally == Some(true) { ws.set_print_settings(&zavora_xlsx::PrintSettings::new().center_horizontally(true)); }
    Ok(success_no_data("Page setup configured"))
}

pub fn add_comment(store: &mut WorkbookStore, input: AddCommentInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    if let Some(ref author) = input.author { ws.add_comment_with_author(row, col, &input.text, author); }
    else { ws.add_comment(row, col, &input.text); }
    Ok(success_no_data(&format!("Comment added at {}", input.cell)))
}

pub fn add_hyperlink(store: &mut WorkbookStore, input: AddHyperlinkInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    let display = input.display_text.as_deref().unwrap_or(&input.url);
    ws.write(row, col, display).map_err(|e| anyhow::anyhow!("{e}"))?;
    ws.write_url(row, col, &input.url, input.display_text.as_deref().unwrap_or(&input.url))?;
    Ok(success_no_data(&format!("Hyperlink added at {}", input.cell)))
}

pub fn add_defined_name(store: &mut WorkbookStore, input: AddDefinedNameInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    entry.data.define_name(&input.name, &input.formula);
    Ok(success_no_data(&format!("Defined name '{}' added", input.name)))
}

pub fn list_defined_names(store: &mut WorkbookStore, input: ListDefinedNamesInput) -> Result<String, anyhow::Error> {
    let entry = match store.get(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let names: Vec<serde_json::Value> = entry.data.defined_names().iter()
        .map(|(n, f)| serde_json::json!({"name": n, "formula": f})).collect();
    Ok(success("Defined names listed", names))
}

pub fn set_sheet_settings(store: &mut WorkbookStore, input: SetSheetSettingsInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    if input.hidden == Some(true) { ws.set_hidden(); }
    if input.very_hidden == Some(true) { ws.set_very_hidden(); }
    if let Some(z) = input.zoom { ws.set_zoom(z); }
    if input.hide_gridlines == Some(true) { ws.hide_gridlines(); }
    if input.hide_headings == Some(true) { ws.hide_headings(); }
    if let Some(ref c) = input.tab_color { ws.set_tab_color(c.as_str()); }
    if input.right_to_left == Some(true) { ws.set_right_to_left(); }
    Ok(success_no_data("Sheet settings updated"))
}

pub fn set_active_sheet(store: &mut WorkbookStore, input: SetActiveSheetInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    entry.data.set_active_sheet(input.sheet_index);
    Ok(success_no_data(&format!("Active sheet set to index {}", input.sheet_index)))
}

// ── Batch 2 ──

pub fn insert_rows(store: &mut WorkbookStore, input: InsertDeleteRowsInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.insert_rows(input.at_row.saturating_sub(1), input.count).map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!("{} rows inserted at row {}", input.count, input.at_row)))
}

pub fn delete_rows(store: &mut WorkbookStore, input: InsertDeleteRowsInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.remove_rows(input.at_row.saturating_sub(1), input.count).map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!("{} rows deleted at row {}", input.count, input.at_row)))
}

pub fn insert_columns(store: &mut WorkbookStore, input: InsertDeleteColumnsInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let col = zavora_xlsx::utility::col_from_letter(&input.at_column).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.insert_columns(col, input.count).map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!("{} columns inserted at {}", input.count, input.at_column)))
}

pub fn delete_columns(store: &mut WorkbookStore, input: InsertDeleteColumnsInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let col = zavora_xlsx::utility::col_from_letter(&input.at_column).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.remove_columns(col, input.count).map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!("{} columns deleted at {}", input.count, input.at_column)))
}

pub fn group_rows(store: &mut WorkbookStore, input: GroupRowsInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.group_rows(input.start.saturating_sub(1), input.end.saturating_sub(1), input.level);
    Ok(success_no_data(&format!("Rows {}-{} grouped at level {}", input.start, input.end, input.level)))
}

pub fn group_columns(store: &mut WorkbookStore, input: GroupColumnsInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let c1 = zavora_xlsx::utility::col_from_letter(&input.start).map_err(|e| anyhow::anyhow!("{e}"))?;
    let c2 = zavora_xlsx::utility::col_from_letter(&input.end).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.group_columns(c1, c2, input.level);
    Ok(success_no_data(&format!("Columns {}-{} grouped at level {}", input.start, input.end, input.level)))
}

pub fn protect_sheet(store: &mut WorkbookStore, input: ProtectSheetInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    if let Some(ref pw) = input.password { ws.protect_with_password(pw); } else { ws.protect(); }
    Ok(success_no_data(&format!("Sheet '{}' protected", input.sheet_name)))
}

pub fn protect_workbook(store: &mut WorkbookStore, input: ProtectWorkbookInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    if let Some(ref pw) = input.password { entry.data.protect_with_password(pw); } else { entry.data.protect(); }
    Ok(success_no_data("Workbook protected"))
}

pub fn autofit_columns(store: &mut WorkbookStore, input: AutofitColumnsInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.autofit().map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data("Columns autofitted"))
}

// ── Batch 3 ──

pub fn add_chart_enhanced(store: &mut WorkbookStore, input: AddChartEnhancedInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let ct = match input.chart_type {
        ChartType::Bar => zavora_xlsx::ChartType::Bar, ChartType::Column => zavora_xlsx::ChartType::Column,
        ChartType::Line => zavora_xlsx::ChartType::Line, ChartType::Pie => zavora_xlsx::ChartType::Pie,
        ChartType::Scatter => zavora_xlsx::ChartType::Scatter, ChartType::Area => zavora_xlsx::ChartType::Area,
        ChartType::Doughnut => zavora_xlsx::ChartType::Doughnut,
    };
    let mut chart = zavora_xlsx::Chart::new(ct);
    if let Some(ref t) = input.title { chart.set_title(t); }
    if let Some(ref x) = input.x_axis_label { chart.set_x_axis_name(x); }
    if let Some(ref y) = input.y_axis_label { chart.set_y_axis_name(y); }
    if let Some(ref lp) = input.legend_position {
        chart.set_legend_position(match lp {
            LegendPosition::Top => zavora_xlsx::LegendPosition::Top, LegendPosition::Bottom => zavora_xlsx::LegendPosition::Bottom,
            LegendPosition::Left => zavora_xlsx::LegendPosition::Left, LegendPosition::Right => zavora_xlsx::LegendPosition::Right,
            LegendPosition::None => zavora_xlsx::LegendPosition::None,
        });
    }
    chart.set_width(input.width); chart.set_height(input.height);
    if let Some(ref ps) = input.pivot_source { chart.set_pivot_source(&ps.pivot_table, &ps.sheet); }
    // Add series
    if !input.series.is_empty() {
        for si in &input.series {
            let s = chart.add_series();
            s.set_values(&si.values);
            if let Some(ref c) = si.categories { s.set_categories(c); }
            if let Some(ref n) = si.name { s.set_name(n); }
            if let Some(ref c) = si.color { s.set_color(c.as_str()); }
            if si.data_labels == Some(true) { s.set_data_labels(true); }
            if si.secondary_axis == Some(true) { s.set_secondary_axis(true); }
            if let Some(ref t) = si.trendline {
                let tt = match t.as_str() {
                    "exponential" => zavora_xlsx::TrendlineType::Exponential,
                    "power" => zavora_xlsx::TrendlineType::Power,
                    "logarithmic" => zavora_xlsx::TrendlineType::Logarithmic,
                    _ => zavora_xlsx::TrendlineType::Linear,
                };
                s.set_trendline(tt);
            }
            if let Some(ref m) = si.marker {
                let mt = match m.as_str() {
                    "circle" => zavora_xlsx::MarkerType::Circle, "diamond" => zavora_xlsx::MarkerType::Diamond,
                    "square" => zavora_xlsx::MarkerType::Square, "triangle" => zavora_xlsx::MarkerType::Triangle,
                    _ => zavora_xlsx::MarkerType::None,
                };
                s.set_marker(mt);
            }
        }
    } else if let Some(ref dr) = input.data_range {
        chart.add_series().set_values(dr);
    }
    let (row, col) = if let Some(ref c) = input.cell {
        zavora_xlsx::utility::parse_cell_ref(c).map_err(|e| anyhow::anyhow!("{e}"))?
    } else { (0, 0) };
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.insert_chart(row, col, &chart)?;
    Ok(success_no_data(&format!("Chart added to '{}'", input.sheet_name)))
}

pub fn add_pivot_table(store: &mut WorkbookStore, input: AddPivotTableInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let mut pt = zavora_xlsx::PivotTable::new(&input.name, &input.source_range);
    for f in &input.row_fields { pt = pt.add_row_field(f); }
    for f in &input.column_fields { pt = pt.add_column_field(f); }
    for vf in &input.value_fields {
        let agg = match vf.aggregation.as_str() {
            "count" => zavora_xlsx::PivotAggregation::Count, "average" => zavora_xlsx::PivotAggregation::Average,
            "max" => zavora_xlsx::PivotAggregation::Max, "min" => zavora_xlsx::PivotAggregation::Min,
            "product" => zavora_xlsx::PivotAggregation::Product, _ => zavora_xlsx::PivotAggregation::Sum,
        };
        pt = pt.add_value_field(&vf.field, agg);
    }
    for f in &input.filter_fields { pt = pt.add_filter_field(f); }
    if let Some(ref s) = input.style { pt = pt.set_style_name(s); }
    if let Some(ref l) = input.layout {
        let layout = match l.as_str() {
            "outline" => zavora_xlsx::PivotLayout::Outline, "tabular" => zavora_xlsx::PivotLayout::Tabular,
            _ => zavora_xlsx::PivotLayout::Compact,
        };
        pt = pt.set_layout(layout);
    }
    let (row, col) = if let Some(ref c) = input.cell {
        zavora_xlsx::utility::parse_cell_ref(c).map_err(|e| anyhow::anyhow!("{e}"))?
    } else { (0, 0) };
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.add_pivot_table(row, col, &pt)?;
    Ok(success_no_data(&format!("Pivot table '{}' added", input.name)))
}

// ── Batch 4 ──

pub fn read_comments(store: &mut WorkbookStore, input: ReadCommentsInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    let comments: Vec<serde_json::Value> = ws.comments().iter().map(|c| {
        serde_json::json!({ "cell": zavora_xlsx::utility::to_a1(c.row, c.col), "author": c.author, "text": c.text })
    }).collect();
    Ok(success("Comments read", comments))
}

pub fn write_rich_text(store: &mut WorkbookStore, input: WriteRichTextInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let mut rt = zavora_xlsx::RichText::new();
    for run in &input.runs {
        let mut built = zavora_xlsx::RichTextRun {
            text: run.text.clone(), bold: run.bold.unwrap_or(false), italic: run.italic.unwrap_or(false),
            font_size: run.font_size, font_name: None, color: None, superscript: false, subscript: false,
        };
        if let Some(ref c) = run.color { built.color = Some(c.clone()); }
        rt.runs.push(built);
    }
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.write_rich_text(row, col, &rt).map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!("Rich text written to {}", input.cell)))
}

// ══════════════════════════════════════════════════════════════════
// Batch 5–8: Remaining 22 tools
// ══════════════════════════════════════════════════════════════════

fn build_format(bold: Option<bool>, italic: Option<bool>, font_size: Option<f64>, font_color: Option<&str>, bg: Option<&str>, nf: Option<&str>) -> zavora_xlsx::Format {
    let mut f = zavora_xlsx::Format::new();
    if bold == Some(true) { f = f.bold(); }
    if italic == Some(true) { f = f.italic(); }
    if let Some(s) = font_size { f = f.font_size(s); }
    if let Some(c) = font_color { f = f.font_color(c); }
    if let Some(c) = bg { f = f.background_color(c); }
    if let Some(n) = nf { f = f.num_format(n); }
    f
}

pub fn set_column_format(store: &mut WorkbookStore, input: SetColumnFormatInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let col = zavora_xlsx::utility::col_from_letter(&input.column).map_err(|e| anyhow::anyhow!("{e}"))?;
    let fmt = build_format(input.bold, input.italic, input.font_size, input.font_color.as_deref(), input.background_color.as_deref(), input.number_format.as_deref());
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.set_column_format(col, &fmt);
    Ok(success_no_data(&format!("Column {} format set", input.column)))
}

pub fn set_row_format(store: &mut WorkbookStore, input: SetRowFormatInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let fmt = build_format(input.bold, input.italic, input.font_size, input.font_color.as_deref(), input.background_color.as_deref(), input.number_format.as_deref());
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.set_row_format(input.row.saturating_sub(1), &fmt);
    Ok(success_no_data(&format!("Row {} format set", input.row)))
}

pub fn set_column_hidden(store: &mut WorkbookStore, input: SetColumnHiddenInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let col = zavora_xlsx::utility::col_from_letter(&input.column).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.set_column_hidden(col, input.hidden);
    Ok(success_no_data(&format!("Column {} hidden={}", input.column, input.hidden)))
}

pub fn set_row_hidden(store: &mut WorkbookStore, input: SetRowHiddenInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.set_row_hidden(input.row.saturating_sub(1), input.hidden);
    Ok(success_no_data(&format!("Row {} hidden={}", input.row, input.hidden)))
}

pub fn set_column_range_width(store: &mut WorkbookStore, input: SetColumnRangeWidthInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let c1 = zavora_xlsx::utility::col_from_letter(&input.first_column).map_err(|e| anyhow::anyhow!("{e}"))?;
    let c2 = zavora_xlsx::utility::col_from_letter(&input.last_column).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.set_column_range_width(c1, c2, input.width);
    Ok(success_no_data(&format!("Columns {}:{} width set to {}", input.first_column, input.last_column, input.width)))
}

pub fn set_default_row_height(store: &mut WorkbookStore, input: SetDefaultRowHeightInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.set_default_row_height(input.height);
    Ok(success_no_data(&format!("Default row height set to {}", input.height)))
}

pub fn set_selection(store: &mut WorkbookStore, input: SetSelectionInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.set_selection(row, col);
    Ok(success_no_data(&format!("Selection set to {}", input.cell)))
}

pub fn set_autofilter(store: &mut WorkbookStore, input: SetAutofilterInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (r1, c1, r2, c2) = zavora_xlsx::utility::parse_range_ref(&input.range).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.set_autofilter(r1, c1, r2, c2);
    Ok(success_no_data(&format!("Autofilter set on {}", input.range)))
}

pub fn filter_column(store: &mut WorkbookStore, input: FilterColumnInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let col = zavora_xlsx::utility::col_from_letter(&input.column).map_err(|e| anyhow::anyhow!("{e}"))?;
    let strs: Vec<&str> = input.values.iter().map(|s| s.as_str()).collect();
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.filter_column(col, &strs);
    Ok(success_no_data(&format!("Filter applied to column {}", input.column)))
}

pub fn ignore_error(store: &mut WorkbookStore, input: IgnoreErrorInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.ignore_error(&input.error_type, &input.range);
    Ok(success_no_data("Error ignored"))
}

pub fn set_page_breaks(store: &mut WorkbookStore, input: SetPageBreaksInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.set_page_breaks(&input.row_breaks, &input.col_breaks);
    Ok(success_no_data("Page breaks set"))
}

pub fn unprotect_range(store: &mut WorkbookStore, input: UnprotectRangeInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    if let Some(ref pw) = input.password { ws.unprotect_range_with_password(&input.range, &input.title, pw); }
    else { ws.unprotect_range(&input.range, &input.title); }
    Ok(success_no_data(&format!("Range {} unprotected", input.range)))
}

pub fn write_formula(store: &mut WorkbookStore, input: WriteFormulaInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    if let Some(result) = input.cached_result {
        ws.write_formula_with_result(row, col, &input.formula, result).map_err(|e| anyhow::anyhow!("{e}"))?;
    } else {
        ws.write_formula(row, col, &input.formula).map_err(|e| anyhow::anyhow!("{e}"))?;
    }
    Ok(success_no_data(&format!("Formula written to {}", input.cell)))
}

pub fn write_array_formula(store: &mut WorkbookStore, input: WriteArrayFormulaInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (r1, c1, r2, c2) = zavora_xlsx::utility::parse_range_ref(&input.range).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.write_array_formula(r1, c1, r2, c2, &input.formula).map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!("Array formula written to {}", input.range)))
}

pub fn write_dynamic_formula(store: &mut WorkbookStore, input: WriteDynamicFormulaInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.write_dynamic_formula(row, col, &input.formula).map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!("Dynamic formula written to {}", input.cell)))
}

pub fn write_blank(store: &mut WorkbookStore, input: WriteBlankInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let fmt = build_format(input.bold, None, None, None, input.background_color.as_deref(), input.number_format.as_deref());
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.write_blank(row, col, &fmt).map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!("Blank cell written at {}", input.cell)))
}

pub fn clear_cell(store: &mut WorkbookStore, input: ClearCellInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.clear_cell(row, col);
    Ok(success_no_data(&format!("Cell {} cleared", input.cell)))
}

pub fn set_calc_mode(store: &mut WorkbookStore, input: SetCalcModeInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let mode = match input.mode.as_str() {
        "manual" => zavora_xlsx::CalcMode::Manual,
        "auto_no_table" => zavora_xlsx::CalcMode::AutoNoTable,
        _ => zavora_xlsx::CalcMode::Auto,
    };
    entry.data.set_calc_mode(mode);
    Ok(success_no_data(&format!("Calc mode set to {}", input.mode)))
}

pub fn set_properties(store: &mut WorkbookStore, input: SetPropertiesInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let mut props = entry.data.properties().clone();
    if let Some(ref t) = input.title { props.title = Some(t.clone()); }
    if let Some(ref a) = input.author { props.author = Some(a.clone()); }
    if let Some(ref s) = input.subject { props.subject = Some(s.clone()); }
    if let Some(ref c) = input.company { props.company = Some(c.clone()); }
    if let Some(ref d) = input.description { props.description = Some(d.clone()); }
    entry.data.set_properties(props);
    Ok(success_no_data("Document properties set"))
}

pub fn move_worksheet(store: &mut WorkbookStore, input: MoveWorksheetInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    entry.data.move_worksheet(idx, input.to_index).map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!("Sheet '{}' moved to position {}", input.sheet_name, input.to_index)))
}

pub fn write_internal_link(store: &mut WorkbookStore, input: WriteInternalLinkInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.write_internal_link(row, col, &input.location, &input.display_text).map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!("Internal link written at {}", input.cell)))
}

// ══════════════════════════════════════════════════════════════════
// Consolidated tools (replacing multiple separate tools)
// ══════════════════════════════════════════════════════════════════

pub fn configure_workbook(store: &mut WorkbookStore, input: ConfigureWorkbookInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    if let Some(ref m) = input.calc_mode {
        let mode = match m.as_str() { "manual" => zavora_xlsx::CalcMode::Manual, "auto_no_table" => zavora_xlsx::CalcMode::AutoNoTable, _ => zavora_xlsx::CalcMode::Auto };
        entry.data.set_calc_mode(mode);
    }
    if let Some(i) = input.active_sheet { entry.data.set_active_sheet(i); }
    let mut props = entry.data.properties().clone();
    if let Some(ref v) = input.title { props.title = Some(v.clone()); }
    if let Some(ref v) = input.author { props.author = Some(v.clone()); }
    if let Some(ref v) = input.subject { props.subject = Some(v.clone()); }
    if let Some(ref v) = input.company { props.company = Some(v.clone()); }
    if let Some(ref v) = input.description { props.description = Some(v.clone()); }
    entry.data.set_properties(props);
    Ok(success_no_data("Workbook configured"))
}

pub fn modify_rows(store: &mut WorkbookStore, input: ModifyRowsInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    let row = input.at_row.saturating_sub(1);
    match input.action.as_str() {
        "delete" => { ws.remove_rows(row, input.count).map_err(|e| anyhow::anyhow!("{e}"))?; }
        _ => { ws.insert_rows(row, input.count).map_err(|e| anyhow::anyhow!("{e}"))?; }
    }
    Ok(success_no_data(&format!("{} {} rows at row {}", input.action, input.count, input.at_row)))
}

pub fn modify_columns(store: &mut WorkbookStore, input: ModifyColumnsInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let col = zavora_xlsx::utility::col_from_letter(&input.at_column).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    match input.action.as_str() {
        "delete" => { ws.remove_columns(col, input.count).map_err(|e| anyhow::anyhow!("{e}"))?; }
        _ => { ws.insert_columns(col, input.count).map_err(|e| anyhow::anyhow!("{e}"))?; }
    }
    Ok(success_no_data(&format!("{} {} columns at {}", input.action, input.count, input.at_column)))
}

pub fn write_formula_consolidated(store: &mut WorkbookStore, input: WriteFormulaConsolidatedInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    match input.formula_type.as_deref().unwrap_or("regular") {
        "array" => {
            let (r1, c1, r2, c2) = zavora_xlsx::utility::parse_range_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
            ws.write_array_formula(r1, c1, r2, c2, &input.formula).map_err(|e| anyhow::anyhow!("{e}"))?;
        }
        "dynamic" => {
            let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
            ws.write_dynamic_formula(row, col, &input.formula).map_err(|e| anyhow::anyhow!("{e}"))?;
        }
        _ => {
            let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
            if let Some(r) = input.cached_result { ws.write_formula_with_result(row, col, &input.formula, r).map_err(|e| anyhow::anyhow!("{e}"))?; }
            else { ws.write_formula(row, col, &input.formula).map_err(|e| anyhow::anyhow!("{e}"))?; }
        }
    }
    Ok(success_no_data(&format!("Formula written to {}", input.cell)))
}

pub fn manage_cell(store: &mut WorkbookStore, input: ManageCellInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    match input.action.as_str() {
        "clear" => { ws.clear_cell(row, col); }
        _ => {
            let fmt = build_format(None, None, None, None, input.background_color.as_deref(), input.number_format.as_deref());
            ws.write_blank(row, col, &fmt).map_err(|e| anyhow::anyhow!("{e}"))?;
        }
    }
    Ok(success_no_data(&format!("Cell {} {}", input.cell, input.action)))
}

pub fn manage_comments(store: &mut WorkbookStore, input: ManageCommentsInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    match input.action.as_str() {
        "read" => {
            let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
            let comments: Vec<serde_json::Value> = ws.comments().iter().map(|c|
                serde_json::json!({"cell": zavora_xlsx::utility::to_a1(c.row, c.col), "author": c.author, "text": c.text})
            ).collect();
            Ok(success("Comments read", comments))
        }
        _ => {
            let cell = input.cell.as_deref().unwrap_or("A1");
            let text = input.text.as_deref().unwrap_or("");
            let (row, col) = zavora_xlsx::utility::parse_cell_ref(cell).map_err(|e| anyhow::anyhow!("{e}"))?;
            let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
            if let Some(ref a) = input.author { ws.add_comment_with_author(row, col, text, a); }
            else { ws.add_comment(row, col, text); }
            Ok(success_no_data(&format!("Comment added at {cell}")))
        }
    }
}

pub fn manage_defined_names(store: &mut WorkbookStore, input: ManageDefinedNamesInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    match input.action.as_str() {
        "list" => {
            let names: Vec<serde_json::Value> = entry.data.defined_names().iter()
                .map(|(n, f)| serde_json::json!({"name": n, "formula": f})).collect();
            Ok(success("Defined names", names))
        }
        _ => {
            let name = input.name.as_deref().unwrap_or("");
            let formula = input.formula.as_deref().unwrap_or("");
            entry.data.define_name(name, formula);
            Ok(success_no_data(&format!("Defined name '{name}' added")))
        }
    }
}

pub fn add_link(store: &mut WorkbookStore, input: AddLinkInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    match input.link_type.as_str() {
        "internal" => { ws.write_internal_link(row, col, &input.target, input.display_text.as_deref().unwrap_or(&input.target)).map_err(|e| anyhow::anyhow!("{e}"))?; }
        _ => { ws.write_url(row, col, &input.target, input.display_text.as_deref().unwrap_or(&input.target)).map_err(|e| anyhow::anyhow!("{e}"))?; }
    }
    Ok(success_no_data(&format!("Link added at {}", input.cell)))
}

pub fn protect_consolidated(store: &mut WorkbookStore, input: ProtectInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    match input.target.as_str() {
        "workbook" => {
            if let Some(ref pw) = input.password { entry.data.protect_with_password(pw); } else { entry.data.protect(); }
            Ok(success_no_data("Workbook protected"))
        }
        "unprotect_range" => {
            let sn = input.sheet_name.as_deref().unwrap_or("Sheet1");
            let idx = match find_sheet(&entry.data, sn) { Some(i) => i, None => return Ok(sheet_err(sn)) };
            let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
            let range = input.range.as_deref().unwrap_or("A1:A1");
            let title = input.range_title.as_deref().unwrap_or("Range");
            if let Some(ref pw) = input.password { ws.unprotect_range_with_password(range, title, pw); }
            else { ws.unprotect_range(range, title); }
            Ok(success_no_data(&format!("Range {range} unprotected")))
        }
        _ => { // "sheet"
            let sn = input.sheet_name.as_deref().unwrap_or("Sheet1");
            let idx = match find_sheet(&entry.data, sn) { Some(i) => i, None => return Ok(sheet_err(sn)) };
            let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
            if let Some(ref pw) = input.password { ws.protect_with_password(pw); } else { ws.protect(); }
            Ok(success_no_data(&format!("Sheet '{sn}' protected")))
        }
    }
}

pub fn set_dimensions(store: &mut WorkbookStore, input: SetDimensionsInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    match input.target.as_str() {
        "row_height" => { ws.set_row_height(input.row.unwrap_or(1).saturating_sub(1), input.value).map_err(|e| anyhow::anyhow!("{e}"))?; }
        "column_range_width" => {
            let c1 = zavora_xlsx::utility::col_from_letter(input.first_column.as_deref().unwrap_or("A")).map_err(|e| anyhow::anyhow!("{e}"))?;
            let c2 = zavora_xlsx::utility::col_from_letter(input.last_column.as_deref().unwrap_or("A")).map_err(|e| anyhow::anyhow!("{e}"))?;
            ws.set_column_range_width(c1, c2, input.value);
        }
        "default_row_height" => { ws.set_default_row_height(input.value); }
        _ => { // "column_width"
            let col = zavora_xlsx::utility::col_from_letter(input.column.as_deref().unwrap_or("A")).map_err(|e| anyhow::anyhow!("{e}"))?;
            ws.set_column_width(col, input.value).map_err(|e| anyhow::anyhow!("{e}"))?;
        }
    }
    Ok(success_no_data(&format!("{} set to {}", input.target, input.value)))
}

pub fn set_visibility(store: &mut WorkbookStore, input: SetVisibilityInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    match input.target.as_str() {
        "row" => { let r: u32 = input.identifier.parse().unwrap_or(1); ws.set_row_hidden(r.saturating_sub(1), input.hidden); }
        _ => { let col = zavora_xlsx::utility::col_from_letter(&input.identifier).map_err(|e| anyhow::anyhow!("{e}"))?; ws.set_column_hidden(col, input.hidden); }
    }
    Ok(success_no_data(&format!("{} {} hidden={}", input.target, input.identifier, input.hidden)))
}

pub fn set_row_column_format(store: &mut WorkbookStore, input: SetRowColumnFormatInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let fmt = build_format(input.bold, input.italic, input.font_size, input.font_color.as_deref(), input.background_color.as_deref(), input.number_format.as_deref());
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    match input.target.as_str() {
        "row" => { let r: u32 = input.identifier.parse().unwrap_or(1); ws.set_row_format(r.saturating_sub(1), &fmt); }
        _ => { let col = zavora_xlsx::utility::col_from_letter(&input.identifier).map_err(|e| anyhow::anyhow!("{e}"))?; ws.set_column_format(col, &fmt); }
    }
    Ok(success_no_data(&format!("{} {} format set", input.target, input.identifier)))
}

pub fn group_consolidated(store: &mut WorkbookStore, input: GroupInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    match input.target.as_str() {
        "columns" => {
            let c1 = zavora_xlsx::utility::col_from_letter(&input.start).map_err(|e| anyhow::anyhow!("{e}"))?;
            let c2 = zavora_xlsx::utility::col_from_letter(&input.end).map_err(|e| anyhow::anyhow!("{e}"))?;
            ws.group_columns(c1, c2, input.level);
        }
        _ => {
            let s: u32 = input.start.parse().unwrap_or(1);
            let e: u32 = input.end.parse().unwrap_or(1);
            ws.group_rows(s.saturating_sub(1), e.saturating_sub(1), input.level);
        }
    }
    Ok(success_no_data(&format!("{} {}-{} grouped at level {}", input.target, input.start, input.end, input.level)))
}

pub fn manage_autofilter(store: &mut WorkbookStore, input: ManageAutofilterInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) { Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)) };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (r1, c1, r2, c2) = zavora_xlsx::utility::parse_range_ref(&input.range).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    ws.set_autofilter(r1, c1, r2, c2);
    if let (Some(ref fc), Some(ref fv)) = (input.filter_column, input.filter_values) {
        let col = zavora_xlsx::utility::col_from_letter(fc).map_err(|e| anyhow::anyhow!("{e}"))?;
        let strs: Vec<&str> = fv.iter().map(|s| s.as_str()).collect();
        ws.filter_column(col, &strs);
    }
    Ok(success_no_data(&format!("Autofilter set on {}", input.range)))
}

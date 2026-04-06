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

//! New tools: page setup, comments, hyperlinks, defined names, sheet settings,
//! active sheet, insert/delete rows/cols, grouping, protection, autofit,
//! enhanced charts, pivot tables, read comments, rich text.

use super::common::workbook_not_found;
use crate::store::WorkbookStore;
use crate::types::enums::{ChartType, LegendPosition};
use crate::types::inputs::*;
use crate::types::responses::*;

fn find_sheet(wb: &zavora_xlsx::Workbook, name: &str) -> Option<usize> {
    wb.sheet_names().iter().position(|n| *n == name)
}
fn sheet_err(name: &str) -> String {
    error(
        ErrorCategory::NotFound,
        &format!("Sheet '{}' not found", name),
        "Check sheet name.",
    )
}

// ── Batch 1 ──

pub fn set_page_setup(
    store: &mut WorkbookStore,
    input: SetPageSetupInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    if input.landscape == Some(true) {
        ws.set_landscape();
    } else if input.landscape == Some(false) {
        ws.set_portrait();
    }
    if let Some(ps) = input.paper_size {
        ws.set_paper_size(ps);
    }
    if let Some(ref m) = input.margins {
        ws.set_margins(
            m.top.unwrap_or(0.75),
            m.bottom.unwrap_or(0.75),
            m.left.unwrap_or(0.7),
            m.right.unwrap_or(0.7),
        );
    }
    if let Some(ref f) = input.fit_to_pages {
        ws.set_fit_to_page(f.width, f.height);
    }
    if let Some(s) = input.print_scale {
        ws.set_print_scale(s);
    }
    if let Some(ref pa) = input.print_area {
        let (r1, c1, r2, c2) =
            zavora_xlsx::utility::parse_range_ref(pa).map_err(|e| anyhow::anyhow!("{e}"))?;
        ws.set_print_area(r1, c1, r2, c2);
    }
    if let Some(ref rr) = input.repeat_rows {
        ws.set_repeat_rows(rr.first, rr.last);
    }
    if let Some(ref h) = input.header {
        ws.set_header(h);
    }
    if let Some(ref f) = input.footer {
        ws.set_footer(f);
    }
    if let Some(true) = input.print_gridlines {
        ws.set_print_settings(&zavora_xlsx::PrintSettings::new().print_gridlines(true));
    }
    if input.center_horizontally == Some(true) {
        ws.set_print_settings(&zavora_xlsx::PrintSettings::new().center_horizontally(true));
    }
    Ok(success_no_data("Page setup configured"))
}

pub fn add_comment(
    store: &mut WorkbookStore,
    input: AddCommentInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (row, col) =
        zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    if let Some(ref author) = input.author {
        ws.add_comment_with_author(row, col, &input.text, author);
    } else {
        ws.add_comment(row, col, &input.text);
    }
    Ok(success_no_data(&format!("Comment added at {}", input.cell)))
}

pub fn add_hyperlink(
    store: &mut WorkbookStore,
    input: AddHyperlinkInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (row, col) =
        zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    let display = input.display_text.as_deref().unwrap_or(&input.url);
    ws.write(row, col, display)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    ws.write_url(
        row,
        col,
        &input.url,
        input.display_text.as_deref().unwrap_or(&input.url),
    )?;
    Ok(success_no_data(&format!(
        "Hyperlink added at {}",
        input.cell
    )))
}

pub fn add_defined_name(
    store: &mut WorkbookStore,
    input: AddDefinedNameInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    entry.data.define_name(&input.name, &input.formula);
    Ok(success_no_data(&format!(
        "Defined name '{}' added",
        input.name
    )))
}

pub fn list_defined_names(
    store: &mut WorkbookStore,
    input: ListDefinedNamesInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let names: Vec<serde_json::Value> = entry
        .data
        .defined_names()
        .iter()
        .map(|(n, f)| serde_json::json!({"name": n, "formula": f}))
        .collect();
    Ok(success("Defined names listed", names))
}

pub fn set_sheet_settings(
    store: &mut WorkbookStore,
    input: SetSheetSettingsInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    if input.hidden == Some(true) {
        ws.set_hidden();
    }
    if input.very_hidden == Some(true) {
        ws.set_very_hidden();
    }
    if let Some(z) = input.zoom {
        ws.set_zoom(z);
    }
    if input.hide_gridlines == Some(true) {
        ws.hide_gridlines();
    }
    if input.hide_headings == Some(true) {
        ws.hide_headings();
    }
    if let Some(ref c) = input.tab_color {
        ws.set_tab_color(c.as_str());
    }
    if input.right_to_left == Some(true) {
        ws.set_right_to_left();
    }
    Ok(success_no_data("Sheet settings updated"))
}

pub fn set_active_sheet(
    store: &mut WorkbookStore,
    input: SetActiveSheetInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    entry.data.set_active_sheet(input.sheet_index);
    Ok(success_no_data(&format!(
        "Active sheet set to index {}",
        input.sheet_index
    )))
}

// ── Batch 2 ──

pub fn insert_rows(
    store: &mut WorkbookStore,
    input: InsertDeleteRowsInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .insert_rows(input.at_row.saturating_sub(1), input.count)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!(
        "{} rows inserted at row {}",
        input.count, input.at_row
    )))
}

pub fn delete_rows(
    store: &mut WorkbookStore,
    input: InsertDeleteRowsInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .remove_rows(input.at_row.saturating_sub(1), input.count)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!(
        "{} rows deleted at row {}",
        input.count, input.at_row
    )))
}

pub fn insert_columns(
    store: &mut WorkbookStore,
    input: InsertDeleteColumnsInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let col = zavora_xlsx::utility::col_from_letter(&input.at_column)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .insert_columns(col, input.count)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!(
        "{} columns inserted at {}",
        input.count, input.at_column
    )))
}

pub fn delete_columns(
    store: &mut WorkbookStore,
    input: InsertDeleteColumnsInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let col = zavora_xlsx::utility::col_from_letter(&input.at_column)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .remove_columns(col, input.count)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!(
        "{} columns deleted at {}",
        input.count, input.at_column
    )))
}

pub fn group_rows(
    store: &mut WorkbookStore,
    input: GroupRowsInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .group_rows(
            input.start.saturating_sub(1),
            input.end.saturating_sub(1),
            input.level,
        );
    Ok(success_no_data(&format!(
        "Rows {}-{} grouped at level {}",
        input.start, input.end, input.level
    )))
}

pub fn group_columns(
    store: &mut WorkbookStore,
    input: GroupColumnsInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let c1 =
        zavora_xlsx::utility::col_from_letter(&input.start).map_err(|e| anyhow::anyhow!("{e}"))?;
    let c2 =
        zavora_xlsx::utility::col_from_letter(&input.end).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .group_columns(c1, c2, input.level);
    Ok(success_no_data(&format!(
        "Columns {}-{} grouped at level {}",
        input.start, input.end, input.level
    )))
}

pub fn protect_sheet(
    store: &mut WorkbookStore,
    input: ProtectSheetInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    if let Some(ref pw) = input.password {
        ws.protect_with_password(pw);
    } else {
        ws.protect();
    }
    Ok(success_no_data(&format!(
        "Sheet '{}' protected",
        input.sheet_name
    )))
}

pub fn protect_workbook(
    store: &mut WorkbookStore,
    input: ProtectWorkbookInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    if let Some(ref pw) = input.password {
        entry.data.protect_with_password(pw);
    } else {
        entry.data.protect();
    }
    Ok(success_no_data("Workbook protected"))
}

pub fn autofit_columns(
    store: &mut WorkbookStore,
    input: AutofitColumnsInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .autofit()
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data("Columns autofitted"))
}

// ── Batch 3 ──

pub fn add_chart_enhanced(
    store: &mut WorkbookStore,
    input: AddChartEnhancedInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let ct = match input.chart_type {
        ChartType::Bar => zavora_xlsx::ChartType::Bar,
        ChartType::Column => zavora_xlsx::ChartType::Column,
        ChartType::Line => zavora_xlsx::ChartType::Line,
        ChartType::Pie => zavora_xlsx::ChartType::Pie,
        ChartType::Scatter => zavora_xlsx::ChartType::Scatter,
        ChartType::Area => zavora_xlsx::ChartType::Area,
        ChartType::Doughnut => zavora_xlsx::ChartType::Doughnut,
    };
    let mut chart = zavora_xlsx::Chart::new(ct);
    if let Some(ref t) = input.title {
        chart.set_title(t);
    }
    if let Some(ref x) = input.x_axis_label {
        chart.set_x_axis_name(x);
    }
    if let Some(ref y) = input.y_axis_label {
        chart.set_y_axis_name(y);
    }
    if let Some(ref lp) = input.legend_position {
        chart.set_legend_position(match lp {
            LegendPosition::Top => zavora_xlsx::LegendPosition::Top,
            LegendPosition::Bottom => zavora_xlsx::LegendPosition::Bottom,
            LegendPosition::Left => zavora_xlsx::LegendPosition::Left,
            LegendPosition::Right => zavora_xlsx::LegendPosition::Right,
            LegendPosition::None => zavora_xlsx::LegendPosition::None,
        });
    }
    chart.set_width(input.width);
    chart.set_height(input.height);
    if let Some(ref ps) = input.pivot_source {
        chart.set_pivot_source(&ps.pivot_table, &ps.sheet);
    }
    // Chart enhancements
    if input.show_data_table == Some(true) {
        chart.show_data_table(true);
    }
    if let Some(ref v) = input.view_3d {
        let mut view = zavora_xlsx::View3D::default();
        if let Some(rx) = v.rot_x {
            view.rot_x = rx;
        }
        if let Some(ry) = v.rot_y {
            view.rot_y = ry;
        }
        if let Some(p) = v.perspective {
            view.perspective = p;
        }
        chart.set_view3d(view);
    }
    if let Some(s) = input.style {
        chart.set_style(s);
    }
    if let Some(ref at) = input.alt_text {
        chart.set_alt_text(&at.title, &at.description);
    }
    if let Some(v) = input.y_axis_min {
        chart.set_y_axis_min(v);
    }
    if let Some(v) = input.y_axis_max {
        chart.set_y_axis_max(v);
    }
    if let Some(v) = input.y_axis_log_base {
        chart.set_y_axis_log_base(v);
    }
    if input.x_axis_reverse == Some(true) {
        chart.set_x_axis_reverse();
    }
    if input.y_axis_reverse == Some(true) {
        chart.set_y_axis_reverse();
    }
    if let Some(ref fmt) = input.x_axis_format {
        let af = zavora_xlsx::AxisFormat {
            num_format: Some(fmt.clone()),
            ..Default::default()
        };
        chart.set_x_axis_format(af);
    }
    if let Some(ref fmt) = input.y_axis_format {
        let af = zavora_xlsx::AxisFormat {
            num_format: Some(fmt.clone()),
            ..Default::default()
        };
        chart.set_y_axis_format(af);
    }
    if input.drop_lines == Some(true) {
        chart.set_drop_lines(true);
    }
    if input.high_low_lines == Some(true) {
        chart.set_high_low_lines(true);
    }
    if let Some(ref fill) = input.plot_area_fill {
        let rgb = parse_hex_color(fill);
        let paf = zavora_xlsx::PlotAreaFormat {
            fill: Some(rgb),
            border: None,
            gradient: None,
        };
        chart.set_plot_area_format(paf);
    }
    // Add series
    if !input.series.is_empty() {
        for si in &input.series {
            let s = chart.add_series();
            s.set_values(&si.values);
            if let Some(ref c) = si.categories {
                s.set_categories(c);
            }
            if let Some(ref n) = si.name {
                s.set_name(n);
            }
            if let Some(ref c) = si.color {
                s.set_color(c.as_str());
            }
            if si.data_labels == Some(true) {
                s.set_data_labels(true);
            }
            if si.secondary_axis == Some(true) {
                s.set_secondary_axis(true);
            }
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
                    "circle" => zavora_xlsx::MarkerType::Circle,
                    "diamond" => zavora_xlsx::MarkerType::Diamond,
                    "square" => zavora_xlsx::MarkerType::Square,
                    "triangle" => zavora_xlsx::MarkerType::Triangle,
                    _ => zavora_xlsx::MarkerType::None,
                };
                s.set_marker(mt);
            }
            if let Some(lw) = si.line_width {
                s.set_line_width(lw);
            }
            if let Some(ref ds) = si.dash_style {
                let style = match ds.as_str() {
                    "dash" => zavora_xlsx::DashStyle::Dash,
                    "dot" => zavora_xlsx::DashStyle::Dot,
                    "dash_dot" => zavora_xlsx::DashStyle::DashDot,
                    "long_dash" => zavora_xlsx::DashStyle::LongDash,
                    "long_dash_dot" => zavora_xlsx::DashStyle::LongDashDot,
                    _ => zavora_xlsx::DashStyle::Solid,
                };
                s.set_dash_style(style);
            }
            if let Some(ref g) = si.gradient {
                let stops: Vec<([u8; 3], f64)> = g
                    .iter()
                    .map(|gs| (parse_hex_color(&gs.color), gs.position))
                    .collect();
                s.set_gradient(stops);
            }
            if let Some(ref bs) = si.bubble_sizes {
                s.set_bubble_sizes(bs);
            }
            if let Some(ref eb) = si.error_bars {
                let bt = match eb.bar_type.as_str() {
                    "plus" => zavora_xlsx::ErrorBarType::Plus,
                    "minus" => zavora_xlsx::ErrorBarType::Minus,
                    _ => zavora_xlsx::ErrorBarType::Both,
                };
                let vt = match eb.value_type.as_str() {
                    "percentage" => zavora_xlsx::ErrorBarValueType::Percentage,
                    "std_dev" => zavora_xlsx::ErrorBarValueType::StandardDeviation,
                    "std_error" => zavora_xlsx::ErrorBarValueType::StandardError,
                    _ => zavora_xlsx::ErrorBarValueType::FixedValue,
                };
                s.set_error_bars(zavora_xlsx::ErrorBar::new(bt, vt, eb.value));
            }
        }
    } else if let Some(ref dr) = input.data_range {
        chart.add_series().set_values(dr);
    }
    let (row, col) = if let Some(ref c) = input.cell {
        zavora_xlsx::utility::parse_cell_ref(c).map_err(|e| anyhow::anyhow!("{e}"))?
    } else {
        (0, 0)
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .insert_chart(row, col, &chart)?;
    Ok(success_no_data(&format!(
        "Chart added to '{}'",
        input.sheet_name
    )))
}

pub fn add_pivot_table(
    store: &mut WorkbookStore,
    input: AddPivotTableInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let mut pt = zavora_xlsx::PivotTable::new(&input.name, &input.source_range);
    for f in &input.row_fields {
        pt = pt.add_row_field(f);
    }
    for f in &input.column_fields {
        pt = pt.add_column_field(f);
    }
    for vf in &input.value_fields {
        let agg = match vf.aggregation.as_str() {
            "count" => zavora_xlsx::PivotAggregation::Count,
            "average" => zavora_xlsx::PivotAggregation::Average,
            "max" => zavora_xlsx::PivotAggregation::Max,
            "min" => zavora_xlsx::PivotAggregation::Min,
            "product" => zavora_xlsx::PivotAggregation::Product,
            _ => zavora_xlsx::PivotAggregation::Sum,
        };
        pt = pt.add_value_field(&vf.field, agg);
    }
    for f in &input.filter_fields {
        pt = pt.add_filter_field(f);
    }
    if let Some(ref s) = input.style {
        pt = pt.set_style_name(s);
    }
    if let Some(ref l) = input.layout {
        let layout = match l.as_str() {
            "outline" => zavora_xlsx::PivotLayout::Outline,
            "tabular" => zavora_xlsx::PivotLayout::Tabular,
            _ => zavora_xlsx::PivotLayout::Compact,
        };
        pt = pt.set_layout(layout);
    }
    // Pivot enhancements
    for cf in &input.calculated_fields {
        pt = pt.add_calculated_field(&cf.name, &cf.formula);
    }
    for dg in &input.date_groups {
        let levels: Vec<zavora_xlsx::DateGroupLevel> = dg
            .levels
            .iter()
            .map(|l| match l.as_str() {
                "years" => zavora_xlsx::DateGroupLevel::Years,
                "quarters" => zavora_xlsx::DateGroupLevel::Quarters,
                "months" => zavora_xlsx::DateGroupLevel::Months,
                "days" => zavora_xlsx::DateGroupLevel::Days,
                "hours" => zavora_xlsx::DateGroupLevel::Hours,
                "minutes" => zavora_xlsx::DateGroupLevel::Minutes,
                "seconds" => zavora_xlsx::DateGroupLevel::Seconds,
                _ => zavora_xlsx::DateGroupLevel::Months,
            })
            .collect();
        pt = pt.group_by_date(&dg.field, &levels);
    }
    for rg in &input.range_groups {
        pt = pt.group_by_range(&rg.field, rg.start, rg.end, rg.interval);
    }
    for vf in &input.value_formats {
        pt = pt.set_value_format(&vf.field, &vf.format);
    }
    for st in &input.subtotals {
        pt = pt.show_subtotals(&st.field, st.show);
    }
    if let (Some(rows), Some(cols)) = (input.grand_total_rows, input.grand_total_cols) {
        pt = pt.show_grand_totals(rows, cols);
    } else if let Some(rows) = input.grand_total_rows {
        pt = pt.show_grand_totals(rows, true);
    } else if let Some(cols) = input.grand_total_cols {
        pt = pt.show_grand_totals(true, cols);
    }
    if let Some(v) = input.show_row_headers {
        pt = pt.show_row_headers(v);
    }
    if let Some(v) = input.show_column_headers {
        pt = pt.show_column_headers(v);
    }
    if let Some(v) = input.show_row_stripes {
        pt = pt.show_row_stripes(v);
    }
    let (row, col) = if let Some(ref c) = input.cell {
        zavora_xlsx::utility::parse_cell_ref(c).map_err(|e| anyhow::anyhow!("{e}"))?
    } else {
        (0, 0)
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .add_pivot_table(row, col, &pt)?;
    Ok(success_no_data(&format!(
        "Pivot table '{}' added",
        input.name
    )))
}

// ── Batch 4 ──

pub fn read_comments(
    store: &mut WorkbookStore,
    input: ReadCommentsInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    let comments: Vec<serde_json::Value> = ws.comments().iter().map(|c| {
        serde_json::json!({ "cell": zavora_xlsx::utility::to_a1(c.row, c.col), "author": c.author, "text": c.text })
    }).collect();
    Ok(success("Comments read", comments))
}

pub fn write_rich_text(
    store: &mut WorkbookStore,
    input: WriteRichTextInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (row, col) =
        zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let mut rt = zavora_xlsx::RichText::new();
    for run in &input.runs {
        let mut built = zavora_xlsx::RichTextRun {
            text: run.text.clone(),
            bold: run.bold.unwrap_or(false),
            italic: run.italic.unwrap_or(false),
            font_size: run.font_size,
            font_name: None,
            color: None,
            superscript: false,
            subscript: false,
        };
        if let Some(ref c) = run.color {
            built.color = Some(c.clone());
        }
        rt.runs.push(built);
    }
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .write_rich_text(row, col, &rt)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!(
        "Rich text written to {}",
        input.cell
    )))
}

// ══════════════════════════════════════════════════════════════════
// Batch 5–8: Remaining 22 tools
// ══════════════════════════════════════════════════════════════════

fn build_format(
    bold: Option<bool>,
    italic: Option<bool>,
    font_size: Option<f64>,
    font_color: Option<&str>,
    bg: Option<&str>,
    nf: Option<&str>,
) -> zavora_xlsx::Format {
    let mut f = zavora_xlsx::Format::new();
    if bold == Some(true) {
        f = f.bold();
    }
    if italic == Some(true) {
        f = f.italic();
    }
    if let Some(s) = font_size {
        f = f.font_size(s);
    }
    if let Some(c) = font_color {
        f = f.font_color(c);
    }
    if let Some(c) = bg {
        f = f.background_color(c);
    }
    if let Some(n) = nf {
        f = f.num_format(n);
    }
    f
}

pub fn set_column_format(
    store: &mut WorkbookStore,
    input: SetColumnFormatInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let col =
        zavora_xlsx::utility::col_from_letter(&input.column).map_err(|e| anyhow::anyhow!("{e}"))?;
    let fmt = build_format(
        input.bold,
        input.italic,
        input.font_size,
        input.font_color.as_deref(),
        input.background_color.as_deref(),
        input.number_format.as_deref(),
    );
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .set_column_format(col, &fmt);
    Ok(success_no_data(&format!(
        "Column {} format set",
        input.column
    )))
}

pub fn set_row_format(
    store: &mut WorkbookStore,
    input: SetRowFormatInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let fmt = build_format(
        input.bold,
        input.italic,
        input.font_size,
        input.font_color.as_deref(),
        input.background_color.as_deref(),
        input.number_format.as_deref(),
    );
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .set_row_format(input.row.saturating_sub(1), &fmt);
    Ok(success_no_data(&format!("Row {} format set", input.row)))
}

pub fn set_column_hidden(
    store: &mut WorkbookStore,
    input: SetColumnHiddenInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let col =
        zavora_xlsx::utility::col_from_letter(&input.column).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .set_column_hidden(col, input.hidden);
    Ok(success_no_data(&format!(
        "Column {} hidden={}",
        input.column, input.hidden
    )))
}

pub fn set_row_hidden(
    store: &mut WorkbookStore,
    input: SetRowHiddenInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .set_row_hidden(input.row.saturating_sub(1), input.hidden);
    Ok(success_no_data(&format!(
        "Row {} hidden={}",
        input.row, input.hidden
    )))
}

pub fn set_column_range_width(
    store: &mut WorkbookStore,
    input: SetColumnRangeWidthInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let c1 = zavora_xlsx::utility::col_from_letter(&input.first_column)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    let c2 = zavora_xlsx::utility::col_from_letter(&input.last_column)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .set_column_range_width(c1, c2, input.width);
    Ok(success_no_data(&format!(
        "Columns {}:{} width set to {}",
        input.first_column, input.last_column, input.width
    )))
}

pub fn set_default_row_height(
    store: &mut WorkbookStore,
    input: SetDefaultRowHeightInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .set_default_row_height(input.height);
    Ok(success_no_data(&format!(
        "Default row height set to {}",
        input.height
    )))
}

pub fn set_selection(
    store: &mut WorkbookStore,
    input: SetSelectionInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (row, col) =
        zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .set_selection(row, col);
    Ok(success_no_data(&format!("Selection set to {}", input.cell)))
}

pub fn set_autofilter(
    store: &mut WorkbookStore,
    input: SetAutofilterInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (r1, c1, r2, c2) =
        zavora_xlsx::utility::parse_range_ref(&input.range).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .set_autofilter(r1, c1, r2, c2);
    Ok(success_no_data(&format!(
        "Autofilter set on {}",
        input.range
    )))
}

pub fn filter_column(
    store: &mut WorkbookStore,
    input: FilterColumnInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let col =
        zavora_xlsx::utility::col_from_letter(&input.column).map_err(|e| anyhow::anyhow!("{e}"))?;
    let strs: Vec<&str> = input.values.iter().map(|s| s.as_str()).collect();
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .filter_column(col, &strs);
    Ok(success_no_data(&format!(
        "Filter applied to column {}",
        input.column
    )))
}

pub fn ignore_error(
    store: &mut WorkbookStore,
    input: IgnoreErrorInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .ignore_error(&input.error_type, &input.range);
    Ok(success_no_data("Error ignored"))
}

pub fn set_page_breaks(
    store: &mut WorkbookStore,
    input: SetPageBreaksInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .set_page_breaks(&input.row_breaks, &input.col_breaks);
    Ok(success_no_data("Page breaks set"))
}

pub fn unprotect_range(
    store: &mut WorkbookStore,
    input: UnprotectRangeInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    if let Some(ref pw) = input.password {
        ws.unprotect_range_with_password(&input.range, &input.title, pw);
    } else {
        ws.unprotect_range(&input.range, &input.title);
    }
    Ok(success_no_data(&format!(
        "Range {} unprotected",
        input.range
    )))
}

pub fn write_formula(
    store: &mut WorkbookStore,
    input: WriteFormulaInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (row, col) =
        zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    if let Some(result) = input.cached_result {
        ws.write_formula_with_result(row, col, &input.formula, result)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
    } else {
        ws.write_formula(row, col, &input.formula)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
    }
    Ok(success_no_data(&format!(
        "Formula written to {}",
        input.cell
    )))
}

pub fn write_array_formula(
    store: &mut WorkbookStore,
    input: WriteArrayFormulaInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (r1, c1, r2, c2) =
        zavora_xlsx::utility::parse_range_ref(&input.range).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .write_array_formula(r1, c1, r2, c2, &input.formula)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!(
        "Array formula written to {}",
        input.range
    )))
}

pub fn write_dynamic_formula(
    store: &mut WorkbookStore,
    input: WriteDynamicFormulaInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (row, col) =
        zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .write_dynamic_formula(row, col, &input.formula)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!(
        "Dynamic formula written to {}",
        input.cell
    )))
}

pub fn write_blank(
    store: &mut WorkbookStore,
    input: WriteBlankInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (row, col) =
        zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let fmt = build_format(
        input.bold,
        None,
        None,
        None,
        input.background_color.as_deref(),
        input.number_format.as_deref(),
    );
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .write_blank(row, col, &fmt)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!(
        "Blank cell written at {}",
        input.cell
    )))
}

pub fn clear_cell(
    store: &mut WorkbookStore,
    input: ClearCellInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (row, col) =
        zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .clear_cell(row, col);
    Ok(success_no_data(&format!("Cell {} cleared", input.cell)))
}

pub fn set_calc_mode(
    store: &mut WorkbookStore,
    input: SetCalcModeInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let mode = match input.mode.as_str() {
        "manual" => zavora_xlsx::CalcMode::Manual,
        "auto_no_table" => zavora_xlsx::CalcMode::AutoNoTable,
        _ => zavora_xlsx::CalcMode::Auto,
    };
    entry.data.set_calc_mode(mode);
    Ok(success_no_data(&format!("Calc mode set to {}", input.mode)))
}

pub fn set_properties(
    store: &mut WorkbookStore,
    input: SetPropertiesInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let mut props = entry.data.properties().clone();
    if let Some(ref t) = input.title {
        props.title = Some(t.clone());
    }
    if let Some(ref a) = input.author {
        props.author = Some(a.clone());
    }
    if let Some(ref s) = input.subject {
        props.subject = Some(s.clone());
    }
    if let Some(ref c) = input.company {
        props.company = Some(c.clone());
    }
    if let Some(ref d) = input.description {
        props.description = Some(d.clone());
    }
    entry.data.set_properties(props);
    Ok(success_no_data("Document properties set"))
}

pub fn move_worksheet(
    store: &mut WorkbookStore,
    input: MoveWorksheetInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    entry
        .data
        .move_worksheet(idx, input.to_index)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!(
        "Sheet '{}' moved to position {}",
        input.sheet_name, input.to_index
    )))
}

pub fn write_internal_link(
    store: &mut WorkbookStore,
    input: WriteInternalLinkInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (row, col) =
        zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .write_internal_link(row, col, &input.location, &input.display_text)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!(
        "Internal link written at {}",
        input.cell
    )))
}

// ══════════════════════════════════════════════════════════════════
// Consolidated tools (replacing multiple separate tools)
// ══════════════════════════════════════════════════════════════════

pub fn configure_workbook(
    store: &mut WorkbookStore,
    input: ConfigureWorkbookInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    if let Some(ref m) = input.calc_mode {
        let mode = match m.as_str() {
            "manual" => zavora_xlsx::CalcMode::Manual,
            "auto_no_table" => zavora_xlsx::CalcMode::AutoNoTable,
            _ => zavora_xlsx::CalcMode::Auto,
        };
        entry.data.set_calc_mode(mode);
    }
    if let Some(i) = input.active_sheet {
        entry.data.set_active_sheet(i);
    }
    let mut props = entry.data.properties().clone();
    if let Some(ref v) = input.title {
        props.title = Some(v.clone());
    }
    if let Some(ref v) = input.author {
        props.author = Some(v.clone());
    }
    if let Some(ref v) = input.subject {
        props.subject = Some(v.clone());
    }
    if let Some(ref v) = input.company {
        props.company = Some(v.clone());
    }
    if let Some(ref v) = input.description {
        props.description = Some(v.clone());
    }
    entry.data.set_properties(props);
    Ok(success_no_data("Workbook configured"))
}

pub fn modify_rows(
    store: &mut WorkbookStore,
    input: ModifyRowsInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    let row = input.at_row.saturating_sub(1);
    match input.action.as_str() {
        "delete" => {
            ws.remove_rows(row, input.count)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
        }
        _ => {
            ws.insert_rows(row, input.count)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
        }
    }
    Ok(success_no_data(&format!(
        "{} {} rows at row {}",
        input.action, input.count, input.at_row
    )))
}

pub fn modify_columns(
    store: &mut WorkbookStore,
    input: ModifyColumnsInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let col = zavora_xlsx::utility::col_from_letter(&input.at_column)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    match input.action.as_str() {
        "delete" => {
            ws.remove_columns(col, input.count)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
        }
        _ => {
            ws.insert_columns(col, input.count)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
        }
    }
    Ok(success_no_data(&format!(
        "{} {} columns at {}",
        input.action, input.count, input.at_column
    )))
}

pub fn write_formula_consolidated(
    store: &mut WorkbookStore,
    input: WriteFormulaConsolidatedInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    match input.formula_type.as_deref().unwrap_or("regular") {
        "array" => {
            let (r1, c1, r2, c2) = zavora_xlsx::utility::parse_range_ref(&input.cell)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
            ws.write_array_formula(r1, c1, r2, c2, &input.formula)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
        }
        "dynamic" => {
            let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.cell)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
            ws.write_dynamic_formula(row, col, &input.formula)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
        }
        _ => {
            let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.cell)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
            if let Some(r) = input.cached_result {
                ws.write_formula_with_result(row, col, &input.formula, r)
                    .map_err(|e| anyhow::anyhow!("{e}"))?;
            } else {
                ws.write_formula(row, col, &input.formula)
                    .map_err(|e| anyhow::anyhow!("{e}"))?;
            }
        }
    }
    Ok(success_no_data(&format!(
        "Formula written to {}",
        input.cell
    )))
}

pub fn manage_cell(
    store: &mut WorkbookStore,
    input: ManageCellInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (row, col) =
        zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    match input.action.as_str() {
        "clear" => {
            ws.clear_cell(row, col);
        }
        _ => {
            let fmt = build_format(
                None,
                None,
                None,
                None,
                input.background_color.as_deref(),
                input.number_format.as_deref(),
            );
            ws.write_blank(row, col, &fmt)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
        }
    }
    Ok(success_no_data(&format!(
        "Cell {} {}",
        input.cell, input.action
    )))
}

pub fn manage_comments(
    store: &mut WorkbookStore,
    input: ManageCommentsInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    match input.action.as_str() {
        "read" => {
            let ws = entry
                .data
                .worksheet(idx)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
            let comments: Vec<serde_json::Value> = ws.comments().iter().map(|c|
                serde_json::json!({"cell": zavora_xlsx::utility::to_a1(c.row, c.col), "author": c.author, "text": c.text})
            ).collect();
            Ok(success("Comments read", comments))
        }
        _ => {
            let cell = input.cell.as_deref().unwrap_or("A1");
            let text = input.text.as_deref().unwrap_or("");
            let (row, col) =
                zavora_xlsx::utility::parse_cell_ref(cell).map_err(|e| anyhow::anyhow!("{e}"))?;
            let ws = entry
                .data
                .worksheet(idx)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
            if let Some(ref a) = input.author {
                ws.add_comment_with_author(row, col, text, a);
            } else {
                ws.add_comment(row, col, text);
            }
            Ok(success_no_data(&format!("Comment added at {cell}")))
        }
    }
}

pub fn manage_defined_names(
    store: &mut WorkbookStore,
    input: ManageDefinedNamesInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    match input.action.as_str() {
        "list" => {
            let names: Vec<serde_json::Value> = entry
                .data
                .defined_names()
                .iter()
                .map(|(n, f)| serde_json::json!({"name": n, "formula": f}))
                .collect();
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
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (row, col) =
        zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    match input.link_type.as_str() {
        "internal" => {
            ws.write_internal_link(
                row,
                col,
                &input.target,
                input.display_text.as_deref().unwrap_or(&input.target),
            )
            .map_err(|e| anyhow::anyhow!("{e}"))?;
        }
        _ => {
            ws.write_url(
                row,
                col,
                &input.target,
                input.display_text.as_deref().unwrap_or(&input.target),
            )
            .map_err(|e| anyhow::anyhow!("{e}"))?;
        }
    }
    Ok(success_no_data(&format!("Link added at {}", input.cell)))
}

pub fn protect_consolidated(
    store: &mut WorkbookStore,
    input: ProtectInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    match input.target.as_str() {
        "workbook" => {
            if let Some(ref pw) = input.password {
                entry.data.protect_with_password(pw);
            } else {
                entry.data.protect();
            }
            Ok(success_no_data("Workbook protected"))
        }
        "unprotect_range" => {
            let sn = input.sheet_name.as_deref().unwrap_or("Sheet1");
            let idx = match find_sheet(&entry.data, sn) {
                Some(i) => i,
                None => return Ok(sheet_err(sn)),
            };
            let ws = entry
                .data
                .worksheet(idx)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
            let range = input.range.as_deref().unwrap_or("A1:A1");
            let title = input.range_title.as_deref().unwrap_or("Range");
            if let Some(ref pw) = input.password {
                ws.unprotect_range_with_password(range, title, pw);
            } else {
                ws.unprotect_range(range, title);
            }
            Ok(success_no_data(&format!("Range {range} unprotected")))
        }
        _ => {
            // "sheet"
            let sn = input.sheet_name.as_deref().unwrap_or("Sheet1");
            let idx = match find_sheet(&entry.data, sn) {
                Some(i) => i,
                None => return Ok(sheet_err(sn)),
            };
            let ws = entry
                .data
                .worksheet(idx)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
            if let Some(ref pw) = input.password {
                ws.protect_with_password(pw);
            } else {
                ws.protect();
            }
            Ok(success_no_data(&format!("Sheet '{sn}' protected")))
        }
    }
}

pub fn set_dimensions(
    store: &mut WorkbookStore,
    input: SetDimensionsInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    match input.target.as_str() {
        "row_height" => {
            ws.set_row_height(input.row.unwrap_or(1).saturating_sub(1), input.value)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
        }
        "column_range_width" => {
            let c1 =
                zavora_xlsx::utility::col_from_letter(input.first_column.as_deref().unwrap_or("A"))
                    .map_err(|e| anyhow::anyhow!("{e}"))?;
            let c2 =
                zavora_xlsx::utility::col_from_letter(input.last_column.as_deref().unwrap_or("A"))
                    .map_err(|e| anyhow::anyhow!("{e}"))?;
            ws.set_column_range_width(c1, c2, input.value);
        }
        "default_row_height" => {
            ws.set_default_row_height(input.value);
        }
        _ => {
            // "column_width"
            let col = zavora_xlsx::utility::col_from_letter(input.column.as_deref().unwrap_or("A"))
                .map_err(|e| anyhow::anyhow!("{e}"))?;
            ws.set_column_width(col, input.value)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
        }
    }
    Ok(success_no_data(&format!(
        "{} set to {}",
        input.target, input.value
    )))
}

pub fn set_visibility(
    store: &mut WorkbookStore,
    input: SetVisibilityInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    match input.target.as_str() {
        "row" => {
            let r: u32 = input.identifier.parse().unwrap_or(1);
            ws.set_row_hidden(r.saturating_sub(1), input.hidden);
        }
        _ => {
            let col = zavora_xlsx::utility::col_from_letter(&input.identifier)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
            ws.set_column_hidden(col, input.hidden);
        }
    }
    Ok(success_no_data(&format!(
        "{} {} hidden={}",
        input.target, input.identifier, input.hidden
    )))
}

pub fn set_row_column_format(
    store: &mut WorkbookStore,
    input: SetRowColumnFormatInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let fmt = build_format(
        input.bold,
        input.italic,
        input.font_size,
        input.font_color.as_deref(),
        input.background_color.as_deref(),
        input.number_format.as_deref(),
    );
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    match input.target.as_str() {
        "row" => {
            let r: u32 = input.identifier.parse().unwrap_or(1);
            ws.set_row_format(r.saturating_sub(1), &fmt);
        }
        _ => {
            let col = zavora_xlsx::utility::col_from_letter(&input.identifier)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
            ws.set_column_format(col, &fmt);
        }
    }
    Ok(success_no_data(&format!(
        "{} {} format set",
        input.target, input.identifier
    )))
}

pub fn group_consolidated(
    store: &mut WorkbookStore,
    input: GroupInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    match input.target.as_str() {
        "columns" => {
            let c1 = zavora_xlsx::utility::col_from_letter(&input.start)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
            let c2 = zavora_xlsx::utility::col_from_letter(&input.end)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
            ws.group_columns(c1, c2, input.level);
        }
        _ => {
            let s: u32 = input.start.parse().unwrap_or(1);
            let e: u32 = input.end.parse().unwrap_or(1);
            ws.group_rows(s.saturating_sub(1), e.saturating_sub(1), input.level);
        }
    }
    Ok(success_no_data(&format!(
        "{} {}-{} grouped at level {}",
        input.target, input.start, input.end, input.level
    )))
}

pub fn manage_autofilter(
    store: &mut WorkbookStore,
    input: ManageAutofilterInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (r1, c1, r2, c2) =
        zavora_xlsx::utility::parse_range_ref(&input.range).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    ws.set_autofilter(r1, c1, r2, c2);
    if let (Some(ref fc), Some(ref fv)) = (input.filter_column, input.filter_values) {
        let col = zavora_xlsx::utility::col_from_letter(fc).map_err(|e| anyhow::anyhow!("{e}"))?;
        let strs: Vec<&str> = fv.iter().map(|s| s.as_str()).collect();
        ws.filter_column(col, &strs);
    }
    Ok(success_no_data(&format!(
        "Autofilter set on {}",
        input.range
    )))
}

// ══════════════════════════════════════════════════════════════════
// Batch 9: Feature-parity tools (waterfall, funnel, treemap, shape, doc properties)
// ══════════════════════════════════════════════════════════════════

/// Parse a hex color string like "#FF0000" into [u8; 3].
fn parse_hex_color(hex: &str) -> [u8; 3] {
    let s = hex.trim_start_matches('#');
    if s.len() == 6 {
        let r = u8::from_str_radix(&s[0..2], 16).unwrap_or(0);
        let g = u8::from_str_radix(&s[2..4], 16).unwrap_or(0);
        let b = u8::from_str_radix(&s[4..6], 16).unwrap_or(0);
        [r, g, b]
    } else {
        [0, 0, 0]
    }
}

pub fn add_waterfall_chart(
    store: &mut WorkbookStore,
    input: AddWaterfallChartInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let mut chart = zavora_xlsx::WaterfallChart::new();
    if let Some(ref t) = input.title {
        chart.set_title(t);
    }
    if let Some(ref n) = input.series_name {
        chart.set_series_name(n);
    }
    chart.set_width(input.width);
    chart.set_height(input.height);
    for pt in &input.points {
        let point_type = match pt.point_type {
            crate::types::enums::WaterfallPointKind::Increase => {
                zavora_xlsx::WaterfallPointType::Increase
            }
            crate::types::enums::WaterfallPointKind::Decrease => {
                zavora_xlsx::WaterfallPointType::Decrease
            }
            crate::types::enums::WaterfallPointKind::Total => {
                zavora_xlsx::WaterfallPointType::Total
            }
        };
        chart.add_point(&pt.category, pt.value, point_type);
    }
    let (row, col) = if let Some(ref c) = input.cell {
        zavora_xlsx::utility::parse_cell_ref(c).map_err(|e| anyhow::anyhow!("{e}"))?
    } else {
        (0, 0)
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .insert_waterfall(row, col, &chart)?;
    Ok(success_no_data(&format!(
        "Waterfall chart added to '{}'",
        input.sheet_name
    )))
}

pub fn add_funnel_chart(
    store: &mut WorkbookStore,
    input: AddFunnelChartInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let mut chart = zavora_xlsx::FunnelChart::new();
    if let Some(ref t) = input.title {
        chart.set_title(t);
    }
    if let Some(ref n) = input.series_name {
        chart.set_series_name(n);
    }
    chart.set_width(input.width);
    chart.set_height(input.height);
    for pt in &input.points {
        chart.add_point(&pt.category, pt.value);
    }
    let (row, col) = if let Some(ref c) = input.cell {
        zavora_xlsx::utility::parse_cell_ref(c).map_err(|e| anyhow::anyhow!("{e}"))?
    } else {
        (0, 0)
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .insert_funnel(row, col, &chart)?;
    Ok(success_no_data(&format!(
        "Funnel chart added to '{}'",
        input.sheet_name
    )))
}

pub fn add_treemap_chart(
    store: &mut WorkbookStore,
    input: AddTreemapChartInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let mut chart = zavora_xlsx::TreemapChart::new();
    if let Some(ref t) = input.title {
        chart.set_title(t);
    }
    if let Some(ref n) = input.series_name {
        chart.set_series_name(n);
    }
    chart.set_width(input.width);
    chart.set_height(input.height);
    for pt in &input.points {
        if let Some(ref c) = pt.color {
            chart.add_point_with_color(&pt.category, pt.value, c.as_str());
        } else {
            chart.add_point(&pt.category, pt.value);
        }
    }
    let (row, col) = if let Some(ref c) = input.cell {
        zavora_xlsx::utility::parse_cell_ref(c).map_err(|e| anyhow::anyhow!("{e}"))?
    } else {
        (0, 0)
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .insert_treemap(row, col, &chart)?;
    Ok(success_no_data(&format!(
        "Treemap chart added to '{}'",
        input.sheet_name
    )))
}

pub fn add_shape(store: &mut WorkbookStore, input: AddShapeInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let st = match input.shape_type {
        crate::types::enums::ShapeKind::Rectangle => zavora_xlsx::ShapeType::Rectangle,
        crate::types::enums::ShapeKind::RoundedRectangle => {
            zavora_xlsx::ShapeType::RoundedRectangle
        }
        crate::types::enums::ShapeKind::Ellipse => zavora_xlsx::ShapeType::Ellipse,
        crate::types::enums::ShapeKind::Triangle => zavora_xlsx::ShapeType::Triangle,
        crate::types::enums::ShapeKind::Diamond => zavora_xlsx::ShapeType::Diamond,
        crate::types::enums::ShapeKind::Arrow => zavora_xlsx::ShapeType::Arrow,
        crate::types::enums::ShapeKind::Callout => zavora_xlsx::ShapeType::Callout,
        crate::types::enums::ShapeKind::TextBox => zavora_xlsx::ShapeType::TextBox,
    };
    let mut shape = zavora_xlsx::Shape::new(st, input.width, input.height);
    if let Some(ref t) = input.text {
        shape = shape.text(t);
    }
    if let Some(ref c) = input.fill_color {
        shape = shape.fill_color(parse_hex_color(c));
    }
    if let Some(ref c) = input.outline_color {
        shape = shape.outline_color(parse_hex_color(c));
    }
    if let Some(w) = input.outline_width {
        shape = shape.outline_width(w);
    }
    if let Some(s) = input.font_size {
        shape = shape.font_size(s);
    }
    if input.bold == Some(true) {
        shape = shape.bold();
    }
    let (row, col) =
        zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .add_shape(row, col, &shape)?;
    Ok(success_no_data(&format!("Shape added at {}", input.cell)))
}

pub fn set_doc_properties(
    store: &mut WorkbookStore,
    input: SetDocPropertiesInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let mut props = zavora_xlsx::DocProperties::new();
    if let Some(ref v) = input.title {
        props = props.title(v);
    }
    if let Some(ref v) = input.author {
        props = props.author(v);
    }
    if let Some(ref v) = input.subject {
        props = props.subject(v);
    }
    if let Some(ref v) = input.description {
        props = props.description(v);
    }
    if let Some(ref v) = input.keywords {
        props = props.keywords(v);
    }
    if let Some(ref v) = input.category {
        props = props.category(v);
    }
    if let Some(ref v) = input.company {
        props = props.company(v);
    }
    entry.data.set_properties(props);
    Ok(success_no_data("Document properties set"))
}

// ══════════════════════════════════════════════════════════════════
// v0.2.0: New tools — 4 chart types, slicers, timelines, form controls,
// advanced save, named ranges CRUD, sheet metadata, chart sheet,
// chart enhancements, protection options
// ══════════════════════════════════════════════════════════════════

pub fn add_sunburst_chart(
    store: &mut WorkbookStore,
    input: AddSunburstChartInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let mut chart = zavora_xlsx::SunburstChart::new();
    if let Some(ref t) = input.title {
        chart.set_title(t);
    }
    if let Some(ref n) = input.series_name {
        chart.set_series_name(n);
    }
    chart.set_width(input.width);
    chart.set_height(input.height);
    // SunburstChart uses add_level for hierarchy labels and set_values for leaf sizes
    let labels: Vec<&str> = input.points.iter().map(|p| p.category.as_str()).collect();
    let values: Vec<f64> = input.points.iter().map(|p| p.value).collect();
    chart.add_level(&labels);
    chart.set_values(&values);
    let (row, col) = if let Some(ref c) = input.cell {
        zavora_xlsx::utility::parse_cell_ref(c).map_err(|e| anyhow::anyhow!("{e}"))?
    } else {
        (0, 0)
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .insert_sunburst(row, col, &chart)?;
    Ok(success_no_data(&format!(
        "Sunburst chart added to '{}'",
        input.sheet_name
    )))
}

pub fn add_histogram_chart(
    store: &mut WorkbookStore,
    input: AddHistogramChartInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let mut chart = if input.pareto == Some(true) {
        zavora_xlsx::HistogramChart::pareto()
    } else {
        zavora_xlsx::HistogramChart::new()
    };
    if let Some(ref t) = input.title {
        chart.set_title(t);
    }
    if let Some(ref n) = input.series_name {
        chart.set_series_name(n);
    }
    chart.set_width(input.width);
    chart.set_height(input.height);
    if let Some(bc) = input.bin_count {
        chart.set_bin_count(bc);
    }
    if let Some(bw) = input.bin_width {
        chart.set_bin_width(bw);
    }
    let values: Vec<f64> = input.points.iter().map(|p| p.value).collect();
    chart.set_values(&values);
    let (row, col) = if let Some(ref c) = input.cell {
        zavora_xlsx::utility::parse_cell_ref(c).map_err(|e| anyhow::anyhow!("{e}"))?
    } else {
        (0, 0)
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .insert_histogram(row, col, &chart)?;
    Ok(success_no_data(&format!(
        "Histogram chart added to '{}'",
        input.sheet_name
    )))
}

pub fn add_box_whisker_chart(
    store: &mut WorkbookStore,
    input: AddBoxWhiskerChartInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let mut chart = zavora_xlsx::BoxWhiskerChart::new();
    if let Some(ref t) = input.title {
        chart.set_title(t);
    }
    if let Some(ref n) = input.series_name {
        chart.set_series_name(n);
    }
    chart.set_width(input.width);
    chart.set_height(input.height);
    if let Some(v) = input.show_outliers {
        chart.set_show_outliers(v);
    }
    if let Some(v) = input.show_mean {
        chart.set_show_mean_markers(v);
    }
    if let Some(v) = input.show_inner_points {
        chart.set_show_inner_points(v);
    }
    // BoxWhiskerChart uses add_data_set(category, &[f64]) for each box
    for pt in &input.points {
        chart.add_data_set(&pt.category, &[pt.value]);
    }
    let (row, col) = if let Some(ref c) = input.cell {
        zavora_xlsx::utility::parse_cell_ref(c).map_err(|e| anyhow::anyhow!("{e}"))?
    } else {
        (0, 0)
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .insert_box_whisker(row, col, &chart)?;
    Ok(success_no_data(&format!(
        "Box & whisker chart added to '{}'",
        input.sheet_name
    )))
}

pub fn add_map_chart(
    store: &mut WorkbookStore,
    input: AddMapChartInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let mut chart = zavora_xlsx::MapChart::new();
    if let Some(ref t) = input.title {
        chart.set_title(t);
    }
    if let Some(ref n) = input.series_name {
        chart.set_series_name(n);
    }
    chart.set_width(input.width);
    chart.set_height(input.height);
    if let Some(ref ml) = input.map_level {
        let level = match ml.as_str() {
            "region" => zavora_xlsx::MapLevel::Region,
            _ => zavora_xlsx::MapLevel::Country,
        };
        chart.set_map_level(level);
    }
    for pt in &input.points {
        chart.add_point(&pt.category, pt.value);
    }
    let (row, col) = if let Some(ref c) = input.cell {
        zavora_xlsx::utility::parse_cell_ref(c).map_err(|e| anyhow::anyhow!("{e}"))?
    } else {
        (0, 0)
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .insert_map(row, col, &chart)?;
    Ok(success_no_data(&format!(
        "Map chart added to '{}'",
        input.sheet_name
    )))
}

pub fn add_slicer(
    store: &mut WorkbookStore,
    input: AddSlicerInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    // Slicer uses builder pattern (takes ownership)
    let mut slicer = zavora_xlsx::Slicer::new(&input.pivot_table_name, &input.field_name);
    if let Some(w) = input.width {
        slicer = slicer.set_width(w);
    }
    if let Some(h) = input.height {
        slicer = slicer.set_height(h);
    }
    let (row, col) = if let Some(ref c) = input.cell {
        zavora_xlsx::utility::parse_cell_ref(c).map_err(|e| anyhow::anyhow!("{e}"))?
    } else {
        (0, 0)
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .add_slicer(row, col, &slicer)?;
    Ok(success_no_data(&format!(
        "Slicer for '{}' added to '{}'",
        input.field_name, input.sheet_name
    )))
}

pub fn add_timeline(
    store: &mut WorkbookStore,
    input: AddTimelineInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let timeline = zavora_xlsx::Timeline::new(&input.pivot_table_name, &input.field_name);
    let (row, col) = if let Some(ref c) = input.cell {
        zavora_xlsx::utility::parse_cell_ref(c).map_err(|e| anyhow::anyhow!("{e}"))?
    } else {
        (0, 0)
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .add_timeline(row, col, &timeline)?;
    Ok(success_no_data(&format!(
        "Timeline for '{}' added to '{}'",
        input.field_name, input.sheet_name
    )))
}

pub fn add_form_control(
    store: &mut WorkbookStore,
    input: AddFormControlInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (row, col) =
        zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let text = input.text.as_deref().unwrap_or("Control");
    let fc = match input.control_type.as_str() {
        "checkbox" => {
            if let Some(ref cl) = input.cell_link {
                zavora_xlsx::FormControl::checkbox_with_link(text, cl)
            } else {
                zavora_xlsx::FormControl::checkbox(text)
            }
        }
        "dropdown" => {
            let items = input
                .input_range
                .as_deref()
                .unwrap_or("")
                .split(',')
                .map(|s| s.trim().to_string())
                .collect();
            zavora_xlsx::FormControl::dropdown(items)
        }
        "spinner" => zavora_xlsx::FormControl::spinner(0, 100, 0),
        _ => zavora_xlsx::FormControl::button(text),
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .add_form_control(row, col, fc);
    Ok(success_no_data(&format!(
        "Form control '{}' added at {}",
        input.control_type, input.cell
    )))
}

pub fn save_workbook_advanced(
    store: &mut WorkbookStore,
    input: SaveWorkbookAdvancedInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    if entry.read_only {
        return Ok(error(
            ErrorCategory::EngineUnsupported,
            "Read-only workbooks cannot be saved",
            "Reopen in edit mode.",
        ));
    }
    let _ = entry.data.recalculate();
    let path = std::path::Path::new(&input.file_path);
    match input.format.as_str() {
        "template" => entry
            .data
            .save_as_template(path)
            .map_err(|e| anyhow::anyhow!("{e}"))?,
        "encrypted" => {
            let pw = input.password.as_deref().unwrap_or("");
            entry
                .data
                .save_encrypted(path, pw)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
        }
        "parallel" => entry
            .data
            .save_parallel(path)
            .map_err(|e| anyhow::anyhow!("{e}"))?,
        _ => entry
            .data
            .save(path)
            .map_err(|e| anyhow::anyhow!("{e}"))?,
    }
    Ok(success_no_data(&format!(
        "Workbook saved as {} to {}",
        input.format, input.file_path
    )))
}

pub fn open_workbook_encrypted(
    store: &mut WorkbookStore,
    input: OpenWorkbookEncryptedInput,
) -> Result<String, anyhow::Error> {
    use crate::store::WorkbookEntry;
    use std::time::Instant;

    if store.is_full() {
        return Ok(error(
            ErrorCategory::CapacityExceeded,
            "Workbook store is at maximum capacity",
            "Save and close an existing workbook first.",
        ));
    }
    let path = std::path::Path::new(&input.file_path);
    if !path.exists() {
        return Ok(error(
            ErrorCategory::NotFound,
            &format!("File not found: {}", input.file_path),
            "Check the file path.",
        ));
    }
    let wb = zavora_xlsx::Workbook::open_with_password(path, &input.password)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    let sheets = crate::engines::zavora::sheet_summaries(&wb);
    let entry = WorkbookEntry {
        id: String::new(),
        data: wb,
        read_only: false,
        last_access: Instant::now(),
    };
    let id = store.insert(entry).map_err(|e| anyhow::anyhow!("{}", e))?;
    Ok(success(
        "Encrypted workbook opened",
        crate::types::responses::WorkbookInfo {
            workbook_id: id,
            engine: "zavora-xlsx".to_string(),
            sheets,
        },
    ))
}

pub fn manage_named_ranges(
    store: &mut WorkbookStore,
    input: ManageNamedRangesInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    match input.action.as_str() {
        "add" => {
            let name = input.name.as_deref().unwrap_or("");
            let formula = input.formula.as_deref().unwrap_or("");
            entry.data.define_name(name, formula);
            Ok(success_no_data(&format!("Named range '{}' added", name)))
        }
        "add_scoped" => {
            let name = input.name.as_deref().unwrap_or("");
            let formula = input.formula.as_deref().unwrap_or("");
            let sheet_idx = input.sheet_index.unwrap_or(0);
            entry.data.define_name_scoped(name, formula, sheet_idx);
            Ok(success_no_data(&format!(
                "Scoped named range '{}' added for sheet {}",
                name, sheet_idx
            )))
        }
        "update" => {
            let name = input.name.as_deref().unwrap_or("");
            let formula = input.formula.as_deref().unwrap_or("");
            let _ = entry.data.update_named_range(name, formula);
            Ok(success_no_data(&format!("Named range '{}' updated", name)))
        }
        "remove" => {
            let name = input.name.as_deref().unwrap_or("");
            let scope = if let Some(idx) = input.sheet_index {
                zavora_xlsx::DefinedNameScope::Sheet(idx)
            } else {
                zavora_xlsx::DefinedNameScope::Workbook
            };
            let _ = entry.data.remove_named_range(name, &scope);
            Ok(success_no_data(&format!("Named range '{}' removed", name)))
        }
        _ => {
            // "list"
            let names: Vec<serde_json::Value> = entry
                .data
                .defined_names_with_scope()
                .iter()
                .map(|dn| {
                    serde_json::json!({
                        "name": dn.name,
                        "formula": dn.formula,
                        "scope": format!("{:?}", dn.scope),
                    })
                })
                .collect();
            Ok(success("Named ranges listed", names))
        }
    }
}

pub fn read_sheet_metadata(
    store: &mut WorkbookStore,
    input: ReadSheetMetadataInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    let mut result = serde_json::Map::new();

    if input.info == "used_range" || input.info == "all" {
        if let Some((r1, c1, r2, c2)) = ws.used_range() {
            let range_str = format!(
                "{}:{}",
                zavora_xlsx::utility::to_a1(r1, c1),
                zavora_xlsx::utility::to_a1(r2, c2)
            );
            result.insert("used_range".into(), serde_json::json!(range_str));
        } else {
            result.insert("used_range".into(), serde_json::Value::Null);
        }
    }

    if input.info == "hyperlinks" || input.info == "all" {
        let links: Vec<serde_json::Value> = ws
            .hyperlinks()
            .iter()
            .map(|h| {
                serde_json::json!({
                    "cell": zavora_xlsx::utility::to_a1(h.row, h.col),
                    "url": h.url,
                    "location": h.location,
                    "tooltip": h.tooltip,
                })
            })
            .collect();
        result.insert("hyperlinks".into(), serde_json::json!(links));
    }

    if input.info == "merge_ranges" || input.info == "all" {
        let merges: Vec<String> = ws
            .merge_ranges()
            .iter()
            .map(|(r1, c1, r2, c2)| {
                format!(
                    "{}:{}",
                    zavora_xlsx::utility::to_a1(*r1, *c1),
                    zavora_xlsx::utility::to_a1(*r2, *c2)
                )
            })
            .collect();
        result.insert("merge_ranges".into(), serde_json::json!(merges));
    }

    if input.info == "charts" || input.info == "all" {
        let charts: Vec<serde_json::Value> = ws
            .charts()
            .iter()
            .map(|c| {
                serde_json::json!({
                    "title": c.title(),
                    "type": format!("{:?}", c.chart_type()),
                    "series_count": c.series().len(),
                })
            })
            .collect();
        result.insert("charts".into(), serde_json::json!(charts));
    }

    Ok(success(
        "Sheet metadata read",
        serde_json::Value::Object(result),
    ))
}

pub fn add_chart_sheet(
    store: &mut WorkbookStore,
    input: AddChartSheetInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let ct = match input.chart_type {
        ChartType::Bar => zavora_xlsx::ChartType::Bar,
        ChartType::Column => zavora_xlsx::ChartType::Column,
        ChartType::Line => zavora_xlsx::ChartType::Line,
        ChartType::Pie => zavora_xlsx::ChartType::Pie,
        ChartType::Scatter => zavora_xlsx::ChartType::Scatter,
        ChartType::Area => zavora_xlsx::ChartType::Area,
        ChartType::Doughnut => zavora_xlsx::ChartType::Doughnut,
    };
    let mut chart = zavora_xlsx::Chart::new(ct);
    if let Some(ref t) = input.title {
        chart.set_title(t);
    }
    if let Some(ref x) = input.x_axis_label {
        chart.set_x_axis_name(x);
    }
    if let Some(ref y) = input.y_axis_label {
        chart.set_y_axis_name(y);
    }
    if let Some(ref lp) = input.legend_position {
        chart.set_legend_position(match lp {
            LegendPosition::Top => zavora_xlsx::LegendPosition::Top,
            LegendPosition::Bottom => zavora_xlsx::LegendPosition::Bottom,
            LegendPosition::Left => zavora_xlsx::LegendPosition::Left,
            LegendPosition::Right => zavora_xlsx::LegendPosition::Right,
            LegendPosition::None => zavora_xlsx::LegendPosition::None,
        });
    }
    if !input.series.is_empty() {
        for si in &input.series {
            let s = chart.add_series();
            s.set_values(&si.values);
            if let Some(ref c) = si.categories {
                s.set_categories(c);
            }
            if let Some(ref n) = si.name {
                s.set_name(n);
            }
        }
    } else if let Some(ref dr) = input.data_range {
        chart.add_series().set_values(dr);
    }
    entry
        .data
        .add_chart_sheet(&input.sheet_name, chart)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!(
        "Chart sheet '{}' added",
        input.sheet_name
    )))
}


// ══════════════════════════════════════════════════════════════════
// v0.2.1: Threaded comments, granular protection, custom properties
// ══════════════════════════════════════════════════════════════════

pub fn add_threaded_comment(
    store: &mut WorkbookStore,
    input: AddThreadedCommentInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (row, col) =
        zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let mut tc = zavora_xlsx::ThreadedComment::new(&input.author, &input.text);
    if let Some(ref ts) = input.timestamp {
        tc = tc.timestamp(ts);
    }
    for reply in &input.replies {
        if let Some(ref ts) = reply.timestamp {
            tc.add_reply_with_timestamp(&reply.author, &reply.text, ts);
        } else {
            tc.add_reply(&reply.author, &reply.text);
        }
    }
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .add_threaded_comment(row, col, tc);
    Ok(success_no_data(&format!(
        "Threaded comment added at {}",
        input.cell
    )))
}

pub fn protect_sheet_advanced(
    store: &mut WorkbookStore,
    input: ProtectSheetAdvancedInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    if let Some(ref pw) = input.password {
        ws.protect_with_password(pw);
    }

    // Build SheetProtection with granular options
    // Default: everything locked (true = locked)
    let mut prot = zavora_xlsx::SheetProtection::default();

    // "allow" means NOT locked, so we invert: allow=true -> field=false
    if let Some(v) = input.allow_insert_rows {
        prot.insert_rows = !v;
    }
    if let Some(v) = input.allow_delete_rows {
        prot.delete_rows = !v;
    }
    if let Some(v) = input.allow_insert_columns {
        prot.insert_columns = !v;
    }
    if let Some(v) = input.allow_delete_columns {
        prot.delete_columns = !v;
    }
    if let Some(v) = input.allow_format_cells {
        prot.format_cells = !v;
    }
    if let Some(v) = input.allow_format_columns {
        prot.format_columns = !v;
    }
    if let Some(v) = input.allow_format_rows {
        prot.format_rows = !v;
    }
    if let Some(v) = input.allow_sort {
        prot.sort = !v;
    }
    if let Some(v) = input.allow_insert_hyperlinks {
        prot.insert_hyperlinks = !v;
    }
    if let Some(v) = input.allow_select_locked_cells {
        prot.select_locked_cells = !v;
    }
    if let Some(v) = input.allow_select_unlocked_cells {
        prot.select_unlocked_cells = !v;
    }
    if let Some(v) = input.allow_pivot_tables {
        prot.pivot_tables = !v;
    }

    ws.protect_with_options(prot);

    Ok(success_no_data(&format!(
        "Sheet '{}' protected with custom options",
        input.sheet_name
    )))
}

pub fn set_custom_property(
    store: &mut WorkbookStore,
    input: SetCustomPropertyInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let val = match input.value_type.as_str() {
        "number" => {
            let n: f64 = input.value.parse().unwrap_or(0.0);
            zavora_xlsx::CustomPropertyValue::Number(n)
        }
        "integer" => {
            let n: i32 = input.value.parse().unwrap_or(0);
            zavora_xlsx::CustomPropertyValue::Integer(n)
        }
        "bool" => {
            let b = input.value == "true" || input.value == "1";
            zavora_xlsx::CustomPropertyValue::Bool(b)
        }
        "datetime" => zavora_xlsx::CustomPropertyValue::DateTime(input.value.clone()),
        _ => zavora_xlsx::CustomPropertyValue::Text(input.value.clone()),
    };
    entry.data.set_custom_property(&input.name, val);
    Ok(success_no_data(&format!(
        "Custom property '{}' set",
        input.name
    )))
}

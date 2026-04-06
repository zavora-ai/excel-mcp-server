use super::common::workbook_not_found;
use crate::store::WorkbookStore;
use crate::types::inputs::*;
use crate::types::responses::*;

pub fn add_chart(store: &mut WorkbookStore, input: AddChartInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let ct = match input.chart_type {
        crate::types::enums::ChartType::Bar => zavora_xlsx::ChartType::Bar,
        crate::types::enums::ChartType::Column => zavora_xlsx::ChartType::Column,
        crate::types::enums::ChartType::Line => zavora_xlsx::ChartType::Line,
        crate::types::enums::ChartType::Pie => zavora_xlsx::ChartType::Pie,
        crate::types::enums::ChartType::Scatter => zavora_xlsx::ChartType::Scatter,
        crate::types::enums::ChartType::Area => zavora_xlsx::ChartType::Area,
        crate::types::enums::ChartType::Doughnut => zavora_xlsx::ChartType::Doughnut,
    };
    let mut chart = zavora_xlsx::Chart::new(ct);
    if let Some(ref t) = input.title { chart.set_title(t); }
    if let Some(ref x) = input.x_axis_label { chart.set_x_axis_name(x); }
    if let Some(ref y) = input.y_axis_label { chart.set_y_axis_name(y); }
    if let Some(ref lp) = input.legend_position {
        chart.set_legend_position(match lp {
            crate::types::enums::LegendPosition::Top => zavora_xlsx::LegendPosition::Top,
            crate::types::enums::LegendPosition::Bottom => zavora_xlsx::LegendPosition::Bottom,
            crate::types::enums::LegendPosition::Left => zavora_xlsx::LegendPosition::Left,
            crate::types::enums::LegendPosition::Right => zavora_xlsx::LegendPosition::Right,
            crate::types::enums::LegendPosition::None => zavora_xlsx::LegendPosition::None,
        });
    }
    chart.set_width(input.width);
    chart.set_height(input.height);
    // Parse data_range and add as series
    let s = chart.add_series();
    s.set_values(&input.data_range);
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    ws.insert_chart(0, 0, &chart)?;
    Ok(success_no_data(&format!("Chart added to '{}'", input.sheet_name)))
}

fn find_sheet(wb: &zavora_xlsx::Workbook, name: &str) -> Option<usize> { wb.sheet_names().iter().position(|n| *n == name) }
fn sheet_err(name: &str) -> String { error(ErrorCategory::NotFound, &format!("Sheet '{}' not found", name), "Check sheet name.") }

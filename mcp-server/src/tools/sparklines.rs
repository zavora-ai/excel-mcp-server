use super::common::workbook_not_found;
use crate::store::WorkbookStore;
use crate::types::inputs::*;
use crate::types::responses::*;

pub fn add_sparkline(store: &mut WorkbookStore, input: AddSparklineInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.target_cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let st = match input.sparkline_type {
        crate::types::enums::SparklineType::Line => zavora_xlsx::SparklineType::Line,
        crate::types::enums::SparklineType::Column => zavora_xlsx::SparklineType::Column,
        crate::types::enums::SparklineType::WinLoss => zavora_xlsx::SparklineType::WinLoss,
    };
    let sparkline = zavora_xlsx::Sparkline::new(&input.data_range, st);
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.add_sparkline(row, col, &sparkline)?;
    Ok(success_no_data(&format!("Sparkline added at {}", input.target_cell)))
}

fn find_sheet(wb: &zavora_xlsx::Workbook, name: &str) -> Option<usize> { wb.sheet_names().iter().position(|n| *n == name) }
fn sheet_err(name: &str) -> String { error(ErrorCategory::NotFound, &format!("Sheet '{}' not found", name), "Check sheet name.") }

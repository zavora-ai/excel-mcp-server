use super::common::workbook_not_found;
use crate::store::WorkbookStore;
use crate::types::inputs::*;
use crate::types::responses::*;

pub fn set_column_width(store: &mut WorkbookStore, input: SetColumnWidthInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let col = zavora_xlsx::utility::col_from_letter(&input.column).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.set_column_width(col, input.width).map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!("Column {} width set to {}", input.column, input.width)))
}

pub fn set_row_height(store: &mut WorkbookStore, input: SetRowHeightInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let row = input.row.saturating_sub(1); // input is 1-based
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.set_row_height(row, input.height).map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!("Row {} height set to {}", input.row, input.height)))
}

pub fn freeze_panes(store: &mut WorkbookStore, input: FreezePanesInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.set_freeze_panes(row, col).map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!("Panes frozen at {}", input.cell)))
}

fn find_sheet(wb: &zavora_xlsx::Workbook, name: &str) -> Option<usize> { wb.sheet_names().iter().position(|n| *n == name) }
fn sheet_err(name: &str) -> String { error(ErrorCategory::NotFound, &format!("Sheet '{}' not found", name), "Check sheet name.") }

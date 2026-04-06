use super::common::workbook_not_found;
use crate::engines::zavora;
use crate::store::WorkbookStore;
use crate::types::inputs::*;
use crate::types::responses::*;

pub fn write_cells(store: &mut WorkbookStore, input: WriteCellsInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i, None => return Ok(sheet_not_found(&input.sheet_name)),
    };
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    let mut count = 0usize;
    for cw in &input.cells {
        let (row, col) = match zavora_xlsx::utility::parse_cell_ref(&cw.cell) {
            Ok(rc) => rc,
            Err(e) => return Ok(error(ErrorCategory::ParseError, &format!("Invalid cell reference '{}': {e}", cw.cell), "Use A1 notation.")),
        };
        if let Err(e) = zavora::write_json_value(ws, row, col, &cw.value) {
            return Ok(error(ErrorCategory::IoError, &format!("Write error: {e}"), "Check value type."));
        }
        count += 1;
    }
    Ok(success("Cells written", WriteResult { cells_written: count, range_covered: format!("{} cells", count) }))
}

pub fn write_row(store: &mut WorkbookStore, input: WriteRowInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i, None => return Ok(sheet_not_found(&input.sheet_name)),
    };
    let (row, start_col) = zavora_xlsx::utility::parse_cell_ref(&input.start_cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    for (i, val) in input.values.iter().enumerate() {
        zavora::write_json_value(ws, row, start_col + i as u16, val).map_err(|e| anyhow::anyhow!("{e}"))?;
    }
    let end = zavora_xlsx::utility::to_a1(row, start_col + input.values.len() as u16 - 1);
    Ok(success("Row written", WriteResult { cells_written: input.values.len(), range_covered: format!("{}:{}", input.start_cell, end) }))
}

pub fn write_column(store: &mut WorkbookStore, input: WriteColumnInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i, None => return Ok(sheet_not_found(&input.sheet_name)),
    };
    let (start_row, col) = zavora_xlsx::utility::parse_cell_ref(&input.start_cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    for (i, val) in input.values.iter().enumerate() {
        zavora::write_json_value(ws, start_row + i as u32, col, val).map_err(|e| anyhow::anyhow!("{e}"))?;
    }
    let end = zavora_xlsx::utility::to_a1(start_row + input.values.len() as u32 - 1, col);
    Ok(success("Column written", WriteResult { cells_written: input.values.len(), range_covered: format!("{}:{}", input.start_cell, end) }))
}

fn find_sheet(wb: &zavora_xlsx::Workbook, name: &str) -> Option<usize> {
    wb.sheet_names().iter().position(|n| *n == name)
}

fn sheet_not_found(name: &str) -> String {
    error(ErrorCategory::NotFound, &format!("Sheet '{}' not found", name), "Check sheet name.")
}

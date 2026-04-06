use super::common::workbook_not_found;
use crate::store::WorkbookStore;
use crate::types::inputs::*;
use crate::types::responses::*;

pub fn add_image(store: &mut WorkbookStore, input: AddImageInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let mut img = zavora_xlsx::Image::from_path(&input.image_path).map_err(|e| anyhow::anyhow!("{e}"))?;
    if let Some(w) = input.width { img.set_scale_width(w as f64 / 200.0); }
    if let Some(h) = input.height { img.set_scale_height(h as f64 / 200.0); }
    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.insert_image(row, col, &img)?;
    Ok(success_no_data(&format!("Image added at {}", input.cell)))
}

fn find_sheet(wb: &zavora_xlsx::Workbook, name: &str) -> Option<usize> { wb.sheet_names().iter().position(|n| *n == name) }
fn sheet_err(name: &str) -> String { error(ErrorCategory::NotFound, &format!("Sheet '{}' not found", name), "Check sheet name.") }

use super::common::workbook_not_found;
use crate::store::WorkbookStore;
use crate::types::inputs::*;
use crate::types::responses::*;

pub fn set_cell_format(store: &mut WorkbookStore, input: SetCellFormatInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (r1, c1, r2, c2) = zavora_xlsx::utility::parse_range_ref(&input.range).map_err(|e| anyhow::anyhow!("{e}"))?;
    let mut fmt = zavora_xlsx::Format::new();
    if input.bold == Some(true) { fmt = fmt.bold(); }
    if input.italic == Some(true) { fmt = fmt.italic(); }
    if input.underline == Some(true) { fmt = fmt.underline(zavora_xlsx::Underline::Single); }
    if let Some(size) = input.font_size { fmt = fmt.font_size(size); }
    if let Some(ref c) = input.font_color { fmt = fmt.font_color(c.as_str()); }
    if let Some(ref c) = input.background_color { fmt = fmt.background_color(c.as_str()); }
    if let Some(ref nf) = input.number_format { fmt = fmt.num_format(nf); }
    if let Some(ref ha) = input.horizontal_alignment {
        fmt = fmt.align(match ha {
            crate::types::enums::HorizontalAlignment::Left => zavora_xlsx::Align::Left,
            crate::types::enums::HorizontalAlignment::Center => zavora_xlsx::Align::Center,
            crate::types::enums::HorizontalAlignment::Right => zavora_xlsx::Align::Right,
            crate::types::enums::HorizontalAlignment::Fill => zavora_xlsx::Align::Left,
            crate::types::enums::HorizontalAlignment::Justify => zavora_xlsx::Align::Left,
        });
    }
    if let Some(ref va) = input.vertical_alignment {
        fmt = fmt.align(match va {
            crate::types::enums::VerticalAlignment::Top => zavora_xlsx::Align::Top,
            crate::types::enums::VerticalAlignment::Center => zavora_xlsx::Align::VerticalCenter,
            crate::types::enums::VerticalAlignment::Bottom => zavora_xlsx::Align::Bottom,
            crate::types::enums::VerticalAlignment::Justify => zavora_xlsx::Align::Bottom,
        });
    }
    if let Some(ref bs) = input.border_style {
        let style = match bs {
            crate::types::enums::BorderStyle::Thin => zavora_xlsx::BorderStyle::Thin,
            crate::types::enums::BorderStyle::Medium => zavora_xlsx::BorderStyle::Medium,
            crate::types::enums::BorderStyle::Thick => zavora_xlsx::BorderStyle::Thick,
            crate::types::enums::BorderStyle::Dashed => zavora_xlsx::BorderStyle::Dashed,
            crate::types::enums::BorderStyle::Dotted => zavora_xlsx::BorderStyle::Dotted,
            crate::types::enums::BorderStyle::Double => zavora_xlsx::BorderStyle::Double,
            crate::types::enums::BorderStyle::None => zavora_xlsx::BorderStyle::None,
        };
        fmt = fmt.border(style);
    }
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    ws.set_range_format(r1, c1, r2, c2, &fmt).map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!("Format applied to {}", input.range)))
}

pub fn merge_cells(store: &mut WorkbookStore, input: MergeCellsInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (r1, c1, r2, c2) = zavora_xlsx::utility::parse_range_ref(&input.range).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    ws.merge_range(r1, c1, r2, c2, "", &zavora_xlsx::Format::new()).map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!("Cells merged: {}", input.range)))
}

fn find_sheet(wb: &zavora_xlsx::Workbook, name: &str) -> Option<usize> { wb.sheet_names().iter().position(|n| *n == name) }
fn sheet_err(name: &str) -> String { error(ErrorCategory::NotFound, &format!("Sheet '{}' not found", name), "Check sheet name.") }

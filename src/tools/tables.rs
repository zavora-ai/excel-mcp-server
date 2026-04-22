use super::common::workbook_not_found;
use crate::store::WorkbookStore;
use crate::types::inputs::*;
use crate::types::responses::*;

pub fn add_table(store: &mut WorkbookStore, input: AddTableInput) -> Result<String, anyhow::Error> {
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
    let cols: Vec<zavora_xlsx::TableColumn> = input
        .columns
        .iter()
        .map(|n| zavora_xlsx::TableColumn::new(n))
        .collect();
    let mut table = zavora_xlsx::Table::new();
    table.set_columns(&cols);
    if let Some(ref style) = input.style {
        let s = style.to_lowercase();
        if s.contains("medium") {
            let n: u8 = s
                .chars()
                .filter(|c| c.is_ascii_digit())
                .collect::<String>()
                .parse()
                .unwrap_or(2);
            table.set_style(zavora_xlsx::TableStyle::Medium(n));
        } else if s.contains("light") {
            let n: u8 = s
                .chars()
                .filter(|c| c.is_ascii_digit())
                .collect::<String>()
                .parse()
                .unwrap_or(1);
            table.set_style(zavora_xlsx::TableStyle::Light(n));
        } else if s.contains("dark") {
            let n: u8 = s
                .chars()
                .filter(|c| c.is_ascii_digit())
                .collect::<String>()
                .parse()
                .unwrap_or(1);
            table.set_style(zavora_xlsx::TableStyle::Dark(n));
        }
    }
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .add_table(r1, c1, r2, c2, &table)?;
    Ok(success_no_data(&format!(
        "Table added to '{}'",
        input.sheet_name
    )))
}

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

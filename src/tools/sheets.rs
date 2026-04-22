use super::common::workbook_not_found;
use crate::engines::zavora;
use crate::store::WorkbookStore;
use crate::types::inputs::*;
use crate::types::responses::*;

pub fn list_sheets(
    store: &mut WorkbookStore,
    input: ListSheetsInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let sheets = zavora::sheet_summaries(&entry.data);
    Ok(success("Sheets listed", sheets))
}

pub fn get_sheet_dimensions(
    store: &mut WorkbookStore,
    input: GetSheetDimensionsInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet_index(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => {
            return Ok(error(
                ErrorCategory::NotFound,
                &format!("Sheet '{}' not found", input.sheet_name),
                "Check sheet name.",
            ))
        }
    };
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    let summary = match ws.used_range() {
        Some((r1, c1, r2, c2)) => SheetSummary {
            name: input.sheet_name,
            dimensions: Some(format!(
                "{}:{}",
                zavora_xlsx::utility::to_a1(r1, c1),
                zavora_xlsx::utility::to_a1(r2, c2)
            )),
            row_count: Some(r2 - r1 + 1),
            col_count: Some(c2 - c1 + 1),
        },
        None => SheetSummary {
            name: input.sheet_name,
            dimensions: None,
            row_count: None,
            col_count: None,
        },
    };
    Ok(success("Sheet dimensions retrieved", summary))
}

pub fn describe_workbook(
    store: &mut WorkbookStore,
    input: DescribeWorkbookInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let names: Vec<String> = entry
        .data
        .sheet_names()
        .iter()
        .map(|s| s.to_string())
        .collect();
    let mut sheets = Vec::new();
    for (i, name) in names.iter().enumerate() {
        let ws = entry
            .data
            .worksheet_ref(i)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
        let (dims, rc, cc, samples) = match ws.used_range() {
            Some((r1, c1, r2, c2)) => {
                let mut rows = Vec::new();
                for r in r1..=(r1 + 4).min(r2) {
                    let mut row = Vec::new();
                    for c in c1..=c2 {
                        row.push(zavora::cell_to_json(&ws.read_cell(r, c)));
                    }
                    rows.push(row);
                }
                (
                    Some(format!(
                        "{}:{}",
                        zavora_xlsx::utility::to_a1(r1, c1),
                        zavora_xlsx::utility::to_a1(r2, c2)
                    )),
                    Some(r2 - r1 + 1),
                    Some(c2 - c1 + 1),
                    rows,
                )
            }
            None => (None, None, None, Vec::new()),
        };
        sheets.push(SheetDescription {
            name: name.clone(),
            dimensions: dims,
            row_count: rc,
            col_count: cc,
            sample_rows: samples,
        });
    }
    Ok(success(
        "Workbook described",
        WorkbookDescription {
            workbook_id: input.workbook_id,
            engine: "zavora-xlsx".to_string(),
            sheets,
        },
    ))
}

pub fn add_sheet(store: &mut WorkbookStore, input: AddSheetInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    entry
        .data
        .add_worksheet_with_name(&input.sheet_name)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!(
        "Sheet '{}' added",
        input.sheet_name
    )))
}

pub fn rename_sheet(
    store: &mut WorkbookStore,
    input: RenameSheetInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet_index(&entry.data, &input.current_name) {
        Some(i) => i,
        None => {
            return Ok(error(
                ErrorCategory::NotFound,
                &format!("Sheet '{}' not found", input.current_name),
                "Check sheet name.",
            ))
        }
    };
    entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?
        .set_name(&input.new_name)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!(
        "Sheet renamed to '{}'",
        input.new_name
    )))
}

pub fn delete_sheet(
    store: &mut WorkbookStore,
    input: DeleteSheetInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet_index(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => {
            return Ok(error(
                ErrorCategory::NotFound,
                &format!("Sheet '{}' not found", input.sheet_name),
                "Check sheet name.",
            ))
        }
    };
    entry
        .data
        .remove_worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    Ok(success_no_data(&format!(
        "Sheet '{}' deleted",
        input.sheet_name
    )))
}

fn find_sheet_index(wb: &zavora_xlsx::Workbook, name: &str) -> Option<usize> {
    wb.sheet_names().iter().position(|n| *n == name)
}

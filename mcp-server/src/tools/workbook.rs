use std::path::Path;
use std::time::Instant;

use super::common::workbook_not_found;
use crate::engines::zavora;
use crate::store::{WorkbookEntry, WorkbookStore};
use crate::types::inputs::*;
use crate::types::responses::*;

pub fn create_workbook(store: &mut WorkbookStore) -> Result<String, anyhow::Error> {
    if store.is_full() {
        return Ok(error(ErrorCategory::CapacityExceeded, "Workbook store is at maximum capacity",
            "Save and close an existing workbook first to free a slot."));
    }
    let wb = zavora::create_workbook();
    let entry = WorkbookEntry { id: String::new(), data: wb, read_only: false, last_access: Instant::now() };
    let id = store.insert(entry).map_err(|e| anyhow::anyhow!("{}", e))?;
    let info = WorkbookInfo {
        workbook_id: id, engine: "zavora-xlsx".to_string(),
        sheets: vec![SheetSummary { name: "Sheet1".to_string(), dimensions: None, row_count: None, col_count: None }],
    };
    Ok(success("Workbook created successfully", info))
}

pub fn open_workbook(store: &mut WorkbookStore, input: OpenWorkbookInput) -> Result<String, anyhow::Error> {
    if store.is_full() {
        return Ok(error(ErrorCategory::CapacityExceeded, "Workbook store is at maximum capacity",
            "Save and close an existing workbook first to free a slot."));
    }
    let path = Path::new(&input.file_path);
    if !path.exists() {
        return Ok(error(ErrorCategory::NotFound, &format!("File not found: {}", input.file_path),
            "Check the file path and try again."));
    }
    let wb = match zavora::open_workbook(&input.file_path, input.read_only) {
        Ok(wb) => wb,
        Err(e) => return Ok(error(ErrorCategory::IoError, &format!("Failed to open file: {e}"),
            "Check the file is a valid Excel file and try again.")),
    };
    let sheets = zavora::sheet_summaries(&wb);
    let mode = if input.read_only { "read-only" } else { "edit" };
    let entry = WorkbookEntry { id: String::new(), data: wb, read_only: input.read_only, last_access: Instant::now() };
    let id = store.insert(entry).map_err(|e| anyhow::anyhow!("{}", e))?;
    Ok(success(&format!("Workbook opened in {mode} mode"), WorkbookInfo { workbook_id: id, engine: "zavora-xlsx".to_string(), sheets }))
}

pub fn save_workbook(store: &mut WorkbookStore, input: SaveWorkbookInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    if entry.read_only {
        return Ok(error(ErrorCategory::EngineUnsupported, "Read-only workbooks cannot be saved",
            "Reopen the file in edit mode (read_only: false) to make changes and save."));
    }
    match zavora::save_workbook(&mut entry.data, &input.file_path) {
        Ok(()) => Ok(success_no_data(&format!("Workbook saved to {}", input.file_path))),
        Err(e) => Ok(error(ErrorCategory::IoError, &format!("Failed to save: {e}"), "Check the file path is writable.")),
    }
}

pub fn close_workbook(store: &mut WorkbookStore, input: CloseWorkbookInput) -> Result<String, anyhow::Error> {
    match store.remove(&input.workbook_id) {
        Some(_) => Ok(success_no_data(&format!("Workbook '{}' closed successfully", input.workbook_id))),
        None => Ok(workbook_not_found(store, &input.workbook_id)),
    }
}

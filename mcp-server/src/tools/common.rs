//! Shared helpers for tool modules.

use crate::store::WorkbookStore;
use crate::types::responses::{error, ErrorCategory};

/// Build a not-found error response that includes the list of open workbook IDs.
pub fn workbook_not_found(store: &WorkbookStore, id: &str) -> String {
    let open = store.open_ids();
    let ids_str = if open.is_empty() {
        "none".to_string()
    } else {
        open.join(", ")
    };
    error(
        ErrorCategory::NotFound,
        &format!("Workbook '{}' not found", id),
        &format!(
            "Currently open workbook IDs: {}. The workbook may have been closed or evicted due to inactivity.",
            ids_str
        ),
    )
}

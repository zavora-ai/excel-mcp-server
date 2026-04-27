//! Example: Find and Replace
//!
//! Demonstrates using `find_replace` to search for and replace values across a sheet.

use excel_mcp_server::store::{WorkbookEntry, WorkbookStore};
use excel_mcp_server::tools::data::find_replace;
use excel_mcp_server::tools::workbook::save_workbook;
use excel_mcp_server::types::inputs::{FindReplaceInput, SaveWorkbookInput};
use std::time::Instant;

fn main() {
    println!("=== Find and Replace Example ===\n");
    std::fs::create_dir_all("output").ok();

    // Step 1: Create store and workbook
    let mut store = WorkbookStore::new();
    let entry = WorkbookEntry {
        id: String::new(),
        data: zavora_xlsx::Workbook::new(),
        read_only: false,
        last_access: Instant::now(),
    };
    let id = store.insert(entry).unwrap();
    println!("Created workbook with ID: {}", id);

    // Step 2: Write data with repeated values to replace
    {
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();

        ws.write(0, 0, "Company").unwrap();
        ws.write(0, 1, "Status").unwrap();
        ws.write(0, 2, "Notes").unwrap();

        ws.write(1, 0, "Acme Corp").unwrap();
        ws.write(1, 1, "Active").unwrap();
        ws.write(1, 2, "Good standing with Acme Corp").unwrap();

        ws.write(2, 0, "Acme Corp").unwrap();
        ws.write(2, 1, "Pending").unwrap();
        ws.write(2, 2, "Renewal for Acme Corp").unwrap();

        ws.write(3, 0, "Beta Inc").unwrap();
        ws.write(3, 1, "Active").unwrap();
        ws.write(3, 2, "New client").unwrap();

        ws.write(4, 0, "Acme Corp").unwrap();
        ws.write(4, 1, "Inactive").unwrap();
        ws.write(4, 2, "Old Acme Corp account").unwrap();
    }
    println!("Wrote data with 'Acme Corp' appearing multiple times");

    // Step 3: Replace "Acme Corp" with "Acme Industries"
    let input = FindReplaceInput {
        workbook_id: id.clone(),
        sheet_name: "Sheet1".to_string(),
        find: "Acme Corp".to_string(),
        replace: "Acme Industries".to_string(),
        range: None, // search entire sheet
        match_case: true,
    };

    let result = find_replace(&mut store, input).unwrap();
    println!("\nFind/replace result:\n{}", result);

    // Step 4: Save
    let save_result = save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: id.clone(),
            file_path: "output/find_replace_example.xlsx".to_string(),
        },
    )
    .unwrap();
    println!("\nSave result: {}", save_result);
    println!("\nDone! File saved to output/find_replace_example.xlsx");
}

//! Example: Copy Range
//!
//! Demonstrates using `copy_range` to copy data from one area to another,
//! both within the same sheet and across sheets.

use excel_mcp_server::store::{WorkbookEntry, WorkbookStore};
use excel_mcp_server::tools::data::copy_range;
use excel_mcp_server::tools::sheets::add_sheet;
use excel_mcp_server::tools::workbook::save_workbook;
use excel_mcp_server::types::inputs::{AddSheetInput, CopyRangeInput, SaveWorkbookInput};
use std::time::Instant;

fn main() {
    println!("=== Copy Range Example ===\n");
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

    // Step 2: Write source data in A1:C3
    {
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();

        ws.write(0, 0, "Region").unwrap();
        ws.write(0, 1, "Revenue").unwrap();
        ws.write(0, 2, "Profit").unwrap();

        ws.write(1, 0, "North").unwrap();
        ws.write(1, 1, 500000.0).unwrap();
        ws.write(1, 2, 150000.0).unwrap();

        ws.write(2, 0, "South").unwrap();
        ws.write(2, 1, 400000.0).unwrap();
        ws.write(2, 2, 120000.0).unwrap();
    }
    println!("Wrote source data in A1:C3");

    // Step 3: Copy range to another location on the same sheet (E1)
    let result = copy_range(
        &mut store,
        CopyRangeInput {
            workbook_id: id.clone(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:C3".to_string(),
            destination_sheet: None, // same sheet
            destination_cell: "E1".to_string(),
        },
    )
    .unwrap();
    println!("Same-sheet copy result: {}", result);

    // Step 4: Add a second sheet and copy data cross-sheet
    let _ = add_sheet(
        &mut store,
        AddSheetInput {
            workbook_id: id.clone(),
            sheet_name: "Summary".to_string(),
        },
    )
    .unwrap();
    println!("Added 'Summary' sheet");

    let result = copy_range(
        &mut store,
        CopyRangeInput {
            workbook_id: id.clone(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:C3".to_string(),
            destination_sheet: Some("Summary".to_string()),
            destination_cell: "A1".to_string(),
        },
    )
    .unwrap();
    println!("Cross-sheet copy result: {}", result);

    // Step 5: Save
    let save_result = save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: id.clone(),
            file_path: "output/copy_range_example.xlsx".to_string(),
        },
    )
    .unwrap();
    println!("\nSave result: {}", save_result);
    println!("\nDone! File saved to output/copy_range_example.xlsx");
}

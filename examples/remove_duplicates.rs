//! Example: Remove Duplicates
//!
//! Demonstrates using `remove_duplicates` to clean up data with duplicate rows.

use excel_mcp_server::store::{WorkbookEntry, WorkbookStore};
use excel_mcp_server::tools::data::remove_duplicates;
use excel_mcp_server::tools::workbook::save_workbook;
use excel_mcp_server::types::inputs::{RemoveDuplicatesInput, SaveWorkbookInput};
use std::time::Instant;

fn main() {
    println!("=== Remove Duplicates Example ===\n");
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

    // Step 2: Write data with duplicate rows
    {
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();

        // Headers
        ws.write(0, 0, "Email").unwrap();
        ws.write(0, 1, "Name").unwrap();
        ws.write(0, 2, "Department").unwrap();

        // Data with duplicates (same email = duplicate)
        let data = [
            ("alice@example.com", "Alice", "Engineering"),
            ("bob@example.com", "Bob", "Marketing"),
            ("alice@example.com", "Alice", "Engineering"), // duplicate
            ("carol@example.com", "Carol", "Sales"),
            ("bob@example.com", "Bob", "Marketing"), // duplicate
            ("dave@example.com", "Dave", "Operations"),
            ("alice@example.com", "Alice", "Engineering"), // duplicate
            ("eve@example.com", "Eve", "HR"),
        ];

        for (i, (email, name, dept)) in data.iter().enumerate() {
            let row = (i + 1) as u32;
            ws.write(row, 0, *email).unwrap();
            ws.write(row, 1, *name).unwrap();
            ws.write(row, 2, *dept).unwrap();
        }
    }
    println!("Wrote 8 rows (3 are duplicates)");

    // Step 3: Remove duplicates based on the Email column (A)
    let input = RemoveDuplicatesInput {
        workbook_id: id.clone(),
        sheet_name: "Sheet1".to_string(),
        range: "A1:C9".to_string(),
        columns: vec!["A".to_string()], // compare by email column only
        has_header: true,
    };

    let result = remove_duplicates(&mut store, input).unwrap();
    println!("\nRemove duplicates result:\n{}", result);

    // Step 4: Save
    let save_result = save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: id.clone(),
            file_path: "output/remove_duplicates_example.xlsx".to_string(),
        },
    )
    .unwrap();
    println!("\nSave result: {}", save_result);
    println!("\nDone! File saved to output/remove_duplicates_example.xlsx");
}

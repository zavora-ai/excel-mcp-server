//! Example: Apply Style Presets
//!
//! Demonstrates applying named style presets ("header", "currency", "percentage", "total")
//! to different ranges in a workbook.

use excel_mcp_server::store::{WorkbookEntry, WorkbookStore};
use excel_mcp_server::tools::format::apply_style;
use excel_mcp_server::tools::workbook::save_workbook;
use excel_mcp_server::types::inputs::{ApplyStyleInput, SaveWorkbookInput};
use std::time::Instant;

fn main() {
    println!("=== Apply Style Presets Example ===\n");
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

    // Step 2: Write sample data
    {
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();

        // Headers
        ws.write(0, 0, "Item").unwrap();
        ws.write(0, 1, "Amount").unwrap();
        ws.write(0, 2, "Tax Rate").unwrap();
        ws.write(0, 3, "Total").unwrap();

        // Data rows
        ws.write(1, 0, "Laptop").unwrap();
        ws.write(1, 1, 1299.99).unwrap();
        ws.write(1, 2, 0.08).unwrap();
        ws.write(1, 3, 1403.99).unwrap();

        ws.write(2, 0, "Monitor").unwrap();
        ws.write(2, 1, 499.99).unwrap();
        ws.write(2, 2, 0.08).unwrap();
        ws.write(2, 3, 539.99).unwrap();

        ws.write(3, 0, "Keyboard").unwrap();
        ws.write(3, 1, 79.99).unwrap();
        ws.write(3, 2, 0.08).unwrap();
        ws.write(3, 3, 86.39).unwrap();

        // Total row
        ws.write(4, 0, "Total").unwrap();
        ws.write(4, 1, 1879.97).unwrap();
        ws.write(4, 2, 0.08).unwrap();
        ws.write(4, 3, 2030.37).unwrap();
    }
    println!("Wrote sample data");

    // Step 3: Apply "header" style to the header row
    let result = apply_style(
        &mut store,
        ApplyStyleInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            range: "A1:D1".to_string(),
            style: "header".to_string(),
        },
    )
    .unwrap();
    println!("Header style: {}", result);

    // Step 4: Apply "currency" style to amount and total columns
    let result = apply_style(
        &mut store,
        ApplyStyleInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            range: "B2:B5,D2:D5".to_string(), // comma-separated ranges
            style: "currency".to_string(),
        },
    )
    .unwrap();
    println!("Currency style: {}", result);

    // Step 5: Apply "percentage" style to tax rate column
    let result = apply_style(
        &mut store,
        ApplyStyleInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            range: "C2:C5".to_string(),
            style: "percentage".to_string(),
        },
    )
    .unwrap();
    println!("Percentage style: {}", result);

    // Step 6: Apply "total" style to the total row
    let result = apply_style(
        &mut store,
        ApplyStyleInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            range: "A5:D5".to_string(),
            style: "total".to_string(),
        },
    )
    .unwrap();
    println!("Total style: {}", result);

    // Step 7: Save
    let save_result = save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: id.clone(),
            file_path: "output/apply_style_example.xlsx".to_string(),
        },
    )
    .unwrap();
    println!("\nSave result: {}", save_result);
    println!("\nDone! File saved to output/apply_style_example.xlsx");
}

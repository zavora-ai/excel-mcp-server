//! Example: Write Grid
//!
//! Demonstrates using `write_grid` to write a 2D block of mixed data
//! (numbers, strings, formulas, booleans) in a single call.

use excel_mcp_server::store::{WorkbookEntry, WorkbookStore};
use excel_mcp_server::tools::workbook::save_workbook;
use excel_mcp_server::tools::write::write_grid;
use excel_mcp_server::types::inputs::{SaveWorkbookInput, WriteGridInput};
use std::time::Instant;

fn main() {
    println!("=== Write Grid Example ===\n");
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

    // Step 2: Write a 2D grid of mixed data types using write_grid
    let input = WriteGridInput {
        workbook_id: id.clone(),
        sheet_name: "Sheet1".to_string(),
        start_cell: "A1".to_string(),
        rows: vec![
            // Row 1: Headers (strings)
            vec![
                serde_json::json!("Product"),
                serde_json::json!("Price"),
                serde_json::json!("Quantity"),
                serde_json::json!("In Stock"),
                serde_json::json!("Total"),
            ],
            // Row 2: Mixed types — string, numbers, boolean, formula
            vec![
                serde_json::json!("Widget A"),
                serde_json::json!(29.99),
                serde_json::json!(100),
                serde_json::json!(true),
                serde_json::json!("=B2*C2"),
            ],
            // Row 3
            vec![
                serde_json::json!("Widget B"),
                serde_json::json!(49.99),
                serde_json::json!(50),
                serde_json::json!(true),
                serde_json::json!("=B3*C3"),
            ],
            // Row 4
            vec![
                serde_json::json!("Widget C"),
                serde_json::json!(9.99),
                serde_json::json!(200),
                serde_json::json!(false),
                serde_json::json!("=B4*C4"),
            ],
            // Row 5: Grand total formula
            vec![
                serde_json::json!("Grand Total"),
                serde_json::json!(null),
                serde_json::json!(null),
                serde_json::json!(null),
                serde_json::json!("=SUM(E2:E4)"),
            ],
        ],
    };

    let result = write_grid(&mut store, input).unwrap();
    println!("Write grid result:\n{}", result);

    // Step 3: Save
    let save_result = save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: id.clone(),
            file_path: "output/write_grid_example.xlsx".to_string(),
        },
    )
    .unwrap();
    println!("\nSave result: {}", save_result);
    println!("\nDone! File saved to output/write_grid_example.xlsx");
}

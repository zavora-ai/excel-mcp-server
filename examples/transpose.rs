//! Example: Transpose Range
//!
//! Demonstrates using `transpose_range` to flip rows and columns of a data range.

use excel_mcp_server::store::{WorkbookEntry, WorkbookStore};
use excel_mcp_server::tools::data::transpose_range;
use excel_mcp_server::tools::workbook::save_workbook;
use excel_mcp_server::types::inputs::{SaveWorkbookInput, TransposeRangeInput};
use std::time::Instant;

fn main() {
    println!("=== Transpose Range Example ===\n");
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

    // Step 2: Write a 3x4 data range (3 rows, 4 columns)
    {
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();

        // Original data in A1:D3
        ws.write(0, 0, "Name").unwrap();
        ws.write(0, 1, "Q1").unwrap();
        ws.write(0, 2, "Q2").unwrap();
        ws.write(0, 3, "Q3").unwrap();

        ws.write(1, 0, "Alice").unwrap();
        ws.write(1, 1, 100.0).unwrap();
        ws.write(1, 2, 120.0).unwrap();
        ws.write(1, 3, 140.0).unwrap();

        ws.write(2, 0, "Bob").unwrap();
        ws.write(2, 1, 90.0).unwrap();
        ws.write(2, 2, 110.0).unwrap();
        ws.write(2, 3, 130.0).unwrap();
    }
    println!("Wrote 3x4 data range in A1:D3");

    // Step 3: Transpose to a new location (F1)
    // The 3x4 range becomes a 4x3 range at F1:H4
    let input = TransposeRangeInput {
        workbook_id: id.clone(),
        sheet_name: "Sheet1".to_string(),
        source_range: "A1:D3".to_string(),
        destination_cell: Some("F1".to_string()),
    };

    let result = transpose_range(&mut store, input).unwrap();
    println!("\nTranspose result:\n{}", result);

    // Step 4: Save
    let save_result = save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: id.clone(),
            file_path: "output/transpose_example.xlsx".to_string(),
        },
    )
    .unwrap();
    println!("\nSave result: {}", save_result);
    println!("\nDone! File saved to output/transpose_example.xlsx");
}

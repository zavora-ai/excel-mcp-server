//! Example: Split Column
//!
//! Demonstrates using `split_column` to split comma-separated values in a column
//! into multiple columns.

use excel_mcp_server::store::{WorkbookEntry, WorkbookStore};
use excel_mcp_server::tools::data::split_column;
use excel_mcp_server::tools::workbook::save_workbook;
use excel_mcp_server::types::inputs::{SaveWorkbookInput, SplitColumnInput};
use std::time::Instant;

fn main() {
    println!("=== Split Column Example ===\n");
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

    // Step 2: Write data with comma-separated values in column B
    {
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();

        // Headers
        ws.write(0, 0, "Employee").unwrap();
        ws.write(0, 1, "Skills").unwrap();

        // Data with comma-separated skills
        ws.write(1, 0, "Alice").unwrap();
        ws.write(1, 1, "Python, JavaScript, SQL").unwrap();

        ws.write(2, 0, "Bob").unwrap();
        ws.write(2, 1, "Java, C++").unwrap();

        ws.write(3, 0, "Carol").unwrap();
        ws.write(3, 1, "Rust, Go, Python, TypeScript").unwrap();

        ws.write(4, 0, "Dave").unwrap();
        ws.write(4, 1, "Excel").unwrap(); // single value, no split needed
    }
    println!("Wrote data with comma-separated skills");

    // Step 3: Split the Skills column (B) by comma
    let input = SplitColumnInput {
        workbook_id: id.clone(),
        sheet_name: "Sheet1".to_string(),
        column: "B".to_string(),
        start_row: 1, // 1-based
        end_row: 5,   // 1-based
        delimiter: ",".to_string(),
        has_header: true,
    };

    let result = split_column(&mut store, input).unwrap();
    println!("\nSplit column result:\n{}", result);

    // Step 4: Save
    let save_result = save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: id.clone(),
            file_path: "output/split_column_example.xlsx".to_string(),
        },
    )
    .unwrap();
    println!("\nSave result: {}", save_result);
    println!("\nDone! File saved to output/split_column_example.xlsx");
}

//! Example: Sort Range
//!
//! Demonstrates using `sort_range` to sort data by one or more columns.

use excel_mcp_server::store::{WorkbookEntry, WorkbookStore};
use excel_mcp_server::tools::data::sort_range;
use excel_mcp_server::tools::workbook::save_workbook;
use excel_mcp_server::types::enums::SortDirection;
use excel_mcp_server::types::inputs::{SaveWorkbookInput, SortKey, SortRangeInput};
use std::time::Instant;

fn main() {
    println!("=== Sort Range Example ===\n");
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

    // Step 2: Write unsorted data
    {
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();

        // Headers
        ws.write(0, 0, "Name").unwrap();
        ws.write(0, 1, "Department").unwrap();
        ws.write(0, 2, "Salary").unwrap();

        // Unsorted data
        let data = [
            ("Charlie", "Sales", 72000.0),
            ("Alice", "Engineering", 95000.0),
            ("Eve", "Engineering", 88000.0),
            ("Bob", "Marketing", 65000.0),
            ("Diana", "Sales", 70000.0),
            ("Frank", "Marketing", 78000.0),
        ];

        for (i, (name, dept, salary)) in data.iter().enumerate() {
            let row = (i + 1) as u32;
            ws.write(row, 0, *name).unwrap();
            ws.write(row, 1, *dept).unwrap();
            ws.write(row, 2, *salary).unwrap();
        }
    }
    println!("Wrote unsorted employee data");

    // Step 3: Sort by Department (ascending), then by Salary (descending)
    let input = SortRangeInput {
        workbook_id: id.clone(),
        sheet_name: "Sheet1".to_string(),
        range: "A1:C7".to_string(),
        sort_keys: vec![
            SortKey {
                column: "B".to_string(),
                direction: Some(SortDirection::Ascending),
            },
            SortKey {
                column: "C".to_string(),
                direction: Some(SortDirection::Descending),
            },
        ],
        has_header: true,
    };

    let result = sort_range(&mut store, input).unwrap();
    println!("\nSort result:\n{}", result);

    // Step 4: Save
    let save_result = save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: id.clone(),
            file_path: "output/sort_range_example.xlsx".to_string(),
        },
    )
    .unwrap();
    println!("\nSave result: {}", save_result);
    println!("\nDone! File saved to output/sort_range_example.xlsx");
}

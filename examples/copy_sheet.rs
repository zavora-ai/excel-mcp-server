//! Example: Copy Sheet
//!
//! Demonstrates using `copy_sheet` to duplicate a sheet with all its data.

use excel_mcp_server::store::{WorkbookEntry, WorkbookStore};
use excel_mcp_server::tools::sheets::copy_sheet;
use excel_mcp_server::tools::workbook::save_workbook;
use excel_mcp_server::types::inputs::{CopySheetInput, SaveWorkbookInput};
use std::time::Instant;

fn main() {
    println!("=== Copy Sheet Example ===\n");
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

    // Step 2: Write data to Sheet1
    {
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();

        ws.write(0, 0, "Month").unwrap();
        ws.write(0, 1, "Sales").unwrap();
        ws.write(0, 2, "Target").unwrap();
        ws.write(0, 3, "Achievement").unwrap();

        let months = ["January", "February", "March", "April", "May", "June"];
        let sales = [45000.0, 52000.0, 48000.0, 61000.0, 55000.0, 67000.0];
        let targets = [50000.0, 50000.0, 50000.0, 55000.0, 55000.0, 60000.0];

        for (i, ((month, sale), target)) in months
            .iter()
            .zip(sales.iter())
            .zip(targets.iter())
            .enumerate()
        {
            let row = (i + 1) as u32;
            ws.write(row, 0, *month).unwrap();
            ws.write(row, 1, *sale).unwrap();
            ws.write(row, 2, *target).unwrap();
            // Achievement formula
            ws.write_formula(row, 3, &format!("B{}/C{}", row + 1, row + 1))
                .unwrap();
        }
    }
    println!("Wrote sales data to Sheet1");

    // Step 3: Copy Sheet1 to a new sheet called "H1 Report"
    let input = CopySheetInput {
        workbook_id: id.clone(),
        source_sheet: "Sheet1".to_string(),
        new_sheet_name: "H1 Report".to_string(),
    };

    let result = copy_sheet(&mut store, input).unwrap();
    println!("\nCopy sheet result:\n{}", result);

    // Step 4: Save
    let save_result = save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: id.clone(),
            file_path: "output/copy_sheet_example.xlsx".to_string(),
        },
    )
    .unwrap();
    println!("\nSave result: {}", save_result);
    println!("\nDone! File saved to output/copy_sheet_example.xlsx");
}

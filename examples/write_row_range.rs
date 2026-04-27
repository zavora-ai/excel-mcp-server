//! Example: Write Row Range (Drag-Fill Formulas)
//!
//! Demonstrates using `write_row_range` to fill a formula across columns
//! with automatic reference adjustment — like Excel's drag-fill behavior.

use excel_mcp_server::store::{WorkbookEntry, WorkbookStore};
use excel_mcp_server::tools::workbook::save_workbook;
use excel_mcp_server::tools::write::write_row_range;
use excel_mcp_server::types::inputs::{SaveWorkbookInput, WriteRowRangeInput};
use std::time::Instant;

fn main() {
    println!("=== Write Row Range (Drag-Fill) Example ===\n");
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

    // Step 2: Write seed data for a financial projection
    {
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();

        // Row labels (column A)
        ws.write(0, 0, "Metric").unwrap();
        ws.write(1, 0, "Revenue").unwrap();
        ws.write(2, 0, "Growth Rate").unwrap();
        ws.write(3, 0, "Expenses").unwrap();
        ws.write(4, 0, "Profit").unwrap();

        // Year headers (row 0)
        ws.write(0, 1, "Year 1").unwrap();
        ws.write(0, 2, "Year 2").unwrap();
        ws.write(0, 3, "Year 3").unwrap();
        ws.write(0, 4, "Year 4").unwrap();
        ws.write(0, 5, "Year 5").unwrap();

        // Year 1 base values (column B)
        ws.write(1, 1, 100000.0).unwrap(); // Revenue
        ws.write(2, 1, 0.10).unwrap();     // Growth rate
        ws.write(3, 1, 60000.0).unwrap();  // Expenses
        ws.write(4, 1, 40000.0).unwrap();  // Profit
    }
    println!("Wrote seed data for Year 1");

    // Step 3: Use write_row_range to fill revenue projection across Years 2-5
    // Formula: previous year revenue * (1 + growth rate)
    // At B2 (row 1, col 1): =B2*(1+B3) → fills to C2, D2, E2, F2
    // The column references adjust: C2*(1+C3), D2*(1+D3), etc.
    let result = write_row_range(
        &mut store,
        WriteRowRangeInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            start_cell: "C2".to_string(),
            end_column: "F".to_string(),
            formula: "=B2*(1+B3)".to_string(),
        },
    )
    .unwrap();
    println!("Revenue projection fill: {}", result);

    // Step 4: Fill profit formula (Revenue - Expenses) across Years 2-5
    let result = write_row_range(
        &mut store,
        WriteRowRangeInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            start_cell: "C5".to_string(),
            end_column: "F".to_string(),
            formula: "=C2-C4".to_string(),
        },
    )
    .unwrap();
    println!("Profit formula fill: {}", result);

    // Step 5: Save
    let save_result = save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: id.clone(),
            file_path: "output/write_row_range_example.xlsx".to_string(),
        },
    )
    .unwrap();
    println!("\nSave result: {}", save_result);
    println!("\nDone! File saved to output/write_row_range_example.xlsx");
}

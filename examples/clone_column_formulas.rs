//! Example: Clone Column Formulas
//!
//! Demonstrates using `clone_column_formulas` to replicate formulas from one column
//! to multiple target columns with automatic reference adjustment.
//! Creates Year 1 formulas in one column, then clones them across Years 2-5.

use excel_mcp_server::store::{WorkbookEntry, WorkbookStore};
use excel_mcp_server::tools::workbook::save_workbook;
use excel_mcp_server::tools::write::clone_column_formulas;
use excel_mcp_server::types::inputs::{CloneColumnFormulasInput, SaveWorkbookInput};
use std::time::Instant;

fn main() {
    println!("=== Clone Column Formulas Example ===\n");
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

    // Step 2: Write a financial model structure
    {
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();

        // Row labels (column A)
        ws.write(0, 0, "Metric").unwrap();
        ws.write(1, 0, "Revenue").unwrap();
        ws.write(2, 0, "COGS").unwrap();
        ws.write(3, 0, "Gross Profit").unwrap();
        ws.write(4, 0, "OpEx").unwrap();
        ws.write(5, 0, "EBITDA").unwrap();

        // Year headers
        ws.write(0, 1, "Year 1").unwrap();
        ws.write(0, 2, "Year 2").unwrap();
        ws.write(0, 3, "Year 3").unwrap();
        ws.write(0, 4, "Year 4").unwrap();
        ws.write(0, 5, "Year 5").unwrap();

        // Year 1 base values (column B)
        ws.write(1, 1, 1000000.0).unwrap(); // Revenue
        ws.write(2, 1, 400000.0).unwrap(); // COGS

        // Year 1 formulas (column B)
        ws.write_formula(3, 1, "B2-B3").unwrap(); // Gross Profit = Revenue - COGS
        ws.write(4, 1, 200000.0).unwrap(); // OpEx (fixed value)
        ws.write_formula(5, 1, "B4-B5").unwrap(); // EBITDA = Gross Profit - OpEx

        // Year 2+ base values (columns C-F) — revenue and COGS grow
        for col in 2u16..=5 {
            ws.write(1, col, 1000000.0 * (1.0 + 0.1 * (col - 1) as f64))
                .unwrap();
            ws.write(2, col, 400000.0 * (1.0 + 0.05 * (col - 1) as f64))
                .unwrap();
            ws.write(4, col, 200000.0 * (1.0 + 0.03 * (col - 1) as f64))
                .unwrap();
        }
    }
    println!("Wrote financial model with Year 1 formulas in column B");

    // Step 3: Clone formulas from column B to columns C, D, E, F
    // This will adjust references: B2-B3 → C2-C3, D2-D3, etc.
    let input = CloneColumnFormulasInput {
        workbook_id: id.clone(),
        sheet_name: "Sheet1".to_string(),
        source_column: "B".to_string(),
        target_columns: vec![
            "C".to_string(),
            "D".to_string(),
            "E".to_string(),
            "F".to_string(),
        ],
        start_row: 1, // 1-based
        end_row: 6,   // 1-based
    };

    let result = clone_column_formulas(&mut store, input).unwrap();
    println!("\nClone column formulas result:\n{}", result);

    // Step 4: Save
    let save_result = save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: id.clone(),
            file_path: "output/clone_formulas_example.xlsx".to_string(),
        },
    )
    .unwrap();
    println!("\nSave result: {}", save_result);
    println!("\nDone! File saved to output/clone_formulas_example.xlsx");
}

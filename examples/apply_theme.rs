//! Example: Apply Theme
//!
//! Demonstrates applying a complete professional theme to a financial report sheet.
//! Uses the "financial_professional" theme with header and total row styling.

use excel_mcp_server::store::{WorkbookEntry, WorkbookStore};
use excel_mcp_server::tools::format::apply_theme;
use excel_mcp_server::tools::workbook::save_workbook;
use excel_mcp_server::types::inputs::{ApplyThemeInput, SaveWorkbookInput};
use std::time::Instant;

fn main() {
    println!("=== Apply Theme Example ===\n");
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

    // Step 2: Write a financial report with headers, data, and totals
    {
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();

        // Header row (row 1 = index 0)
        ws.write(0, 0, "Department").unwrap();
        ws.write(0, 1, "Budget").unwrap();
        ws.write(0, 2, "Actual").unwrap();
        ws.write(0, 3, "Variance").unwrap();

        // Data rows
        let departments = ["Engineering", "Marketing", "Sales", "Operations", "HR"];
        let budgets = [500000.0, 200000.0, 300000.0, 150000.0, 100000.0];
        let actuals = [480000.0, 220000.0, 310000.0, 140000.0, 95000.0];

        for (i, (dept, (budget, actual))) in departments
            .iter()
            .zip(budgets.iter().zip(actuals.iter()))
            .enumerate()
        {
            let row = (i + 1) as u32;
            ws.write(row, 0, *dept).unwrap();
            ws.write(row, 1, *budget).unwrap();
            ws.write(row, 2, *actual).unwrap();
            ws.write(row, 3, *actual - *budget).unwrap();
        }

        // Total row (row 7 = index 6)
        ws.write(6, 0, "Total").unwrap();
        ws.write(6, 1, 1250000.0).unwrap();
        ws.write(6, 2, 1245000.0).unwrap();
        ws.write(6, 3, -5000.0).unwrap();
    }
    println!("Wrote financial report data");

    // Step 3: Apply the "financial_professional" theme
    let input = ApplyThemeInput {
        workbook_id: id.clone(),
        sheet_name: "Sheet1".to_string(),
        theme: "financial_professional".to_string(),
        header_rows: vec![1],  // 1-based: row 1 is the header
        total_rows: vec![7],   // 1-based: row 7 is the total
        auto_detect_formats: false,
    };

    let result = apply_theme(&mut store, input).unwrap();
    println!("\nApply theme result:\n{}", result);

    // Step 4: Save
    let save_result = save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: id.clone(),
            file_path: "output/apply_theme_example.xlsx".to_string(),
        },
    )
    .unwrap();
    println!("\nSave result: {}", save_result);
    println!("\nDone! File saved to output/apply_theme_example.xlsx");
}

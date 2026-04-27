//! Example: Batch Format
//!
//! Demonstrates applying multiple formatting operations in a single call using `batch_format`.
//! Creates a workbook with financial data, applies bold headers, currency formatting on numbers,
//! and borders — all in one batch operation.

use excel_mcp_server::store::{WorkbookEntry, WorkbookStore};
use excel_mcp_server::tools::format::batch_format;
use excel_mcp_server::tools::workbook::save_workbook;
use excel_mcp_server::types::inputs::{
    BatchFormatInput, FormatOperation, SaveWorkbookInput,
};
use std::time::Instant;

fn main() {
    println!("=== Batch Format Example ===\n");
    std::fs::create_dir_all("output").ok();

    // Step 1: Create a WorkbookStore and insert a new workbook
    let mut store = WorkbookStore::new();
    let entry = WorkbookEntry {
        id: String::new(),
        data: zavora_xlsx::Workbook::new(),
        read_only: false,
        last_access: Instant::now(),
    };
    let id = store.insert(entry).unwrap();
    println!("Created workbook with ID: {}", id);

    // Step 2: Write sample financial data
    {
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();

        // Headers (row 0)
        ws.write(0, 0, "Category").unwrap();
        ws.write(0, 1, "Q1").unwrap();
        ws.write(0, 2, "Q2").unwrap();
        ws.write(0, 3, "Q3").unwrap();
        ws.write(0, 4, "Q4").unwrap();

        // Data rows
        ws.write(1, 0, "Revenue").unwrap();
        ws.write(1, 1, 150000.0).unwrap();
        ws.write(1, 2, 175000.0).unwrap();
        ws.write(1, 3, 200000.0).unwrap();
        ws.write(1, 4, 225000.0).unwrap();

        ws.write(2, 0, "Expenses").unwrap();
        ws.write(2, 1, 80000.0).unwrap();
        ws.write(2, 2, 85000.0).unwrap();
        ws.write(2, 3, 90000.0).unwrap();
        ws.write(2, 4, 95000.0).unwrap();

        ws.write(3, 0, "Profit").unwrap();
        ws.write(3, 1, 70000.0).unwrap();
        ws.write(3, 2, 90000.0).unwrap();
        ws.write(3, 3, 110000.0).unwrap();
        ws.write(3, 4, 130000.0).unwrap();
    }
    println!("Wrote financial data (headers + 3 data rows)");

    // Step 3: Apply batch formatting — bold headers, currency numbers, borders
    let input = BatchFormatInput {
        workbook_id: id.clone(),
        sheet_name: "Sheet1".to_string(),
        operations: vec![
            // Operation 1: Bold headers with blue background and white font
            FormatOperation {
                range: "A1:E1".to_string(),
                bold: Some(true),
                italic: None,
                underline: None,
                font_size: Some(12.0),
                font_color: Some("#FFFFFF".to_string()),
                background_color: Some("#4472C4".to_string()),
                number_format: None,
                horizontal_alignment: None,
                vertical_alignment: None,
                border_style: None,
            },
            // Operation 2: Currency format on numeric cells
            FormatOperation {
                range: "B2:E4".to_string(),
                bold: None,
                italic: None,
                underline: None,
                font_size: None,
                font_color: None,
                background_color: None,
                number_format: Some("currency".to_string()), // semantic format
                horizontal_alignment: None,
                vertical_alignment: None,
                border_style: None,
            },
            // Operation 3: Thin borders on the entire table
            FormatOperation {
                range: "A1:E4".to_string(),
                bold: None,
                italic: None,
                underline: None,
                font_size: None,
                font_color: None,
                background_color: None,
                number_format: None,
                horizontal_alignment: None,
                vertical_alignment: None,
                border_style: Some(excel_mcp_server::types::enums::BorderStyle::Thin),
            },
        ],
    };

    let result = batch_format(&mut store, input).unwrap();
    println!("\nBatch format result:\n{}", result);

    // Step 4: Save to file
    let save_result = save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: id.clone(),
            file_path: "output/batch_format_example.xlsx".to_string(),
        },
    )
    .unwrap();
    println!("\nSave result: {}", save_result);
    println!("\nDone! File saved to output/batch_format_example.xlsx");
}

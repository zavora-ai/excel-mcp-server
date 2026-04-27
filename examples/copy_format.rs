//! Example: Copy Format
//!
//! Demonstrates copying formatting from one range to multiple target ranges.
//! Formats a header row with bold + colors, then replicates that formatting
//! to other rows using `copy_format`.

use excel_mcp_server::store::{WorkbookEntry, WorkbookStore};
use excel_mcp_server::tools::format::{copy_format, set_cell_format};
use excel_mcp_server::tools::workbook::save_workbook;
use excel_mcp_server::types::inputs::{
    CopyFormatInput, SaveWorkbookInput, SetCellFormatInput,
};
use std::time::Instant;

fn main() {
    println!("=== Copy Format Example ===\n");
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

    // Step 2: Write data — multiple sections each with a header row
    {
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();

        // Section 1 header (row 0)
        ws.write(0, 0, "Product").unwrap();
        ws.write(0, 1, "Price").unwrap();
        ws.write(0, 2, "Quantity").unwrap();
        // Section 1 data
        ws.write(1, 0, "Widget A").unwrap();
        ws.write(1, 1, 29.99).unwrap();
        ws.write(1, 2, 100.0).unwrap();

        // Section 2 header (row 3)
        ws.write(3, 0, "Service").unwrap();
        ws.write(3, 1, "Rate").unwrap();
        ws.write(3, 2, "Hours").unwrap();
        // Section 2 data
        ws.write(4, 0, "Consulting").unwrap();
        ws.write(4, 1, 150.0).unwrap();
        ws.write(4, 2, 40.0).unwrap();

        // Section 3 header (row 6)
        ws.write(6, 0, "Region").unwrap();
        ws.write(6, 1, "Revenue").unwrap();
        ws.write(6, 2, "Growth").unwrap();
        // Section 3 data
        ws.write(7, 0, "North").unwrap();
        ws.write(7, 1, 500000.0).unwrap();
        ws.write(7, 2, 0.12).unwrap();
    }
    println!("Wrote data with 3 section headers");

    // Step 3: Format the first header row (A1:C1) with bold + colors
    let fmt_result = set_cell_format(
        &mut store,
        SetCellFormatInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            range: "A1:C1".to_string(),
            bold: Some(true),
            italic: None,
            underline: None,
            font_size: Some(11.0),
            font_color: Some("#FFFFFF".to_string()),
            background_color: Some("#2E75B6".to_string()),
            number_format: None,
            horizontal_alignment: Some(excel_mcp_server::types::enums::HorizontalAlignment::Center),
            vertical_alignment: None,
            border_style: None,
        },
    )
    .unwrap();
    println!("Formatted source header: {}", fmt_result);

    // Step 4: Save and reopen so cell_format() can read the formatting back
    save_workbook(&mut store, SaveWorkbookInput {
        workbook_id: id.clone(),
        file_path: "output/copy_format_example.xlsx".into(),
    }).unwrap();

    let open_result = excel_mcp_server::tools::workbook::open_workbook(
        &mut store,
        excel_mcp_server::types::inputs::OpenWorkbookInput {
            file_path: "output/copy_format_example.xlsx".into(),
            read_only: false,
        },
    ).unwrap();
    let v: serde_json::Value = serde_json::from_str(&open_result).unwrap();
    let id = v["data"]["workbook_id"].as_str().unwrap().to_string();
    println!("Reopened workbook with ID: {}", id);

    // Step 5: Copy that formatting to the other two header rows
    let input = CopyFormatInput {
        workbook_id: id.clone(),
        sheet_name: "Sheet1".to_string(),
        source_range: "A1:C1".to_string(),
        target_ranges: vec!["A4:C4".to_string(), "A7:C7".to_string()],
    };

    let result = copy_format(&mut store, input).unwrap();
    println!("\nCopy format result:\n{}", result);

    // Step 6: Save
    let save_result = save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: id.clone(),
            file_path: "output/copy_format_example.xlsx".to_string(),
        },
    )
    .unwrap();
    println!("\nSave result: {}", save_result);
    println!("\nDone! File saved to output/copy_format_example.xlsx");
}

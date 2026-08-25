//! Example: Describe Formatting
//!
//! Demonstrates using `describe_formatting` to read back formatting information
//! from a range. Creates a workbook, applies various formatting, then inspects it.

use excel_mcp_server::store::{WorkbookEntry, WorkbookStore};
use excel_mcp_server::tools::format::set_cell_format;
use excel_mcp_server::tools::read::describe_formatting;
use excel_mcp_server::tools::workbook::save_workbook;
use excel_mcp_server::types::inputs::{
    DescribeFormattingInput, SaveWorkbookInput, SetCellFormatInput,
};
use std::time::Instant;

fn main() {
    println!("=== Describe Formatting Example ===\n");
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

    // Step 2: Write some data
    {
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();

        ws.write(0, 0, "Header A").unwrap();
        ws.write(0, 1, "Header B").unwrap();
        ws.write(0, 2, "Header C").unwrap();
        ws.write(1, 0, 1000.0).unwrap();
        ws.write(1, 1, 2000.0).unwrap();
        ws.write(1, 2, 3000.0).unwrap();
        ws.write(2, 0, 0.15).unwrap();
        ws.write(2, 1, 0.25).unwrap();
        ws.write(2, 2, 0.35).unwrap();
    }
    println!("Wrote sample data");

    // Step 3: Apply different formatting to different ranges
    // Bold headers with blue background
    let _ = set_cell_format(
        &mut store,
        SetCellFormatInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            range: "A1:C1".to_string(),
            bold: Some(true),
            italic: None,
            underline: None,
            font_size: None,
            font_color: Some("#FFFFFF".to_string()),
            background_color: Some("#4472C4".to_string()),
            number_format: None,
            horizontal_alignment: None,
            vertical_alignment: None,
            border_style: None,
        },
    )
    .unwrap();
    println!("Applied bold + blue background to headers");

    // Currency format on row 2
    let _ = set_cell_format(
        &mut store,
        SetCellFormatInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            range: "A2:C2".to_string(),
            bold: None,
            italic: None,
            underline: None,
            font_size: None,
            font_color: None,
            background_color: None,
            number_format: Some("currency".to_string()),
            horizontal_alignment: None,
            vertical_alignment: None,
            border_style: None,
        },
    )
    .unwrap();
    println!("Applied currency format to row 2");

    // Percentage format on row 3
    let _ = set_cell_format(
        &mut store,
        SetCellFormatInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            range: "A3:C3".to_string(),
            bold: None,
            italic: None,
            underline: None,
            font_size: None,
            font_color: None,
            background_color: None,
            number_format: Some("percentage".to_string()),
            horizontal_alignment: None,
            vertical_alignment: None,
            border_style: None,
        },
    )
    .unwrap();
    println!("Applied percentage format to row 3");

    // Step 4: Save and reopen so cell_format() can read the formatting back
    save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: id.clone(),
            file_path: "output/describe_formatting_example.xlsx".into(),
        },
    )
    .unwrap();

    let open_result = excel_mcp_server::tools::workbook::open_workbook(
        &mut store,
        excel_mcp_server::types::inputs::OpenWorkbookInput {
            file_path: "output/describe_formatting_example.xlsx".into(),
            read_only: false,
        },
    )
    .unwrap();
    let v: serde_json::Value = serde_json::from_str(&open_result).unwrap();
    let id = v["data"]["workbook_id"].as_str().unwrap().to_string();
    println!("Reopened workbook with ID: {}", id);

    // Step 5: Describe the formatting on the entire range
    let result = describe_formatting(
        &mut store,
        DescribeFormattingInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            range: "A1:C3".to_string(),
        },
    )
    .unwrap();

    println!("\n=== Formatting Description ===");
    // Pretty-print the JSON
    if let Ok(parsed) = serde_json::from_str::<serde_json::Value>(&result) {
        println!("{}", serde_json::to_string_pretty(&parsed).unwrap());
    } else {
        println!("{}", result);
    }

    println!("\nDone! File saved to output/describe_formatting_example.xlsx");
}

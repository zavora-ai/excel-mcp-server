//! Example: Fill Series
//!
//! Demonstrates using `fill_series` with linear, copy, and date fill types
//! to extend patterns from seed values.

use excel_mcp_server::store::{WorkbookEntry, WorkbookStore};
use excel_mcp_server::tools::data::fill_series;
use excel_mcp_server::tools::workbook::save_workbook;
use excel_mcp_server::types::enums::{FillDirection, FillType};
use excel_mcp_server::types::inputs::{FillSeriesInput, SaveWorkbookInput};
use std::time::Instant;

fn main() {
    println!("=== Fill Series Example ===\n");
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

    // Step 2: Write seed values
    {
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();

        // Column A: Linear numeric series (seed: 10, 20)
        ws.write(0, 0, "Linear").unwrap();
        ws.write(1, 0, 10.0).unwrap();
        ws.write(2, 0, 20.0).unwrap();

        // Column B: Date series (seed: two dates)
        ws.write(0, 1, "Dates").unwrap();
        ws.write(1, 1, "2024-01-01").unwrap();
        ws.write(2, 1, "2024-01-08").unwrap(); // weekly interval

        // Column C: Copy series (seed: A, B, C)
        ws.write(0, 2, "Copy").unwrap();
        ws.write(1, 2, "Red").unwrap();
        ws.write(2, 2, "Green").unwrap();
        ws.write(3, 2, "Blue").unwrap();
    }
    println!("Wrote seed values for 3 series types");

    // Step 3: Fill linear series (extend 10, 20 → 30, 40, 50, 60, 70)
    let result = fill_series(
        &mut store,
        FillSeriesInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            source_range: "A2:A3".to_string(),
            fill_count: 5,
            direction: Some(FillDirection::Down),
            fill_type: Some(FillType::Linear),
        },
    )
    .unwrap();
    println!("Linear fill result: {}", result);

    // Step 4: Fill date series (extend weekly dates)
    let result = fill_series(
        &mut store,
        FillSeriesInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            source_range: "B2:B3".to_string(),
            fill_count: 5,
            direction: Some(FillDirection::Down),
            fill_type: Some(FillType::Date),
        },
    )
    .unwrap();
    println!("Date fill result: {}", result);

    // Step 5: Fill copy series (repeat Red, Green, Blue cyclically)
    let result = fill_series(
        &mut store,
        FillSeriesInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            source_range: "C2:C4".to_string(),
            fill_count: 6,
            direction: Some(FillDirection::Down),
            fill_type: Some(FillType::Copy),
        },
    )
    .unwrap();
    println!("Copy fill result: {}", result);

    // Step 6: Save
    let save_result = save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: id.clone(),
            file_path: "output/fill_series_example.xlsx".to_string(),
        },
    )
    .unwrap();
    println!("\nSave result: {}", save_result);
    println!("\nDone! File saved to output/fill_series_example.xlsx");
}

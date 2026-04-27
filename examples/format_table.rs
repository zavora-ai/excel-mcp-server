//! Example: Format as Table Header + Table Range
//!
//! Demonstrates using `format_as_table_header` and `format_as_table_range` to style
//! tabular data with header formatting, freeze panes, autofilter, and banded rows.

use excel_mcp_server::store::{WorkbookEntry, WorkbookStore};
use excel_mcp_server::tools::format::{format_as_table_header, format_as_table_range};
use excel_mcp_server::tools::workbook::save_workbook;
use excel_mcp_server::types::inputs::{
    FormatAsTableHeaderInput, FormatAsTableRangeInput, SaveWorkbookInput,
};
use std::time::Instant;

fn main() {
    println!("=== Format Table Example ===\n");
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

    // Step 2: Write tabular data
    {
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();

        // Headers
        ws.write(0, 0, "Employee").unwrap();
        ws.write(0, 1, "Department").unwrap();
        ws.write(0, 2, "Salary").unwrap();
        ws.write(0, 3, "Start Date").unwrap();
        ws.write(0, 4, "Rating").unwrap();

        // Data
        let employees = [
            ("Alice Johnson", "Engineering", 95000.0, "2020-03-15", 4.5),
            ("Bob Smith", "Marketing", 72000.0, "2019-07-01", 3.8),
            ("Carol White", "Sales", 68000.0, "2021-01-10", 4.2),
            ("David Brown", "Engineering", 105000.0, "2018-11-20", 4.8),
            ("Eve Davis", "Operations", 62000.0, "2022-05-01", 3.5),
            ("Frank Miller", "HR", 58000.0, "2020-09-15", 4.0),
        ];

        for (i, (name, dept, salary, date, rating)) in employees.iter().enumerate() {
            let row = (i + 1) as u32;
            ws.write(row, 0, *name).unwrap();
            ws.write(row, 1, *dept).unwrap();
            ws.write(row, 2, *salary).unwrap();
            ws.write(row, 3, *date).unwrap();
            ws.write(row, 4, *rating).unwrap();
        }
    }
    println!("Wrote employee data (6 rows)");

    // Step 3: Format as table header (applies bold, colors, freeze panes, autofilter)
    let header_result = format_as_table_header(
        &mut store,
        FormatAsTableHeaderInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            header_row: None, // defaults to row 1
            background_color: None, // defaults to #4472C4
            font_color: None, // defaults to #FFFFFF
        },
    )
    .unwrap();
    println!("\nFormat as table header result:\n{}", header_result);

    // Step 4: Format as table range (applies banded rows, borders)
    let range_result = format_as_table_range(
        &mut store,
        FormatAsTableRangeInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            range: "A1:E7".to_string(),
            style: Some("blue".to_string()),
        },
    )
    .unwrap();
    println!("\nFormat as table range result:\n{}", range_result);

    // Step 5: Save
    let save_result = save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: id.clone(),
            file_path: "output/format_table_example.xlsx".to_string(),
        },
    )
    .unwrap();
    println!("\nSave result: {}", save_result);
    println!("\nDone! File saved to output/format_table_example.xlsx");
}

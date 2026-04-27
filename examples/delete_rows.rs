//! Example: Delete Rows Where
//!
//! Demonstrates using `delete_rows_where` to remove rows matching a condition.

use excel_mcp_server::store::{WorkbookEntry, WorkbookStore};
use excel_mcp_server::tools::data::delete_rows_where;
use excel_mcp_server::tools::workbook::save_workbook;
use excel_mcp_server::types::enums::ConditionOperator;
use excel_mcp_server::types::inputs::{
    DeleteRowsWhereInput, RowCondition, SaveWorkbookInput,
};
use std::time::Instant;

fn main() {
    println!("=== Delete Rows Where Example ===\n");
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

    // Step 2: Write data with rows to delete
    {
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();

        // Headers
        ws.write(0, 0, "Product").unwrap();
        ws.write(0, 1, "Status").unwrap();
        ws.write(0, 2, "Price").unwrap();

        // Data — some rows have "Discontinued" status
        let data = [
            ("Widget A", "Active", 29.99),
            ("Widget B", "Discontinued", 19.99),
            ("Widget C", "Active", 39.99),
            ("Widget D", "Discontinued", 14.99),
            ("Widget E", "Active", 49.99),
            ("Widget F", "Discontinued", 9.99),
            ("Widget G", "Active", 59.99),
        ];

        for (i, (product, status, price)) in data.iter().enumerate() {
            let row = (i + 1) as u32;
            ws.write(row, 0, *product).unwrap();
            ws.write(row, 1, *status).unwrap();
            ws.write(row, 2, *price).unwrap();
        }
    }
    println!("Wrote 7 product rows (3 discontinued)");

    // Step 3: Delete rows where Status = "Discontinued"
    let input = DeleteRowsWhereInput {
        workbook_id: id.clone(),
        sheet_name: "Sheet1".to_string(),
        condition: RowCondition {
            column: "B".to_string(),
            operator: ConditionOperator::Equals,
            value: Some("Discontinued".to_string()),
        },
        has_header: true,
    };

    let result = delete_rows_where(&mut store, input).unwrap();
    println!("\nDelete rows result:\n{}", result);

    // Step 4: Save
    let save_result = save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: id.clone(),
            file_path: "output/delete_rows_example.xlsx".to_string(),
        },
    )
    .unwrap();
    println!("\nSave result: {}", save_result);
    println!("\nDone! File saved to output/delete_rows_example.xlsx");
}

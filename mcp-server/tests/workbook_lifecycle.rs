//! Integration tests for workbook lifecycle flows (Task 10.3).
//!
//! Validates Requirements 2.1, 3.1, 3.2, 4.1, 21.3, 21.4:
//! - create → write → save → close flow
//! - open (edit) → read → write → save flow
//! - open (read-only) → read → search → csv export flow
//! - eviction after TTL expiry

use std::time::Duration;

use excel_mcp_server::store::WorkbookStore;
use excel_mcp_server::tools;
use excel_mcp_server::types::inputs::*;

/// Parse a JSON response and return it as serde_json::Value.
fn parse(json: &str) -> serde_json::Value {
    serde_json::from_str(json).expect("response should be valid JSON")
}

/// Assert the response has status "success" and return the data field.
fn assert_success(json: &str) -> serde_json::Value {
    let v = parse(json);
    assert_eq!(
        v["status"].as_str(),
        Some("success"),
        "expected success, got: {json}"
    );
    v["data"].clone()
}

/// Returns the file path.
fn create_test_xlsx() -> String {
    let dir = std::env::temp_dir().join("excel_mcp_lifecycle_tests");
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join(format!("test_{}.xlsx", uuid::Uuid::new_v4()));

    let mut wb = zavora_xlsx::Workbook::new();
    let ws = wb.worksheet(0).unwrap();
    ws.set_name("Data").unwrap();
    ws.write(0, 0, "Name").unwrap();
    ws.write(0, 1, "Value").unwrap();
    ws.write(1, 0, "Alice").unwrap();
    ws.write(1, 1, 42.0).unwrap();
    ws.write(2, 0, "Bob").unwrap();
    ws.write(2, 1, 99.0).unwrap();
    wb.save(&path).unwrap();

    path.to_str().unwrap().to_string()
}

// ─── Test 1: create → write → save → close ───

#[test]
fn test_create_write_save_close() {
    let mut store = WorkbookStore::new();

    // 1. Create workbook
    let result = tools::workbook::create_workbook(&mut store).unwrap();
    let data = assert_success(&result);
    let workbook_id = data["workbook_id"].as_str().unwrap().to_string();
    assert!(!workbook_id.is_empty());

    // 2. Write cells
    let write_result = tools::write::write_cells(
        &mut store,
        WriteCellsInput {
            workbook_id: workbook_id.clone(),
            sheet_name: "Sheet1".into(),
            cells: vec![
                CellWrite {
                    cell: "A1".into(),
                    value: serde_json::json!("Hello"),
                },
                CellWrite {
                    cell: "B1".into(),
                    value: serde_json::json!(42),
                },
                CellWrite {
                    cell: "C1".into(),
                    value: serde_json::json!(true),
                },
            ],
        },
    )
    .unwrap();
    let write_data = assert_success(&write_result);
    assert_eq!(write_data["cells_written"].as_u64(), Some(3));

    // 3. Save to temp file
    let dir = std::env::temp_dir().join("excel_mcp_lifecycle_tests");
    std::fs::create_dir_all(&dir).unwrap();
    let save_path = dir.join(format!("create_flow_{}.xlsx", uuid::Uuid::new_v4()));
    let save_path_str = save_path.to_str().unwrap().to_string();

    let save_result = tools::workbook::save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: workbook_id.clone(),
            file_path: save_path_str.clone(),
        },
    )
    .unwrap();
    assert_success(&save_result);

    // 4. Close workbook
    let close_result = tools::workbook::close_workbook(
        &mut store,
        CloseWorkbookInput {
            workbook_id: workbook_id.clone(),
        },
    )
    .unwrap();
    assert_success(&close_result);

    // 5. Verify file exists on disk
    assert!(
        save_path.exists(),
        "saved file should exist at {}",
        save_path_str
    );

    // Cleanup
    let _ = std::fs::remove_file(&save_path);
}

// ─── Test 2: open (edit) → read → write → save ───

#[test]
fn test_open_edit_read_write_save() {
    use std::time::Instant;
    use excel_mcp_server::store::*;

    let source_path = create_test_xlsx();
    let mut store = WorkbookStore::new();

    // 1. Open in edit mode (non-lazy to allow immediate reads)
    let entry = WorkbookEntry {
        id: String::new(),
        data: zavora_xlsx::Workbook::open(std::path::Path::new(&source_path)).unwrap(),
        read_only: false,
        last_access: Instant::now(),
    };
    let workbook_id = store.insert(entry).unwrap();

    // 2. Read a cell
    let read_result = tools::read::read_cell(
        &mut store,
        ReadCellInput {
            workbook_id: workbook_id.clone(),
            sheet_name: "Data".into(),
            cell: "A1".into(),
        },
    )
    .unwrap();
    let read_data = assert_success(&read_result);
    // The cell should contain "Name"
    assert_eq!(read_data["value"].as_str(), Some("Name"));

    // 3. Write new cells
    let write_result = tools::write::write_cells(
        &mut store,
        WriteCellsInput {
            workbook_id: workbook_id.clone(),
            sheet_name: "Data".into(),
            cells: vec![CellWrite {
                cell: "A4".into(),
                value: serde_json::json!("Charlie"),
            }],
        },
    )
    .unwrap();
    assert_success(&write_result);

    // 4. Save to a new path
    let dir = std::env::temp_dir().join("excel_mcp_lifecycle_tests");
    let save_path = dir.join(format!("edit_flow_{}.xlsx", uuid::Uuid::new_v4()));
    let save_path_str = save_path.to_str().unwrap().to_string();

    let save_result = tools::workbook::save_workbook(
        &mut store,
        SaveWorkbookInput {
            workbook_id: workbook_id.clone(),
            file_path: save_path_str.clone(),
        },
    )
    .unwrap();
    assert_success(&save_result);
    assert!(save_path.exists(), "saved file should exist");

    // Cleanup
    let _ = std::fs::remove_file(&save_path);
    let _ = std::fs::remove_file(&source_path);
}


// ─── Test 3: open (read-only) → read → search → csv export ───

#[test]
fn test_open_readonly_read_search_csv() {
    let source_path = create_test_xlsx();
    let mut store = WorkbookStore::new();

    // 1. Open in read-only mode
    let open_result = tools::workbook::open_workbook(
        &mut store,
        OpenWorkbookInput {
            file_path: source_path.clone(),
            read_only: true,
        },
    )
    .unwrap();
    let open_data = assert_success(&open_result);
    let workbook_id = open_data["workbook_id"].as_str().unwrap().to_string();

    // 2. Read sheet data
    let read_result = tools::read::read_sheet(
        &mut store,
        ReadSheetInput {
            workbook_id: workbook_id.clone(),
            sheet_name: "Data".into(),
            range: None,
            continuation_token: None,
        },
    )
    .unwrap();
    let read_data = assert_success(&read_result);
    // Should have rows of data
    let rows = read_data["rows"].as_array().unwrap();
    assert!(rows.len() >= 2, "should have at least 2 data rows");

    // 3. Search for "Alice"
    let search_result = tools::read::search_cells(
        &mut store,
        SearchCellsInput {
            workbook_id: workbook_id.clone(),
            sheet_name: Some("Data".into()),
            query: "Alice".into(),
            match_mode: excel_mcp_server::types::enums::MatchMode::Exact,
        },
    )
    .unwrap();
    let search_data = assert_success(&search_result);
    let matches = search_data["matches"].as_array().unwrap();
    assert!(
        !matches.is_empty(),
        "search for 'Alice' should find at least one match"
    );
    // Verify the match is in the expected cell
    assert_eq!(matches[0]["value"].as_str(), Some("Alice"));

    // 4. CSV export
    let csv_result = tools::read::sheet_to_csv(
        &mut store,
        SheetToCsvInput {
            workbook_id: workbook_id.clone(),
            sheet_name: "Data".into(),
            delimiter: ",".into(),
        },
    )
    .unwrap();
    let csv_data = assert_success(&csv_result);
    let csv_str = csv_data["csv"].as_str().unwrap();
    assert!(
        csv_str.contains("Alice"),
        "CSV should contain 'Alice', got: {csv_str}"
    );
    assert!(
        csv_str.contains("Bob"),
        "CSV should contain 'Bob', got: {csv_str}"
    );

    // Cleanup
    let _ = std::fs::remove_file(&source_path);
}

// ─── Test 4: TTL eviction ───

#[test]
fn test_ttl_eviction() {
    // Create store with very short TTL (1ms)
    let mut store = WorkbookStore::with_config(10, Duration::from_millis(1));

    // Create a workbook
    let result = tools::workbook::create_workbook(&mut store).unwrap();
    let data = assert_success(&result);
    let workbook_id = data["workbook_id"].as_str().unwrap().to_string();

    // Sleep long enough for TTL to expire
    std::thread::sleep(Duration::from_millis(10));

    // Try to access the workbook — should be evicted
    let write_result = tools::write::write_cells(
        &mut store,
        WriteCellsInput {
            workbook_id: workbook_id.clone(),
            sheet_name: "Sheet1".into(),
            cells: vec![CellWrite {
                cell: "A1".into(),
                value: serde_json::json!("test"),
            }],
        },
    )
    .unwrap();

    let v = parse(&write_result);
    assert_eq!(
        v["status"].as_str(),
        Some("error"),
        "should get error after TTL eviction"
    );
    assert_eq!(
        v["data"]["category"].as_str(),
        Some("not_found"),
        "error category should be not_found for evicted workbook"
    );
}

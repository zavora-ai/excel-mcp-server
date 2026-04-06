//! Integration tests for structured error responses (Task 10.2).
//!
//! Validates Requirements 23.1, 23.2, 23.3, 23.4, 23.5:
//! - All error responses contain `status: "error"`, `category`, `description`, `suggestion`
//! - Not-found workbook error includes open IDs
//! - Invalid cell reference error identifies the bad reference
//! - Capacity exceeded error message

use std::time::{Duration, Instant};

use excel_mcp_server::store::*;
use excel_mcp_server::tools;
use excel_mcp_server::types::inputs::*;

fn insert_rxw(store: &mut WorkbookStore) -> String {
    let entry = WorkbookEntry {
        id: String::new(),
        data: zavora_xlsx::Workbook::new(), read_only: false,
        last_access: Instant::now(),
    };
    store.insert(entry).expect("insert rxw workbook")
}

/// Helper: parse a JSON response string and return it as serde_json::Value.
fn parse_response(json: &str) -> serde_json::Value {
    serde_json::from_str(json).expect("response should be valid JSON")
}

/// Assert that a JSON response has the standard error structure:
/// status="error", and data contains category, description, suggestion.
fn assert_error_structure(v: &serde_json::Value) {
    assert_eq!(
        v["status"].as_str(),
        Some("error"),
        "expected status 'error', got: {:?}",
        v["status"]
    );
    let data = &v["data"];
    assert!(
        data["category"].is_string(),
        "expected data.category to be a string, got: {:?}",
        data["category"]
    );
    assert!(
        data["description"].is_string(),
        "expected data.description to be a string, got: {:?}",
        data["description"]
    );
    assert!(
        data["suggestion"].is_string(),
        "expected data.suggestion to be a string, got: {:?}",
        data["suggestion"]
    );
}

// ─── Requirement 23.1, 23.3: Error response structure ───

#[test]
fn test_error_response_structure() {
    // Trigger an error by reading from a non-existent sheet.
    let mut store = WorkbookStore::new();
    let id = insert_rxw(&mut store);

    let input = ReadCellInput {
        workbook_id: id,
        sheet_name: "NonExistentSheet".into(),
        cell: "A1".into(),
    };
    let result = tools::read::read_cell(&mut store, input).unwrap();
    let v = parse_response(&result);

    assert_error_structure(&v);
    assert_eq!(v["data"]["category"].as_str(), Some("not_found"));
}

// ─── Requirement 23.3, 23.5: Not-found workbook error includes open IDs ───

#[test]
fn test_not_found_workbook_includes_open_ids() {
    let mut store = WorkbookStore::new();
    // Insert a workbook so we have a known open ID
    let known_id = insert_rxw(&mut store);

    // Try to access a non-existent workbook
    let input = WriteCellsInput {
        workbook_id: "nonexistent-id".into(),
        sheet_name: "Sheet1".into(),
        cells: vec![CellWrite {
            cell: "A1".into(),
            value: serde_json::json!("test"),
        }],
    };
    let result = tools::write::write_cells(&mut store, input).unwrap();
    let v = parse_response(&result);

    assert_error_structure(&v);
    assert_eq!(v["data"]["category"].as_str(), Some("not_found"));

    // The suggestion should mention the open workbook ID
    let suggestion = v["data"]["suggestion"].as_str().unwrap();
    assert!(
        suggestion.contains(&known_id),
        "suggestion should contain the open workbook ID '{}', got: {}",
        known_id,
        suggestion
    );
}

// ─── Requirement 23.3, 23.5: Invalid cell reference error identifies the bad reference ───

#[test]
fn test_invalid_cell_reference_error() {
    let mut store = WorkbookStore::new();
    let id = insert_rxw(&mut store);

    let bad_ref = "ZZZZZ1";
    let input = WriteCellsInput {
        workbook_id: id,
        sheet_name: "Sheet1".into(),
        cells: vec![CellWrite {
            cell: bad_ref.into(),
            value: serde_json::json!("test"),
        }],
    };
    let result = tools::write::write_cells(&mut store, input).unwrap();
    let v = parse_response(&result);

    assert_error_structure(&v);
    assert_eq!(v["data"]["category"].as_str(), Some("parse_error"));

    // The description should mention the bad reference
    let description = v["data"]["description"].as_str().unwrap();
    assert!(
        description.contains(bad_ref),
        "description should contain the bad reference '{}', got: {}",
        bad_ref,
        description
    );
}

// ─── Requirement 23.3: Capacity exceeded error message ───

#[test]
fn test_capacity_exceeded_error() {
    // Create a store with max_capacity=1
    let mut store = WorkbookStore::with_config(1, Duration::from_secs(600));

    // Fill the single slot
    let _id = insert_rxw(&mut store);

    // Creating another workbook should fail with capacity exceeded
    let result = tools::workbook::create_workbook(&mut store).unwrap();
    let v = parse_response(&result);

    assert_error_structure(&v);
    assert_eq!(v["data"]["category"].as_str(), Some("capacity_exceeded"));

    let description = v["data"]["description"].as_str().unwrap();
    assert!(
        description.to_lowercase().contains("capacity"),
        "description should mention capacity, got: {}",
        description
    );
}

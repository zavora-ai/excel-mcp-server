use super::common::{adjust_formula_col_refs, workbook_not_found};
use crate::engines::zavora;
use crate::store::WorkbookStore;
use crate::types::inputs::*;
use crate::types::responses::*;

pub fn write_cells(
    store: &mut WorkbookStore,
    input: WriteCellsInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_not_found(&input.sheet_name)),
    };
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    let mut count = 0usize;
    for cw in &input.cells {
        let (row, col) = match zavora_xlsx::utility::parse_cell_ref(&cw.cell) {
            Ok(rc) => rc,
            Err(e) => {
                return Ok(error(
                    ErrorCategory::ParseError,
                    &format!("Invalid cell reference '{}': {e}", cw.cell),
                    "Use A1 notation.",
                ))
            }
        };
        if let Err(e) = zavora::write_json_value(ws, row, col, &cw.value) {
            return Ok(error(
                ErrorCategory::IoError,
                &format!("Write error: {e}"),
                "Check value type.",
            ));
        }
        count += 1;
    }
    Ok(success(
        "Cells written",
        WriteResult {
            cells_written: count,
            range_covered: format!("{} cells", count),
        },
    ))
}

pub fn write_row(store: &mut WorkbookStore, input: WriteRowInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_not_found(&input.sheet_name)),
    };
    let (row, start_col) = zavora_xlsx::utility::parse_cell_ref(&input.start_cell)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    for (i, val) in input.values.iter().enumerate() {
        zavora::write_json_value(ws, row, start_col + i as u16, val)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
    }
    let end = zavora_xlsx::utility::to_a1(row, start_col + input.values.len() as u16 - 1);
    Ok(success(
        "Row written",
        WriteResult {
            cells_written: input.values.len(),
            range_covered: format!("{}:{}", input.start_cell, end),
        },
    ))
}

pub fn write_column(
    store: &mut WorkbookStore,
    input: WriteColumnInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_not_found(&input.sheet_name)),
    };
    let (start_row, col) = zavora_xlsx::utility::parse_cell_ref(&input.start_cell)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    for (i, val) in input.values.iter().enumerate() {
        zavora::write_json_value(ws, start_row + i as u32, col, val)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
    }
    let end = zavora_xlsx::utility::to_a1(start_row + input.values.len() as u32 - 1, col);
    Ok(success(
        "Column written",
        WriteResult {
            cells_written: input.values.len(),
            range_covered: format!("{}:{}", input.start_cell, end),
        },
    ))
}

pub fn write_grid(
    store: &mut WorkbookStore,
    input: WriteGridInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_not_found(&input.sheet_name)),
    };
    let (start_row, start_col) = zavora_xlsx::utility::parse_cell_ref(&input.start_cell)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    let mut cells_written = 0usize;
    let mut max_cols = 0usize;
    let mut failures = Vec::new();

    for (ri, row_vals) in input.rows.iter().enumerate() {
        if row_vals.len() > max_cols {
            max_cols = row_vals.len();
        }
        for (ci, val) in row_vals.iter().enumerate() {
            let r = start_row + ri as u32;
            let c = start_col + ci as u16;
            if let Err(e) = zavora::write_json_value(ws, r, c, val) {
                let cell_ref = zavora_xlsx::utility::to_a1(r, c);
                failures.push(format!("{}: {}", cell_ref, e));
            } else {
                cells_written += 1;
            }
        }
    }

    Ok(success(
        "Grid written",
        WriteGridResult {
            rows_written: input.rows.len(),
            columns_written: max_cols,
            cells_written,
            failures,
        },
    ))
}

pub fn write_row_range(
    store: &mut WorkbookStore,
    input: WriteRowRangeInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_not_found(&input.sheet_name)),
    };
    let (row, start_col) = zavora_xlsx::utility::parse_cell_ref(&input.start_cell)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    let end_col = zavora_xlsx::utility::col_from_letter(&input.end_column)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    if start_col >= end_col {
        return Ok(error(
            ErrorCategory::InvalidInput,
            &format!(
                "Start column ({}) must be less than end column ({})",
                input.start_cell, input.end_column
            ),
            "Provide a start cell whose column is before the end column.",
        ));
    }

    // Strip leading "=" if present
    let base_formula = input.formula.strip_prefix('=').unwrap_or(&input.formula);

    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    let mut cells_written = 0usize;
    for col in start_col..=end_col {
        let offset = col as i16 - start_col as i16;
        let adjusted = adjust_formula_col_refs(base_formula, offset);
        ws.write_formula(row, col, &adjusted)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
        cells_written += 1;
    }

    Ok(success(
        "Row range written",
        WriteRowRangeResult { cells_written },
    ))
}

pub fn clone_column_formulas(
    store: &mut WorkbookStore,
    input: CloneColumnFormulasInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_not_found(&input.sheet_name)),
    };
    let source_col = zavora_xlsx::utility::col_from_letter(&input.source_column)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    let mut target_cols = Vec::with_capacity(input.target_columns.len());
    for tc in &input.target_columns {
        let c = zavora_xlsx::utility::col_from_letter(tc)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
        target_cols.push(c);
    }

    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    // Convert 1-based row numbers to 0-based
    let start_row = input.start_row.saturating_sub(1);
    let end_row = input.end_row.saturating_sub(1);

    // First, collect formulas from the source column
    let mut source_formulas: Vec<(u32, String)> = Vec::new();
    for r in start_row..=end_row {
        let val = ws.read_cell(r, source_col);
        if let zavora_xlsx::CellValue::Formula { formula, .. } = val {
            source_formulas.push((r, formula));
        }
    }

    let mut formulas_cloned = 0usize;
    let mut columns_filled = 0usize;

    for &target_col in &target_cols {
        let offset = target_col as i16 - source_col as i16;
        let mut col_had_formulas = false;
        for (r, formula) in &source_formulas {
            let adjusted = adjust_formula_col_refs(formula, offset);
            ws.write_formula(*r, target_col, &adjusted)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
            formulas_cloned += 1;
            col_had_formulas = true;
        }
        if col_had_formulas {
            columns_filled += 1;
        }
    }

    Ok(success(
        "Column formulas cloned",
        CloneFormulasResult {
            formulas_cloned,
            columns_filled,
        },
    ))
}

fn find_sheet(wb: &zavora_xlsx::Workbook, name: &str) -> Option<usize> {
    wb.sheet_names().iter().position(|n| *n == name)
}

fn sheet_not_found(name: &str) -> String {
    error(
        ErrorCategory::NotFound,
        &format!("Sheet '{}' not found", name),
        "Check sheet name.",
    )
}


#[cfg(test)]
mod tests {
    use super::*;
    use crate::store::{WorkbookEntry, WorkbookStore};
    use std::time::Instant;

    /// Helper: create a store with a workbook and sheet, return (store, workbook_id).
    fn setup() -> (WorkbookStore, String) {
        let mut store = WorkbookStore::new();
        let entry = WorkbookEntry {
            id: String::new(),
            data: zavora_xlsx::Workbook::new(),
            read_only: false,
            last_access: Instant::now(),
        };
        let id = store.insert(entry).unwrap();
        (store, id)
    }

    // ── write_grid tests ──

    #[test]
    fn test_write_grid_mixed_types() {
        let (mut store, id) = setup();
        let input = WriteGridInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            start_cell: "A1".to_string(),
            rows: vec![
                vec![
                    serde_json::json!(42),
                    serde_json::json!("hello"),
                    serde_json::json!(true),
                ],
                vec![
                    serde_json::json!("=A1+1"),
                    serde_json::json!(3.14),
                    serde_json::json!(false),
                ],
            ],
        };
        let result = write_grid(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""), "Expected success: {}", result);
        assert!(result.contains("\"rows_written\":2"), "Expected 2 rows: {}", result);
        assert!(result.contains("\"columns_written\":3"), "Expected 3 cols: {}", result);
        assert!(result.contains("\"cells_written\":6"), "Expected 6 cells: {}", result);

        // Verify values were written
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();
        // A1 = 42
        match ws.read_cell(0, 0) {
            zavora_xlsx::CellValue::Number(n) => assert_eq!(n, 42.0),
            other => panic!("Expected Number(42), got {:?}", other),
        }
        // B1 = "hello"
        match ws.read_cell(0, 1) {
            zavora_xlsx::CellValue::String(s) => assert_eq!(s, "hello"),
            other => panic!("Expected String(hello), got {:?}", other),
        }
        // C1 = true
        match ws.read_cell(0, 2) {
            zavora_xlsx::CellValue::Bool(b) => assert!(b),
            other => panic!("Expected Bool(true), got {:?}", other),
        }
        // A2 = formula "=A1+1"
        match ws.read_cell(1, 0) {
            zavora_xlsx::CellValue::Formula { formula, .. } => assert_eq!(formula, "A1+1"),
            other => panic!("Expected Formula, got {:?}", other),
        }
    }

    #[test]
    fn test_write_grid_partial_failure() {
        let (mut store, id) = setup();
        // Write a grid with some valid and some null values (nulls are skipped, not failures)
        let input = WriteGridInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            start_cell: "A1".to_string(),
            rows: vec![
                vec![serde_json::json!(1), serde_json::json!(null)],
                vec![serde_json::json!("text"), serde_json::json!(2)],
            ],
        };
        let result = write_grid(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""), "Expected success: {}", result);
        // Null values are skipped by write_json_value (not failures)
        assert!(result.contains("\"rows_written\":2"), "Expected 2 rows: {}", result);
    }

    // ── write_row_range tests ──

    #[test]
    fn test_write_row_range_basic() {
        let (mut store, id) = setup();
        let input = WriteRowRangeInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            start_cell: "B10".to_string(),
            end_column: "D".to_string(),
            formula: "=B5*(1+0.05)".to_string(),
        };
        let result = write_row_range(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""), "Expected success: {}", result);
        assert!(result.contains("\"cells_written\":3"), "Expected 3 cells: {}", result);

        // Verify formulas
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();
        // B10 (row=9, col=1): B5*(1+0.05)
        match ws.read_cell(9, 1) {
            zavora_xlsx::CellValue::Formula { formula, .. } => {
                assert_eq!(formula, "B5*(1+0.05)");
            }
            other => panic!("Expected Formula at B10, got {:?}", other),
        }
        // C10 (row=9, col=2): C5*(1+0.05)
        match ws.read_cell(9, 2) {
            zavora_xlsx::CellValue::Formula { formula, .. } => {
                assert_eq!(formula, "C5*(1+0.05)");
            }
            other => panic!("Expected Formula at C10, got {:?}", other),
        }
        // D10 (row=9, col=3): D5*(1+0.05)
        match ws.read_cell(9, 3) {
            zavora_xlsx::CellValue::Formula { formula, .. } => {
                assert_eq!(formula, "D5*(1+0.05)");
            }
            other => panic!("Expected Formula at D10, got {:?}", other),
        }
    }

    #[test]
    fn test_write_row_range_absolute_refs() {
        let (mut store, id) = setup();
        let input = WriteRowRangeInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            start_cell: "B1".to_string(),
            end_column: "D".to_string(),
            formula: "$A$1*B2".to_string(), // no leading "="
        };
        let result = write_row_range(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""), "Expected success: {}", result);

        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();
        // B1: $A$1*B2
        match ws.read_cell(0, 1) {
            zavora_xlsx::CellValue::Formula { formula, .. } => {
                assert_eq!(formula, "$A$1*B2");
            }
            other => panic!("Expected Formula at B1, got {:?}", other),
        }
        // C1: $A$1*C2 (B shifted to C, $A$1 preserved)
        match ws.read_cell(0, 2) {
            zavora_xlsx::CellValue::Formula { formula, .. } => {
                assert_eq!(formula, "$A$1*C2");
            }
            other => panic!("Expected Formula at C1, got {:?}", other),
        }
        // D1: $A$1*D2
        match ws.read_cell(0, 3) {
            zavora_xlsx::CellValue::Formula { formula, .. } => {
                assert_eq!(formula, "$A$1*D2");
            }
            other => panic!("Expected Formula at D1, got {:?}", other),
        }
    }

    #[test]
    fn test_write_row_range_invalid_columns() {
        let (mut store, id) = setup();
        // Start column D >= end column B → error
        let input = WriteRowRangeInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            start_cell: "D1".to_string(),
            end_column: "B".to_string(),
            formula: "=A1".to_string(),
        };
        let result = write_row_range(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"error\""), "Expected error: {}", result);
        assert!(result.contains("must be less than"), "Expected column order error: {}", result);
    }

    // ── clone_column_formulas tests ──

    #[test]
    fn test_clone_column_no_formulas() {
        let (mut store, id) = setup();
        // Write plain values in column C (no formulas)
        {
            let entry = store.get_mut(&id).unwrap();
            let ws = entry.data.worksheet(0).unwrap();
            ws.write(0, 2, 100.0).unwrap(); // C1
            ws.write(1, 2, 200.0).unwrap(); // C2
        }
        let input = CloneColumnFormulasInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            source_column: "C".to_string(),
            target_columns: vec!["D".to_string(), "E".to_string()],
            start_row: 1,
            end_row: 2,
        };
        let result = clone_column_formulas(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""), "Expected success: {}", result);
        assert!(result.contains("\"formulas_cloned\":0"), "Expected 0 formulas: {}", result);
        assert!(result.contains("\"columns_filled\":0"), "Expected 0 columns: {}", result);
    }

    #[test]
    fn test_clone_column_skips_values() {
        let (mut store, id) = setup();
        // Write a mix of formulas and values in column C
        {
            let entry = store.get_mut(&id).unwrap();
            let ws = entry.data.worksheet(0).unwrap();
            ws.write_formula(0, 2, "A1+B1").unwrap(); // C1 = formula
            ws.write(1, 2, "plain text").unwrap();     // C2 = value (not formula)
            ws.write_formula(2, 2, "A3*2").unwrap();   // C3 = formula
        }
        let input = CloneColumnFormulasInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".to_string(),
            source_column: "C".to_string(),
            target_columns: vec!["D".to_string()],
            start_row: 1,
            end_row: 3,
        };
        let result = clone_column_formulas(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""), "Expected success: {}", result);
        // Only 2 formulas should be cloned (C1 and C3), C2 is skipped
        assert!(result.contains("\"formulas_cloned\":2"), "Expected 2 formulas: {}", result);
        assert!(result.contains("\"columns_filled\":1"), "Expected 1 column: {}", result);

        // Verify D1 has adjusted formula: B1+C1 (offset +1 from C→D)
        // Wait, source is C (col 2), target is D (col 3), offset = 1
        // C1 formula "A1+B1" → D1 formula "B1+C1"
        let entry = store.get_mut(&id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();
        match ws.read_cell(0, 3) {
            zavora_xlsx::CellValue::Formula { formula, .. } => {
                assert_eq!(formula, "B1+C1");
            }
            other => panic!("Expected Formula at D1, got {:?}", other),
        }
        // D2 should still be empty (value was skipped)
        match ws.read_cell(1, 3) {
            zavora_xlsx::CellValue::Empty => {}
            other => panic!("Expected Empty at D2, got {:?}", other),
        }
        // D3: A3*2 → B3*2
        match ws.read_cell(2, 3) {
            zavora_xlsx::CellValue::Formula { formula, .. } => {
                assert_eq!(formula, "B3*2");
            }
            other => panic!("Expected Formula at D3, got {:?}", other),
        }
    }

    // ── Property-based tests ──

    use proptest::prelude::*;

    // **Validates: Requirements 10.1, 10.3**
    //
    // Property 5: Write Grid Round-Trip
    // Write random 2D arrays (1-10 rows, 1-10 cols) of numbers/strings/booleans,
    // read back, verify equivalence and dimensions match.
    proptest! {
        #[test]
        fn prop_write_grid_round_trip(
            num_rows in 1usize..=10,
            num_cols in 1usize..=10,
            seed in 0u64..10000,
        ) {
            use std::collections::hash_map::DefaultHasher;
            use std::hash::{Hash, Hasher};

            // Generate deterministic data based on seed
            let mut rows: Vec<Vec<serde_json::Value>> = Vec::new();
            for ri in 0..num_rows {
                let mut row = Vec::new();
                for ci in 0..num_cols {
                    let mut hasher = DefaultHasher::new();
                    (seed, ri, ci).hash(&mut hasher);
                    let h = hasher.finish();
                    let val = match h % 3 {
                        0 => serde_json::json!((h % 10000) as f64 / 100.0),
                        1 => serde_json::json!(format!("str_{}_{}", ri, ci)),
                        _ => serde_json::json!(h % 2 == 0),
                    };
                    row.push(val);
                }
                rows.push(row);
            }

            let mut store = WorkbookStore::new();
            let entry = WorkbookEntry {
                id: String::new(),
                data: zavora_xlsx::Workbook::new(),
                read_only: false,
                last_access: Instant::now(),
            };
            let id = store.insert(entry).unwrap();

            let input = WriteGridInput {
                workbook_id: id.clone(),
                sheet_name: "Sheet1".to_string(),
                start_cell: "A1".to_string(),
                rows: rows.clone(),
            };
            let result = write_grid(&mut store, input).unwrap();
            let parsed: serde_json::Value = serde_json::from_str(&result).unwrap();
            let data = &parsed["data"];

            // Verify dimensions
            prop_assert_eq!(data["rows_written"].as_u64().unwrap() as usize, num_rows);
            prop_assert_eq!(data["columns_written"].as_u64().unwrap() as usize, num_cols);
            prop_assert_eq!(data["cells_written"].as_u64().unwrap() as usize, num_rows * num_cols);
            prop_assert!(data["failures"].as_array().unwrap().is_empty());

            // Read back and verify
            let entry = store.get_mut(&id).unwrap();
            let ws = entry.data.worksheet(0).unwrap();
            for (ri, row) in rows.iter().enumerate() {
                for (ci, val) in row.iter().enumerate() {
                    let cell = ws.read_cell(ri as u32, ci as u16);
                    match val {
                        serde_json::Value::Number(n) => {
                            let expected = n.as_f64().unwrap();
                            match cell {
                                zavora_xlsx::CellValue::Number(actual) => {
                                    prop_assert!((actual - expected).abs() < 1e-10,
                                        "Number mismatch at ({},{}): expected {}, got {}",
                                        ri, ci, expected, actual);
                                }
                                other => prop_assert!(false,
                                    "Expected Number at ({},{}), got {:?}", ri, ci, other),
                            }
                        }
                        serde_json::Value::String(s) => {
                            match cell {
                                zavora_xlsx::CellValue::String(actual) => {
                                    prop_assert_eq!(&actual, s,
                                        "String mismatch at ({},{})", ri, ci);
                                }
                                other => prop_assert!(false,
                                    "Expected String at ({},{}), got {:?}", ri, ci, other),
                            }
                        }
                        serde_json::Value::Bool(b) => {
                            match cell {
                                zavora_xlsx::CellValue::Bool(actual) => {
                                    prop_assert_eq!(actual, *b,
                                        "Bool mismatch at ({},{})", ri, ci);
                                }
                                other => prop_assert!(false,
                                    "Expected Bool at ({},{}), got {:?}", ri, ci, other),
                            }
                        }
                        _ => {}
                    }
                }
            }
        }
    }
}

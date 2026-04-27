//! Data operation tools: sort_range, find_replace, fill_series, delete_rows_where,
//! copy_range, transpose_range, remove_duplicates, split_column.

use super::common::workbook_not_found;
use crate::store::WorkbookStore;
use crate::types::inputs::*;
use crate::types::responses::*;

fn find_sheet(wb: &zavora_xlsx::Workbook, name: &str) -> Option<usize> {
    wb.sheet_names().iter().position(|n| *n == name)
}

fn sheet_err(name: &str) -> String {
    error(
        ErrorCategory::NotFound,
        &format!("Sheet '{}' not found", name),
        "Check sheet name.",
    )
}

/// Convert a cell value to its displayed string representation.
fn cell_display_value(val: &zavora_xlsx::CellValue) -> String {
    match val {
        zavora_xlsx::CellValue::Empty => String::new(),
        zavora_xlsx::CellValue::String(s) => s.clone(),
        zavora_xlsx::CellValue::Number(n) => {
            if *n == (*n as i64) as f64 {
                format!("{}", *n as i64)
            } else {
                format!("{}", n)
            }
        }
        zavora_xlsx::CellValue::Bool(b) => {
            if *b {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }
        }
        zavora_xlsx::CellValue::DateTime(dt) => dt.to_iso_string(),
        zavora_xlsx::CellValue::Error(e) => format!("#ERR:{e}"),
        zavora_xlsx::CellValue::Formula { cached_value, .. } => cell_display_value(cached_value),
        zavora_xlsx::CellValue::RichText(rt) => rt.plain_text(),
    }
}

/// Write a CellValue back to a worksheet cell.
fn write_cell_value(
    ws: &mut zavora_xlsx::Worksheet,
    row: u32,
    col: u16,
    val: &zavora_xlsx::CellValue,
) -> Result<(), String> {
    match val {
        zavora_xlsx::CellValue::Empty => Ok(()),
        zavora_xlsx::CellValue::String(s) => {
            ws.write(row, col, s.as_str()).map(|_| ()).map_err(|e| e.to_string())
        }
        zavora_xlsx::CellValue::Number(n) => {
            ws.write(row, col, *n).map(|_| ()).map_err(|e| e.to_string())
        }
        zavora_xlsx::CellValue::Bool(b) => {
            ws.write(row, col, *b).map(|_| ()).map_err(|e| e.to_string())
        }
        zavora_xlsx::CellValue::DateTime(dt) => {
            ws.write(row, col, dt.clone()).map(|_| ()).map_err(|e| e.to_string())
        }
        zavora_xlsx::CellValue::Formula { formula, .. } => {
            ws.write_formula(row, col, formula).map(|_| ()).map_err(|e| e.to_string())
        }
        zavora_xlsx::CellValue::RichText(rt) => {
            ws.write(row, col, rt.plain_text().as_str())
                .map(|_| ())
                .map_err(|e| e.to_string())
        }
        zavora_xlsx::CellValue::Error(_) => Ok(()),
    }
}

// ── sort_range ─────────────────────────────────────────────────────

pub fn sort_range(
    store: &mut WorkbookStore,
    input: SortRangeInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (r1, c1, r2, c2) = zavora_xlsx::utility::parse_range_ref(&input.range)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    // Validate sort keys
    let mut key_cols: Vec<(u16, bool)> = Vec::new(); // (col_index_within_range, ascending)
    for sk in &input.sort_keys {
        let col = zavora_xlsx::utility::col_from_letter(&sk.column)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
        if col < c1 || col > c2 {
            return Ok(error(
                ErrorCategory::InvalidInput,
                &format!(
                    "Sort key column '{}' is outside range {}",
                    sk.column, input.range
                ),
                "Use a column letter within the specified range.",
            ));
        }
        let ascending = match &sk.direction {
            Some(crate::types::enums::SortDirection::Descending) => false,
            _ => true,
        };
        key_cols.push((col, ascending));
    }

    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    // Read all rows
    let data_start = if input.has_header { r1 + 1 } else { r1 };
    let num_cols = (c2 - c1 + 1) as usize;

    let mut rows: Vec<Vec<zavora_xlsx::CellValue>> = Vec::new();
    for r in data_start..=r2 {
        let mut row = Vec::with_capacity(num_cols);
        for c in c1..=c2 {
            row.push(ws.read_cell(r, c));
        }
        rows.push(row);
    }

    // Stable sort
    rows.sort_by(|a, b| {
        for &(col, ascending) in &key_cols {
            let ci = (col - c1) as usize;
            let va = a.get(ci);
            let vb = b.get(ci);
            let ord = compare_cell_values(va, vb);
            let ord = if ascending { ord } else { ord.reverse() };
            if ord != std::cmp::Ordering::Equal {
                return ord;
            }
        }
        std::cmp::Ordering::Equal
    });

    // Write sorted rows back
    for (ri, row) in rows.iter().enumerate() {
        let r = data_start + ri as u32;
        for (ci, val) in row.iter().enumerate() {
            let c = c1 + ci as u16;
            let _ = write_cell_value(ws, r, c, val);
        }
    }

    Ok(success(
        "Range sorted",
        SortResult {
            rows_sorted: rows.len(),
        },
    ))
}

fn compare_cell_values(
    a: Option<&zavora_xlsx::CellValue>,
    b: Option<&zavora_xlsx::CellValue>,
) -> std::cmp::Ordering {
    let a = a.unwrap_or(&zavora_xlsx::CellValue::Empty);
    let b = b.unwrap_or(&zavora_xlsx::CellValue::Empty);

    // Empty cells sort last
    let a_empty = matches!(a, zavora_xlsx::CellValue::Empty);
    let b_empty = matches!(b, zavora_xlsx::CellValue::Empty);
    if a_empty && b_empty {
        return std::cmp::Ordering::Equal;
    }
    if a_empty {
        return std::cmp::Ordering::Greater;
    }
    if b_empty {
        return std::cmp::Ordering::Less;
    }

    let a_num = cell_as_number(a);
    let b_num = cell_as_number(b);

    match (a_num, b_num) {
        (Some(an), Some(bn)) => an.partial_cmp(&bn).unwrap_or(std::cmp::Ordering::Equal),
        _ => {
            let sa = cell_display_value(a);
            let sb = cell_display_value(b);
            sa.cmp(&sb)
        }
    }
}

fn cell_as_number(val: &zavora_xlsx::CellValue) -> Option<f64> {
    match val {
        zavora_xlsx::CellValue::Number(n) => Some(*n),
        zavora_xlsx::CellValue::Formula { cached_value, .. } => cell_as_number(cached_value),
        _ => None,
    }
}

// ── find_replace ───────────────────────────────────────────────────

pub fn find_replace(
    store: &mut WorkbookStore,
    input: FindReplaceInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };

    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    // Determine search range
    let (r1, c1, r2, c2) = if let Some(ref range_str) = input.range {
        zavora_xlsx::utility::parse_range_ref(range_str)
            .map_err(|e| anyhow::anyhow!("{e}"))?
    } else {
        match ws.used_range() {
            Some(r) => r,
            None => return Ok(success("Find/replace complete", FindReplaceResult { replacements: 0 })),
        }
    };

    let mut replacements = 0usize;

    for r in r1..=r2 {
        for c in c1..=c2 {
            let val = ws.read_cell(r, c);
            // Only replace in displayed values, not formulas
            if matches!(val, zavora_xlsx::CellValue::Formula { .. }) {
                continue;
            }
            let display = cell_display_value(&val);
            if display.is_empty() {
                continue;
            }

            let new_val = if input.match_case {
                if display.contains(&input.find) {
                    display.replace(&input.find, &input.replace)
                } else {
                    continue;
                }
            } else {
                let lower_display = display.to_lowercase();
                let lower_find = input.find.to_lowercase();
                if lower_display.contains(&lower_find) {
                    // Case-insensitive replace
                    case_insensitive_replace(&display, &input.find, &input.replace)
                } else {
                    continue;
                }
            };

            // Count occurrences
            let count = if input.match_case {
                display.matches(&input.find).count()
            } else {
                display.to_lowercase().matches(&input.find.to_lowercase()).count()
            };
            replacements += count;

            // Write the new value back
            let _ = ws.write(r, c, new_val.as_str());
        }
    }

    Ok(success(
        "Find/replace complete",
        FindReplaceResult { replacements },
    ))
}

/// Case-insensitive string replacement.
fn case_insensitive_replace(text: &str, find: &str, replace: &str) -> String {
    if find.is_empty() {
        return text.to_string();
    }
    let lower_text = text.to_lowercase();
    let lower_find = find.to_lowercase();
    let mut result = String::with_capacity(text.len());
    let mut start = 0;
    while let Some(pos) = lower_text[start..].find(&lower_find) {
        result.push_str(&text[start..start + pos]);
        result.push_str(replace);
        start += pos + find.len();
    }
    result.push_str(&text[start..]);
    result
}

// ── fill_series ────────────────────────────────────────────────────

pub fn fill_series(
    store: &mut WorkbookStore,
    input: FillSeriesInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (r1, c1, r2, c2) = zavora_xlsx::utility::parse_range_ref(&input.source_range)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    let fill_type = input.fill_type.unwrap_or(crate::types::enums::FillType::Linear);
    let direction = input.direction.unwrap_or(crate::types::enums::FillDirection::Down);

    // Read seed values
    let seeds: Vec<zavora_xlsx::CellValue> = match direction {
        crate::types::enums::FillDirection::Down => {
            (r1..=r2).map(|r| ws.read_cell(r, c1)).collect()
        }
        crate::types::enums::FillDirection::Right => {
            (c1..=c2).map(|c| ws.read_cell(r1, c)).collect()
        }
    };

    let mut cells_filled = 0u32;

    match fill_type {
        crate::types::enums::FillType::Linear => {
            // Extract numeric seeds
            let nums: Vec<f64> = seeds
                .iter()
                .filter_map(|v| cell_as_number(v))
                .collect();
            if nums.len() >= 2 {
                let step = nums[nums.len() - 1] - nums[nums.len() - 2];
                let last = nums[nums.len() - 1];
                for i in 0..input.fill_count {
                    let val = last + step * (i as f64 + 1.0);
                    match direction {
                        crate::types::enums::FillDirection::Down => {
                            let _ = ws.write(r2 + 1 + i, c1, val);
                        }
                        crate::types::enums::FillDirection::Right => {
                            let _ = ws.write(r1, c2 + 1 + i as u16, val);
                        }
                    }
                    cells_filled += 1;
                }
            } else if nums.len() == 1 {
                // Single seed: increment by 1
                let last = nums[0];
                for i in 0..input.fill_count {
                    let val = last + (i as f64 + 1.0);
                    match direction {
                        crate::types::enums::FillDirection::Down => {
                            let _ = ws.write(r2 + 1 + i, c1, val);
                        }
                        crate::types::enums::FillDirection::Right => {
                            let _ = ws.write(r1, c2 + 1 + i as u16, val);
                        }
                    }
                    cells_filled += 1;
                }
            }
        }
        crate::types::enums::FillType::Date => {
            // Try to detect date seeds and interval
            let date_strs: Vec<String> = seeds
                .iter()
                .map(|v| cell_display_value(v))
                .collect();
            // Simple date handling: try to parse as yyyy-mm-dd and detect day interval
            let dates: Vec<Option<(i32, u32, u32)>> = date_strs
                .iter()
                .map(|s| parse_simple_date(s))
                .collect();
            let valid_dates: Vec<(i32, u32, u32)> = dates.iter().filter_map(|d| *d).collect();
            if valid_dates.len() >= 2 {
                let last = valid_dates[valid_dates.len() - 1];
                let prev = valid_dates[valid_dates.len() - 2];
                let day_diff = days_between(prev, last);
                for i in 0..input.fill_count {
                    let new_date = add_days(last, day_diff * (i as i32 + 1));
                    let date_str = format!("{:04}-{:02}-{:02}", new_date.0, new_date.1, new_date.2);
                    match direction {
                        crate::types::enums::FillDirection::Down => {
                            let _ = ws.write(r2 + 1 + i, c1, date_str.as_str());
                        }
                        crate::types::enums::FillDirection::Right => {
                            let _ = ws.write(r1, c2 + 1 + i as u16, date_str.as_str());
                        }
                    }
                    cells_filled += 1;
                }
            }
        }
        crate::types::enums::FillType::Copy => {
            if !seeds.is_empty() {
                for i in 0..input.fill_count {
                    let seed_idx = i as usize % seeds.len();
                    let val = &seeds[seed_idx];
                    match direction {
                        crate::types::enums::FillDirection::Down => {
                            let _ = write_cell_value(ws, r2 + 1 + i, c1, val);
                        }
                        crate::types::enums::FillDirection::Right => {
                            let _ = write_cell_value(ws, r1, c2 + 1 + i as u16, val);
                        }
                    }
                    cells_filled += 1;
                }
            }
        }
    }

    Ok(success(
        "Series filled",
        FillSeriesResult {
            cells_filled: cells_filled as usize,
        },
    ))
}

fn parse_simple_date(s: &str) -> Option<(i32, u32, u32)> {
    let parts: Vec<&str> = s.split('-').collect();
    if parts.len() != 3 {
        return None;
    }
    let y = parts[0].parse::<i32>().ok()?;
    let m = parts[1].parse::<u32>().ok()?;
    let d = parts[2].parse::<u32>().ok()?;
    if m < 1 || m > 12 || d < 1 || d > 31 {
        return None;
    }
    Some((y, m, d))
}

fn date_to_days(y: i32, m: u32, d: u32) -> i32 {
    // Simple Julian day number approximation
    let m = m as i32;
    let d = d as i32;
    let a = (14 - m) / 12;
    let y2 = y + 4800 - a;
    let m2 = m + 12 * a - 3;
    d + (153 * m2 + 2) / 5 + 365 * y2 + y2 / 4 - y2 / 100 + y2 / 400 - 32045
}

fn days_to_date(jdn: i32) -> (i32, u32, u32) {
    let a = jdn + 32044;
    let b = (4 * a + 3) / 146097;
    let c = a - (146097 * b) / 4;
    let d = (4 * c + 3) / 1461;
    let e = c - (1461 * d) / 4;
    let m = (5 * e + 2) / 153;
    let day = e - (153 * m + 2) / 5 + 1;
    let month = m + 3 - 12 * (m / 10);
    let year = 100 * b + d - 4800 + m / 10;
    (year, month as u32, day as u32)
}

fn days_between(a: (i32, u32, u32), b: (i32, u32, u32)) -> i32 {
    date_to_days(b.0, b.1, b.2) - date_to_days(a.0, a.1, a.2)
}

fn add_days(date: (i32, u32, u32), days: i32) -> (i32, u32, u32) {
    let jdn = date_to_days(date.0, date.1, date.2) + days;
    days_to_date(jdn)
}


// ── delete_rows_where ──────────────────────────────────────────────

pub fn delete_rows_where(
    store: &mut WorkbookStore,
    input: DeleteRowsWhereInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };

    let cond_col = zavora_xlsx::utility::col_from_letter(&input.condition.column)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    let (r1, _c1, r2, _c2) = match ws.used_range() {
        Some(r) => r,
        None => return Ok(success("No rows to delete", DeleteRowsResult { rows_deleted: 0 })),
    };

    let start_row = if input.has_header { r1 + 1 } else { r1 };

    // Collect rows to delete (from bottom to top)
    let mut rows_to_delete: Vec<u32> = Vec::new();
    for r in start_row..=r2 {
        let val = ws.read_cell(r, cond_col);
        if matches_condition(&val, &input.condition) {
            rows_to_delete.push(r);
        }
    }

    // Delete from bottom to top to preserve indices
    rows_to_delete.reverse();
    for &r in &rows_to_delete {
        ws.remove_rows(r, 1).map_err(|e| anyhow::anyhow!("{e}"))?;
    }

    Ok(success(
        "Rows deleted",
        DeleteRowsResult {
            rows_deleted: rows_to_delete.len(),
        },
    ))
}

fn matches_condition(val: &zavora_xlsx::CellValue, cond: &RowCondition) -> bool {
    let display = cell_display_value(val);
    let cond_val = cond.value.as_deref().unwrap_or("");

    match cond.operator {
        crate::types::enums::ConditionOperator::Equals => display == cond_val,
        crate::types::enums::ConditionOperator::NotEquals => display != cond_val,
        crate::types::enums::ConditionOperator::Contains => display.contains(cond_val),
        crate::types::enums::ConditionOperator::GreaterThan => {
            match (display.parse::<f64>(), cond_val.parse::<f64>()) {
                (Ok(a), Ok(b)) => a > b,
                _ => display.as_str() > cond_val,
            }
        }
        crate::types::enums::ConditionOperator::LessThan => {
            match (display.parse::<f64>(), cond_val.parse::<f64>()) {
                (Ok(a), Ok(b)) => a < b,
                _ => display.as_str() < cond_val,
            }
        }
        crate::types::enums::ConditionOperator::StartsWith => display.starts_with(cond_val),
        crate::types::enums::ConditionOperator::EndsWith => display.ends_with(cond_val),
        crate::types::enums::ConditionOperator::IsEmpty => {
            matches!(val, zavora_xlsx::CellValue::Empty) || display.is_empty()
        }
    }
}

// ── copy_range ─────────────────────────────────────────────────────

pub fn copy_range(
    store: &mut WorkbookStore,
    input: CopyRangeInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };

    let src_idx = match find_sheet(&entry.data, &input.source_sheet) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.source_sheet)),
    };

    let dest_sheet_name = input
        .destination_sheet
        .as_deref()
        .unwrap_or(&input.source_sheet);
    let dest_idx = match find_sheet(&entry.data, dest_sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(dest_sheet_name)),
    };

    let (r1, c1, r2, c2) = zavora_xlsx::utility::parse_range_ref(&input.source_range)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    let (dest_r, dest_c) = zavora_xlsx::utility::parse_cell_ref(&input.destination_cell)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    // Read source values
    let src_ws = entry
        .data
        .worksheet_ref(src_idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    let mut values: Vec<Vec<zavora_xlsx::CellValue>> = Vec::new();
    for r in r1..=r2 {
        let mut row = Vec::new();
        for c in c1..=c2 {
            row.push(src_ws.read_cell(r, c));
        }
        values.push(row);
    }

    // Write to destination
    let dest_ws = entry
        .data
        .worksheet(dest_idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    for (ri, row) in values.iter().enumerate() {
        for (ci, val) in row.iter().enumerate() {
            let _ = write_cell_value(dest_ws, dest_r + ri as u32, dest_c + ci as u16, val);
        }
    }

    Ok(success_no_data(&format!(
        "Range copied from {} to {}",
        input.source_range, input.destination_cell
    )))
}

// ── transpose_range ────────────────────────────────────────────────

pub fn transpose_range(
    store: &mut WorkbookStore,
    input: TransposeRangeInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (r1, c1, r2, c2) = zavora_xlsx::utility::parse_range_ref(&input.source_range)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    let num_rows = (r2 - r1 + 1) as usize;
    let num_cols = (c2 - c1 + 1) as usize;

    // Read source values
    let mut values: Vec<Vec<zavora_xlsx::CellValue>> = Vec::new();
    for r in r1..=r2 {
        let mut row = Vec::new();
        for c in c1..=c2 {
            row.push(ws.read_cell(r, c));
        }
        values.push(row);
    }

    // Determine destination
    let (dest_r, dest_c) = if let Some(ref dest) = input.destination_cell {
        zavora_xlsx::utility::parse_cell_ref(dest)
            .map_err(|e| anyhow::anyhow!("{e}"))?
    } else {
        (r1, c1)
    };

    // If writing in-place, clear the source range first
    if input.destination_cell.is_none() {
        for r in r1..=r2 {
            for c in c1..=c2 {
                let _ = ws.write(r, c, "");
            }
        }
    }

    // Write transposed
    for (ri, row) in values.iter().enumerate() {
        for (ci, val) in row.iter().enumerate() {
            // Transpose: row becomes col, col becomes row
            let _ = write_cell_value(ws, dest_r + ci as u32, dest_c + ri as u16, val);
        }
    }

    Ok(success(
        "Range transposed",
        TransposeResult {
            original_rows: num_rows,
            original_columns: num_cols,
            transposed_rows: num_cols,
            transposed_columns: num_rows,
        },
    ))
}

// ── remove_duplicates ──────────────────────────────────────────────

pub fn remove_duplicates(
    store: &mut WorkbookStore,
    input: RemoveDuplicatesInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (r1, c1, r2, c2) = zavora_xlsx::utility::parse_range_ref(&input.range)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    // Determine which columns to compare
    let compare_cols: Vec<u16> = if input.columns.is_empty() {
        (c1..=c2).collect()
    } else {
        let mut cols = Vec::new();
        for col_str in &input.columns {
            let c = zavora_xlsx::utility::col_from_letter(col_str)
                .map_err(|e| anyhow::anyhow!("{e}"))?;
            cols.push(c);
        }
        cols
    };

    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    let data_start = if input.has_header { r1 + 1 } else { r1 };
    let num_cols = (c2 - c1 + 1) as usize;

    // Read all data rows
    let mut rows: Vec<(u32, Vec<zavora_xlsx::CellValue>)> = Vec::new();
    for r in data_start..=r2 {
        let mut row = Vec::with_capacity(num_cols);
        for c in c1..=c2 {
            row.push(ws.read_cell(r, c));
        }
        rows.push((r, row));
    }

    // Find duplicates
    let mut seen: std::collections::HashSet<String> = std::collections::HashSet::new();
    let mut rows_to_delete: Vec<u32> = Vec::new();

    for (row_num, row_data) in &rows {
        let key: String = compare_cols
            .iter()
            .map(|&c| {
                let ci = (c - c1) as usize;
                row_data
                    .get(ci)
                    .map(|v| cell_display_value(v))
                    .unwrap_or_default()
            })
            .collect::<Vec<_>>()
            .join("\x00");

        if !seen.insert(key) {
            rows_to_delete.push(*row_num);
        }
    }

    // Delete from bottom to top
    rows_to_delete.reverse();
    for &r in &rows_to_delete {
        ws.remove_rows(r, 1).map_err(|e| anyhow::anyhow!("{e}"))?;
    }

    let total_data_rows = rows.len();
    let removed = rows_to_delete.len();

    Ok(success(
        "Duplicates removed",
        RemoveDuplicatesResult {
            rows_removed: removed,
            rows_remaining: total_data_rows - removed,
        },
    ))
}

// ── split_column ───────────────────────────────────────────────────

pub fn split_column(
    store: &mut WorkbookStore,
    input: SplitColumnInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let col = zavora_xlsx::utility::col_from_letter(&input.column)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    // Convert 1-based to 0-based
    let start_row = input.start_row.saturating_sub(1);
    let end_row = input.end_row.saturating_sub(1);
    let data_start = if input.has_header { start_row + 1 } else { start_row };

    let mut max_parts = 0usize;
    let mut rows_split = 0usize;

    // First pass: determine max parts
    for r in data_start..=end_row {
        let val = ws.read_cell(r, col);
        let display = cell_display_value(&val);
        if !display.is_empty() {
            let parts: Vec<&str> = display.split(&input.delimiter).collect();
            if parts.len() > max_parts {
                max_parts = parts.len();
            }
        }
    }

    // Second pass: write split parts
    for r in data_start..=end_row {
        let val = ws.read_cell(r, col);
        let display = cell_display_value(&val);
        if !display.is_empty() {
            let parts: Vec<&str> = display.split(&input.delimiter).collect();
            if parts.len() > 1 {
                rows_split += 1;
            }
            for (i, part) in parts.iter().enumerate() {
                let target_col = col + 1 + i as u16;
                let trimmed = part.trim();
                let _ = ws.write(r, target_col, trimmed);
            }
        }
    }

    Ok(success(
        "Column split",
        SplitColumnResult {
            rows_split,
            output_columns: max_parts,
        },
    ))
}


#[cfg(test)]
mod tests {
    use super::*;
    use crate::store::{WorkbookEntry, WorkbookStore};
    use std::time::Instant;

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

    fn write_data(store: &mut WorkbookStore, id: &str, data: &[&[&str]]) {
        let entry = store.get_mut(id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();
        for (ri, row) in data.iter().enumerate() {
            for (ci, val) in row.iter().enumerate() {
                if let Ok(n) = val.parse::<f64>() {
                    let _ = ws.write(ri as u32, ci as u16, n);
                } else {
                    let _ = ws.write(ri as u32, ci as u16, *val);
                }
            }
        }
    }

    fn read_cell_str(store: &mut WorkbookStore, id: &str, row: u32, col: u16) -> String {
        let entry = store.get_mut(id).unwrap();
        let ws = entry.data.worksheet(0).unwrap();
        cell_display_value(&ws.read_cell(row, col))
    }

    // ── sort_range tests ──

    #[test]
    fn test_sort_single_key() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["3"], &["1"], &["2"]]);
        let input = SortRangeInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            range: "A1:A3".into(),
            sort_keys: vec![SortKey { column: "A".into(), direction: None }],
            has_header: false,
        };
        let result = sort_range(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""));
        assert_eq!(read_cell_str(&mut store, &id, 0, 0), "1");
        assert_eq!(read_cell_str(&mut store, &id, 1, 0), "2");
        assert_eq!(read_cell_str(&mut store, &id, 2, 0), "3");
    }

    #[test]
    fn test_sort_multi_key() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[
            &["B", "2"], &["A", "3"], &["A", "1"], &["B", "1"],
        ]);
        let input = SortRangeInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            range: "A1:B4".into(),
            sort_keys: vec![
                SortKey { column: "A".into(), direction: None },
                SortKey { column: "B".into(), direction: None },
            ],
            has_header: false,
        };
        let result = sort_range(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""));
        assert_eq!(read_cell_str(&mut store, &id, 0, 0), "A");
        assert_eq!(read_cell_str(&mut store, &id, 0, 1), "1");
        assert_eq!(read_cell_str(&mut store, &id, 1, 0), "A");
        assert_eq!(read_cell_str(&mut store, &id, 1, 1), "3");
        assert_eq!(read_cell_str(&mut store, &id, 2, 0), "B");
        assert_eq!(read_cell_str(&mut store, &id, 2, 1), "1");
    }

    #[test]
    fn test_sort_with_header() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["Name"], &["Charlie"], &["Alice"], &["Bob"]]);
        let input = SortRangeInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            range: "A1:A4".into(),
            sort_keys: vec![SortKey { column: "A".into(), direction: None }],
            has_header: true,
        };
        let result = sort_range(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""));
        assert_eq!(read_cell_str(&mut store, &id, 0, 0), "Name");
        assert_eq!(read_cell_str(&mut store, &id, 1, 0), "Alice");
        assert_eq!(read_cell_str(&mut store, &id, 2, 0), "Bob");
        assert_eq!(read_cell_str(&mut store, &id, 3, 0), "Charlie");
    }

    #[test]
    fn test_sort_key_outside_range() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["1"], &["2"]]);
        let input = SortRangeInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            range: "A1:A2".into(),
            sort_keys: vec![SortKey { column: "C".into(), direction: None }],
            has_header: false,
        };
        let result = sort_range(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"error\""));
        assert!(result.contains("outside range"));
    }

    // ── find_replace tests ──

    #[test]
    fn test_find_replace_case_insensitive() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["Hello World"], &["hello there"]]);
        let input = FindReplaceInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            find: "hello".into(),
            replace: "Hi".into(),
            range: None,
            match_case: false,
        };
        let result = find_replace(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""));
        assert_eq!(read_cell_str(&mut store, &id, 0, 0), "Hi World");
        assert_eq!(read_cell_str(&mut store, &id, 1, 0), "Hi there");
    }

    #[test]
    fn test_find_replace_no_matches() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["Hello"], &["World"]]);
        let input = FindReplaceInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            find: "xyz".into(),
            replace: "abc".into(),
            range: None,
            match_case: true,
        };
        let result = find_replace(&mut store, input).unwrap();
        assert!(result.contains("\"replacements\":0"));
    }

    #[test]
    fn test_find_replace_in_range() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["foo", "foo"], &["foo", "bar"]]);
        let input = FindReplaceInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            find: "foo".into(),
            replace: "baz".into(),
            range: Some("A1:A2".into()),
            match_case: true,
        };
        let result = find_replace(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""));
        assert_eq!(read_cell_str(&mut store, &id, 0, 0), "baz");
        assert_eq!(read_cell_str(&mut store, &id, 0, 1), "foo"); // outside range
    }

    // ── fill_series tests ──

    #[test]
    fn test_fill_series_linear_integers() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["1"], &["2"], &["3"]]);
        let input = FillSeriesInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            source_range: "A1:A3".into(),
            fill_count: 3,
            direction: None,
            fill_type: None,
        };
        let result = fill_series(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""));
        assert_eq!(read_cell_str(&mut store, &id, 3, 0), "4");
        assert_eq!(read_cell_str(&mut store, &id, 4, 0), "5");
        assert_eq!(read_cell_str(&mut store, &id, 5, 0), "6");
    }

    #[test]
    fn test_fill_series_copy() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["A"], &["B"], &["C"]]);
        let input = FillSeriesInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            source_range: "A1:A3".into(),
            fill_count: 6,
            direction: None,
            fill_type: Some(crate::types::enums::FillType::Copy),
        };
        let result = fill_series(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""));
        assert_eq!(read_cell_str(&mut store, &id, 3, 0), "A");
        assert_eq!(read_cell_str(&mut store, &id, 4, 0), "B");
        assert_eq!(read_cell_str(&mut store, &id, 5, 0), "C");
        assert_eq!(read_cell_str(&mut store, &id, 6, 0), "A");
        assert_eq!(read_cell_str(&mut store, &id, 7, 0), "B");
        assert_eq!(read_cell_str(&mut store, &id, 8, 0), "C");
    }

    #[test]
    fn test_fill_series_date() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["2024-01-01"], &["2024-01-02"], &["2024-01-03"]]);
        let input = FillSeriesInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            source_range: "A1:A3".into(),
            fill_count: 2,
            direction: None,
            fill_type: Some(crate::types::enums::FillType::Date),
        };
        let result = fill_series(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""));
        assert_eq!(read_cell_str(&mut store, &id, 3, 0), "2024-01-04");
        assert_eq!(read_cell_str(&mut store, &id, 4, 0), "2024-01-05");
    }

    // ── delete_rows_where tests ──

    #[test]
    fn test_delete_rows_each_operator() {
        // Test equals
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["A"], &["B"], &["A"]]);
        let input = DeleteRowsWhereInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            condition: RowCondition {
                column: "A".into(),
                operator: crate::types::enums::ConditionOperator::Equals,
                value: Some("A".into()),
            },
            has_header: false,
        };
        let result = delete_rows_where(&mut store, input).unwrap();
        assert!(result.contains("\"rows_deleted\":2"));

        // Test contains
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["hello world"], &["goodbye"], &["hello"]]);
        let input = DeleteRowsWhereInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            condition: RowCondition {
                column: "A".into(),
                operator: crate::types::enums::ConditionOperator::Contains,
                value: Some("hello".into()),
            },
            has_header: false,
        };
        let result = delete_rows_where(&mut store, input).unwrap();
        assert!(result.contains("\"rows_deleted\":2"));

        // Test greater_than
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["10"], &["5"], &["20"]]);
        let input = DeleteRowsWhereInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            condition: RowCondition {
                column: "A".into(),
                operator: crate::types::enums::ConditionOperator::GreaterThan,
                value: Some("9".into()),
            },
            has_header: false,
        };
        let result = delete_rows_where(&mut store, input).unwrap();
        assert!(result.contains("\"rows_deleted\":2"));

        // Test is_empty
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["A"], &[""], &["C"]]);
        let input = DeleteRowsWhereInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            condition: RowCondition {
                column: "A".into(),
                operator: crate::types::enums::ConditionOperator::IsEmpty,
                value: None,
            },
            has_header: false,
        };
        let result = delete_rows_where(&mut store, input).unwrap();
        assert!(result.contains("\"rows_deleted\":1"));
    }

    #[test]
    fn test_delete_rows_with_header() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["Name"], &["Alice"], &["Bob"], &["Alice"]]);
        let input = DeleteRowsWhereInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            condition: RowCondition {
                column: "A".into(),
                operator: crate::types::enums::ConditionOperator::Equals,
                value: Some("Alice".into()),
            },
            has_header: true,
        };
        let result = delete_rows_where(&mut store, input).unwrap();
        assert!(result.contains("\"rows_deleted\":2"));
        // Header should remain
        assert_eq!(read_cell_str(&mut store, &id, 0, 0), "Name");
    }

    // ── copy_range tests ──

    #[test]
    fn test_copy_range_same_sheet() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["1", "2"], &["3", "4"]]);
        let input = CopyRangeInput {
            workbook_id: id.clone(),
            source_sheet: "Sheet1".into(),
            source_range: "A1:B2".into(),
            destination_sheet: None,
            destination_cell: "D1".into(),
        };
        let result = copy_range(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""));
        assert_eq!(read_cell_str(&mut store, &id, 0, 3), "1");
        assert_eq!(read_cell_str(&mut store, &id, 0, 4), "2");
        assert_eq!(read_cell_str(&mut store, &id, 1, 3), "3");
        assert_eq!(read_cell_str(&mut store, &id, 1, 4), "4");
    }

    #[test]
    fn test_copy_range_cross_sheet() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["hello"]]);
        // Add a second sheet
        {
            let entry = store.get_mut(&id).unwrap();
            entry.data.add_worksheet_with_name("Sheet2").unwrap();
        }
        let input = CopyRangeInput {
            workbook_id: id.clone(),
            source_sheet: "Sheet1".into(),
            source_range: "A1:A1".into(),
            destination_sheet: Some("Sheet2".into()),
            destination_cell: "B2".into(),
        };
        let result = copy_range(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""));
        // Read from Sheet2
        let entry = store.get_mut(&id).unwrap();
        let ws2 = entry.data.worksheet(1).unwrap();
        let val = cell_display_value(&ws2.read_cell(1, 1));
        assert_eq!(val, "hello");
    }

    // ── transpose tests ──

    #[test]
    fn test_transpose_basic() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["1", "2", "3"], &["4", "5", "6"]]);
        let input = TransposeRangeInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            source_range: "A1:C2".into(),
            destination_cell: Some("E1".into()),
        };
        let result = transpose_range(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""));
        assert!(result.contains("\"original_rows\":2"));
        assert!(result.contains("\"original_columns\":3"));
        assert!(result.contains("\"transposed_rows\":3"));
        assert!(result.contains("\"transposed_columns\":2"));
        // Check transposed values at E1
        assert_eq!(read_cell_str(&mut store, &id, 0, 4), "1");
        assert_eq!(read_cell_str(&mut store, &id, 1, 4), "2");
        assert_eq!(read_cell_str(&mut store, &id, 2, 4), "3");
        assert_eq!(read_cell_str(&mut store, &id, 0, 5), "4");
        assert_eq!(read_cell_str(&mut store, &id, 1, 5), "5");
        assert_eq!(read_cell_str(&mut store, &id, 2, 5), "6");
    }

    #[test]
    fn test_transpose_in_place() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["A", "B"], &["C", "D"]]);
        let input = TransposeRangeInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            source_range: "A1:B2".into(),
            destination_cell: None,
        };
        let result = transpose_range(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""));
        assert_eq!(read_cell_str(&mut store, &id, 0, 0), "A");
        assert_eq!(read_cell_str(&mut store, &id, 1, 0), "B");
        assert_eq!(read_cell_str(&mut store, &id, 0, 1), "C");
        assert_eq!(read_cell_str(&mut store, &id, 1, 1), "D");
    }

    // ── remove_duplicates tests ──

    #[test]
    fn test_remove_duplicates_all_columns() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[
            &["A", "1"], &["B", "2"], &["A", "1"], &["C", "3"],
        ]);
        let input = RemoveDuplicatesInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            range: "A1:B4".into(),
            columns: vec![],
            has_header: false,
        };
        let result = remove_duplicates(&mut store, input).unwrap();
        assert!(result.contains("\"rows_removed\":1"));
        assert!(result.contains("\"rows_remaining\":3"));
    }

    #[test]
    fn test_remove_duplicates_specific_columns() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[
            &["A", "1"], &["A", "2"], &["B", "1"],
        ]);
        let input = RemoveDuplicatesInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            range: "A1:B3".into(),
            columns: vec!["A".into()],
            has_header: false,
        };
        let result = remove_duplicates(&mut store, input).unwrap();
        assert!(result.contains("\"rows_removed\":1"));
        assert!(result.contains("\"rows_remaining\":2"));
    }

    #[test]
    fn test_remove_duplicates_with_header() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[
            &["Name"], &["Alice"], &["Bob"], &["Alice"],
        ]);
        let input = RemoveDuplicatesInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            range: "A1:A4".into(),
            columns: vec![],
            has_header: true,
        };
        let result = remove_duplicates(&mut store, input).unwrap();
        assert!(result.contains("\"rows_removed\":1"));
        assert_eq!(read_cell_str(&mut store, &id, 0, 0), "Name");
    }

    // ── split_column tests ──

    #[test]
    fn test_split_column_comma() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["a,b,c"], &["d,e"]]);
        let input = SplitColumnInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            column: "A".into(),
            start_row: 1,
            end_row: 2,
            delimiter: ",".into(),
            has_header: false,
        };
        let result = split_column(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""));
        assert!(result.contains("\"output_columns\":3"));
        assert_eq!(read_cell_str(&mut store, &id, 0, 1), "a");
        assert_eq!(read_cell_str(&mut store, &id, 0, 2), "b");
        assert_eq!(read_cell_str(&mut store, &id, 0, 3), "c");
    }

    #[test]
    fn test_split_column_custom_delimiter() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["a|b|c"]]);
        let input = SplitColumnInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            column: "A".into(),
            start_row: 1,
            end_row: 1,
            delimiter: "|".into(),
            has_header: false,
        };
        let result = split_column(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""));
        assert_eq!(read_cell_str(&mut store, &id, 0, 1), "a");
        assert_eq!(read_cell_str(&mut store, &id, 0, 2), "b");
        assert_eq!(read_cell_str(&mut store, &id, 0, 3), "c");
    }

    #[test]
    fn test_split_column_with_header() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["Tags"], &["a,b"], &["c,d"]]);
        let input = SplitColumnInput {
            workbook_id: id.clone(),
            sheet_name: "Sheet1".into(),
            column: "A".into(),
            start_row: 1,
            end_row: 3,
            delimiter: ",".into(),
            has_header: true,
        };
        let result = split_column(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""));
        // Header row should not be split
        assert_eq!(read_cell_str(&mut store, &id, 0, 0), "Tags");
    }

    // ── copy_sheet tests (in sheets.rs but tested here for convenience) ──

    #[test]
    fn test_copy_sheet_basic() {
        let (mut store, id) = setup();
        write_data(&mut store, &id, &[&["hello", "world"]]);
        let input = CopySheetInput {
            workbook_id: id.clone(),
            source_sheet: "Sheet1".into(),
            new_sheet_name: "Copy1".into(),
        };
        let result = crate::tools::sheets::copy_sheet(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"success\""));
        // Verify data in copy
        let entry = store.get_mut(&id).unwrap();
        let copy_idx = entry.data.sheet_names().iter().position(|n| *n == "Copy1").unwrap();
        let ws = entry.data.worksheet(copy_idx).unwrap();
        match ws.read_cell(0, 0) {
            zavora_xlsx::CellValue::String(s) => assert_eq!(s, "hello"),
            other => panic!("Expected String, got {:?}", other),
        }
    }

    #[test]
    fn test_copy_sheet_not_found() {
        let (mut store, id) = setup();
        let input = CopySheetInput {
            workbook_id: id.clone(),
            source_sheet: "NonExistent".into(),
            new_sheet_name: "Copy1".into(),
        };
        let result = crate::tools::sheets::copy_sheet(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"error\""));
        assert!(result.contains("not found"));
    }

    #[test]
    fn test_copy_sheet_duplicate_name() {
        let (mut store, id) = setup();
        let input = CopySheetInput {
            workbook_id: id.clone(),
            source_sheet: "Sheet1".into(),
            new_sheet_name: "Sheet1".into(),
        };
        let result = crate::tools::sheets::copy_sheet(&mut store, input).unwrap();
        assert!(result.contains("\"status\":\"error\""));
        assert!(result.contains("already exists"));
    }

    // ── Property-based tests ──

    use proptest::prelude::*;

    // **Validates: Requirements 14.1, 14.2, 14.3**
    //
    // Property 7: Sort Correctness
    // For any data range and sort keys, rows are ordered correctly;
    // header row unchanged if has_header.
    proptest! {
        #[test]
        fn prop_sort_correctness(
            data in proptest::collection::vec(
                proptest::collection::vec(-1000.0f64..1000.0f64, 1..=3),
                2..=10
            ),
            ascending in proptest::bool::ANY,
            has_header in proptest::bool::ANY,
        ) {
            let mut store = WorkbookStore::new();
            let entry = WorkbookEntry {
                id: String::new(),
                data: zavora_xlsx::Workbook::new(),
                read_only: false,
                last_access: Instant::now(),
            };
            let id = store.insert(entry).unwrap();

            // Write data
            let num_rows = data.len();
            let num_cols = data[0].len();
            {
                let entry = store.get_mut(&id).unwrap();
                let ws = entry.data.worksheet(0).unwrap();
                for (ri, row) in data.iter().enumerate() {
                    for (ci, val) in row.iter().enumerate() {
                        let _ = ws.write(ri as u32, ci as u16, *val);
                    }
                }
            }

            let end_col = (b'A' + (num_cols - 1) as u8) as char;
            let range = format!("A1:{}{}", end_col, num_rows);
            let direction = if ascending {
                None
            } else {
                Some(crate::types::enums::SortDirection::Descending)
            };

            let header_row_val = if has_header {
                let entry = store.get_mut(&id).unwrap();
                let ws = entry.data.worksheet(0).unwrap();
                Some(ws.read_cell(0, 0))
            } else {
                None
            };

            let input = SortRangeInput {
                workbook_id: id.clone(),
                sheet_name: "Sheet1".into(),
                range,
                sort_keys: vec![SortKey {
                    column: "A".into(),
                    direction,
                }],
                has_header,
            };
            let result = sort_range(&mut store, input).unwrap();
            prop_assert!(result.contains("\"status\":\"success\""));

            // Verify ordering
            let entry = store.get_mut(&id).unwrap();
            let ws = entry.data.worksheet(0).unwrap();

            if has_header {
                // Header should be unchanged
                let hdr = ws.read_cell(0, 0);
                if let (Some(zavora_xlsx::CellValue::Number(expected)), zavora_xlsx::CellValue::Number(actual)) = (&header_row_val, &hdr) {
                    prop_assert!((actual - expected).abs() < 1e-10);
                }
            }

            let start = if has_header { 1u32 } else { 0u32 };
            let end = num_rows as u32 - 1;
            for r in start..end {
                let a = ws.read_cell(r, 0);
                let b = ws.read_cell(r + 1, 0);
                if let (zavora_xlsx::CellValue::Number(va), zavora_xlsx::CellValue::Number(vb)) = (&a, &b) {
                    if ascending {
                        prop_assert!(va <= vb, "Not ascending at row {}: {} > {}", r, va, vb);
                    } else {
                        prop_assert!(va >= vb, "Not descending at row {}: {} < {}", r, va, vb);
                    }
                }
            }
        }
    }

    // **Validates: Requirements 15.1, 15.3**
    //
    // Property 11: Find-Replace Completeness
    // After replace, no cell contains the find string; returned count equals replacements made.
    proptest! {
        #[test]
        fn prop_find_replace_completeness(
            base_strings in proptest::collection::vec("[a-z]{3,8}", 2..=5),
            find_str in "[a-z]{2,4}",
        ) {
            let mut store = WorkbookStore::new();
            let entry = WorkbookEntry {
                id: String::new(),
                data: zavora_xlsx::Workbook::new(),
                read_only: false,
                last_access: Instant::now(),
            };
            let id = store.insert(entry).unwrap();

            // Write data, some containing the find string
            {
                let entry = store.get_mut(&id).unwrap();
                let ws = entry.data.worksheet(0).unwrap();
                for (i, s) in base_strings.iter().enumerate() {
                    // Inject find_str into some cells
                    let val = if i % 2 == 0 {
                        format!("{}{}{}", s, find_str, s)
                    } else {
                        s.clone()
                    };
                    let _ = ws.write(i as u32, 0, val.as_str());
                }
            }

            let input = FindReplaceInput {
                workbook_id: id.clone(),
                sheet_name: "Sheet1".into(),
                find: find_str.clone(),
                replace: "REPLACED".into(),
                range: None,
                match_case: true,
            };
            let result = find_replace(&mut store, input).unwrap();
            prop_assert!(result.contains("\"status\":\"success\""));

            // Verify no cell contains the find string
            let entry = store.get_mut(&id).unwrap();
            let ws = entry.data.worksheet(0).unwrap();
            for i in 0..base_strings.len() {
                let val = ws.read_cell(i as u32, 0);
                let display = cell_display_value(&val);
                prop_assert!(!display.contains(&find_str),
                    "Cell at row {} still contains '{}': '{}'", i, find_str, display);
            }
        }
    }

    // **Validates: Requirements 16.2**
    //
    // Property 12: Fill Series Linear Continuation
    // For 2+ numeric seeds with constant step, filled values continue the arithmetic progression.
    proptest! {
        #[test]
        fn prop_fill_series_linear_continuation(
            start in -100.0f64..100.0f64,
            step in -10.0f64..10.0f64,
            num_seeds in 2usize..=5,
            fill_count in 1u32..=10,
        ) {
            let mut store = WorkbookStore::new();
            let entry = WorkbookEntry {
                id: String::new(),
                data: zavora_xlsx::Workbook::new(),
                read_only: false,
                last_access: Instant::now(),
            };
            let id = store.insert(entry).unwrap();

            // Write seed values
            {
                let entry = store.get_mut(&id).unwrap();
                let ws = entry.data.worksheet(0).unwrap();
                for i in 0..num_seeds {
                    let val = start + step * i as f64;
                    let _ = ws.write(i as u32, 0, val);
                }
            }

            let range = format!("A1:A{}", num_seeds);
            let input = FillSeriesInput {
                workbook_id: id.clone(),
                sheet_name: "Sheet1".into(),
                source_range: range,
                fill_count,
                direction: None,
                fill_type: Some(crate::types::enums::FillType::Linear),
            };
            let result = fill_series(&mut store, input).unwrap();
            prop_assert!(result.contains("\"status\":\"success\""));

            // Verify filled values continue the progression
            let entry = store.get_mut(&id).unwrap();
            let ws = entry.data.worksheet(0).unwrap();
            for i in 0..fill_count {
                let expected = start + step * (num_seeds as f64 + i as f64);
                let val = ws.read_cell(num_seeds as u32 + i, 0);
                if let zavora_xlsx::CellValue::Number(actual) = val {
                    prop_assert!((actual - expected).abs() < 1e-6,
                        "Fill at index {}: expected {}, got {}", i, expected, actual);
                } else {
                    prop_assert!(false, "Expected Number at fill index {}, got {:?}", i, val);
                }
            }
        }
    }

    // **Validates: Requirements 16.4**
    //
    // Property 13: Fill Series Copy Cycling
    // For any seed values and fill count, "copy" type repeats seeds cyclically.
    proptest! {
        #[test]
        fn prop_fill_series_copy_cycling(
            seeds in proptest::collection::vec("[a-z]{1,5}", 1..=5),
            fill_count in 1u32..=15,
        ) {
            let mut store = WorkbookStore::new();
            let entry = WorkbookEntry {
                id: String::new(),
                data: zavora_xlsx::Workbook::new(),
                read_only: false,
                last_access: Instant::now(),
            };
            let id = store.insert(entry).unwrap();

            {
                let entry = store.get_mut(&id).unwrap();
                let ws = entry.data.worksheet(0).unwrap();
                for (i, s) in seeds.iter().enumerate() {
                    let _ = ws.write(i as u32, 0, s.as_str());
                }
            }

            let range = format!("A1:A{}", seeds.len());
            let input = FillSeriesInput {
                workbook_id: id.clone(),
                sheet_name: "Sheet1".into(),
                source_range: range,
                fill_count,
                direction: None,
                fill_type: Some(crate::types::enums::FillType::Copy),
            };
            let result = fill_series(&mut store, input).unwrap();
            prop_assert!(result.contains("\"status\":\"success\""));

            let entry = store.get_mut(&id).unwrap();
            let ws = entry.data.worksheet(0).unwrap();
            for i in 0..fill_count {
                let expected = &seeds[i as usize % seeds.len()];
                let val = ws.read_cell(seeds.len() as u32 + i, 0);
                let display = cell_display_value(&val);
                prop_assert_eq!(&display, expected,
                    "Copy cycling mismatch at fill index {}", i);
            }
        }
    }

    // **Validates: Requirements 17.1, 17.5**
    //
    // Property 9: Delete Rows Completeness
    // After deletion, no remaining non-header row matches the condition;
    // header unchanged if has_header.
    proptest! {
        #[test]
        fn prop_delete_rows_completeness(
            values in proptest::collection::vec("[a-c]", 3..=10),
            target in "[a-c]",
            has_header in proptest::bool::ANY,
        ) {
            let mut store = WorkbookStore::new();
            let entry = WorkbookEntry {
                id: String::new(),
                data: zavora_xlsx::Workbook::new(),
                read_only: false,
                last_access: Instant::now(),
            };
            let id = store.insert(entry).unwrap();

            let mut all_values = values.clone();
            if has_header {
                all_values.insert(0, "Header".to_string());
            }

            {
                let entry = store.get_mut(&id).unwrap();
                let ws = entry.data.worksheet(0).unwrap();
                for (i, v) in all_values.iter().enumerate() {
                    let _ = ws.write(i as u32, 0, v.as_str());
                }
            }

            let input = DeleteRowsWhereInput {
                workbook_id: id.clone(),
                sheet_name: "Sheet1".into(),
                condition: RowCondition {
                    column: "A".into(),
                    operator: crate::types::enums::ConditionOperator::Equals,
                    value: Some(target.clone()),
                },
                has_header,
            };
            let result = delete_rows_where(&mut store, input).unwrap();
            prop_assert!(result.contains("\"status\":\"success\""));

            // Verify no remaining row matches
            let entry = store.get_mut(&id).unwrap();
            let ws = entry.data.worksheet_ref(0).unwrap();
            if let Some((r1, _, r2, _)) = ws.used_range() {
                let start = if has_header { r1 + 1 } else { r1 };
                for r in start..=r2 {
                    let val = ws.read_cell(r, 0);
                    let display = cell_display_value(&val);
                    prop_assert_ne!(display, target.clone(),
                        "Row {} still matches target '{}' after deletion", r, target);
                }
                // Header should be preserved
                if has_header {
                    let hdr = cell_display_value(&ws.read_cell(r1, 0));
                    prop_assert_eq!(hdr, "Header");
                }
            }
        }
    }

    // **Validates: Requirements 20.1, 20.3**
    //
    // Property 8: Transpose Involution
    // Transposing twice restores original data; transposed dimensions = swapped original dimensions.
    proptest! {
        #[test]
        fn prop_transpose_involution(
            num_rows in 1usize..=5,
            num_cols in 1usize..=5,
            seed in 0u64..10000,
        ) {
            use std::collections::hash_map::DefaultHasher;
            use std::hash::{Hash, Hasher};

            let mut store = WorkbookStore::new();
            let entry = WorkbookEntry {
                id: String::new(),
                data: zavora_xlsx::Workbook::new(),
                read_only: false,
                last_access: Instant::now(),
            };
            let id = store.insert(entry).unwrap();

            // Write original data
            let mut original: Vec<Vec<String>> = Vec::new();
            {
                let entry = store.get_mut(&id).unwrap();
                let ws = entry.data.worksheet(0).unwrap();
                for ri in 0..num_rows {
                    let mut row = Vec::new();
                    for ci in 0..num_cols {
                        let mut hasher = DefaultHasher::new();
                        (seed, ri, ci).hash(&mut hasher);
                        let _h = hasher.finish();
                        let val = format!("v{}_{}", ri, ci);
                        let _ = ws.write(ri as u32, ci as u16, val.as_str());
                        row.push(val);
                    }
                    original.push(row);
                }
            }

            let end_col = (b'A' + (num_cols - 1) as u8) as char;
            let range = format!("A1:{}{}", end_col, num_rows);

            // First transpose: write to a separate area
            let input1 = TransposeRangeInput {
                workbook_id: id.clone(),
                sheet_name: "Sheet1".into(),
                source_range: range.clone(),
                destination_cell: Some("A20".into()),
            };
            let result1 = transpose_range(&mut store, input1).unwrap();
            prop_assert!(result1.contains("\"status\":\"success\""));

            // Second transpose: transpose the transposed data back
            let t_end_col = (b'A' + (num_rows - 1) as u8) as char;
            let t_range = format!("A20:{}{}", t_end_col, 20 + num_cols - 1);
            let input2 = TransposeRangeInput {
                workbook_id: id.clone(),
                sheet_name: "Sheet1".into(),
                source_range: t_range,
                destination_cell: Some("A40".into()),
            };
            let result2 = transpose_range(&mut store, input2).unwrap();
            prop_assert!(result2.contains("\"status\":\"success\""));

            // Verify double-transposed data matches original
            let entry = store.get_mut(&id).unwrap();
            let ws = entry.data.worksheet(0).unwrap();
            for ri in 0..num_rows {
                for ci in 0..num_cols {
                    let val = cell_display_value(&ws.read_cell(40 - 1 + ri as u32, ci as u16));
                    prop_assert_eq!(&val, &original[ri][ci],
                        "Mismatch at ({},{}): expected '{}', got '{}'",
                        ri, ci, original[ri][ci], val);
                }
            }
        }
    }

    // **Validates: Requirements 21.1, 21.4**
    //
    // Property 10: Remove Duplicates Uniqueness
    // After dedup, no two non-header rows have identical values in specified columns;
    // first occurrence preserved.
    proptest! {
        #[test]
        fn prop_remove_duplicates_uniqueness(
            values in proptest::collection::vec("[a-c]", 3..=10),
        ) {
            let mut store = WorkbookStore::new();
            let entry = WorkbookEntry {
                id: String::new(),
                data: zavora_xlsx::Workbook::new(),
                read_only: false,
                last_access: Instant::now(),
            };
            let id = store.insert(entry).unwrap();

            {
                let entry = store.get_mut(&id).unwrap();
                let ws = entry.data.worksheet(0).unwrap();
                for (i, v) in values.iter().enumerate() {
                    let _ = ws.write(i as u32, 0, v.as_str());
                }
            }

            let range = format!("A1:A{}", values.len());
            let input = RemoveDuplicatesInput {
                workbook_id: id.clone(),
                sheet_name: "Sheet1".into(),
                range,
                columns: vec![],
                has_header: false,
            };
            let result = remove_duplicates(&mut store, input).unwrap();
            prop_assert!(result.contains("\"status\":\"success\""));

            // Verify uniqueness
            let entry = store.get_mut(&id).unwrap();
            let ws = entry.data.worksheet_ref(0).unwrap();
            let mut seen = std::collections::HashSet::new();
            if let Some((r1, _, r2, _)) = ws.used_range() {
                for r in r1..=r2 {
                    let val = cell_display_value(&ws.read_cell(r, 0));
                    if !val.is_empty() {
                        prop_assert!(seen.insert(val.clone()),
                            "Duplicate found after dedup: '{}'", val);
                    }
                }
            }
        }
    }

    // **Validates: Requirements 22.1**
    //
    // Property 14: Split Column Round-Trip
    // For any cell value containing delimiter, splitting and re-joining
    // with same delimiter reconstructs the original value.
    proptest! {
        #[test]
        fn prop_split_column_round_trip(
            parts in proptest::collection::vec("[a-z]{1,5}", 2..=5),
        ) {
            let delimiter = ",";
            let original = parts.join(delimiter);

            let mut store = WorkbookStore::new();
            let entry = WorkbookEntry {
                id: String::new(),
                data: zavora_xlsx::Workbook::new(),
                read_only: false,
                last_access: Instant::now(),
            };
            let id = store.insert(entry).unwrap();

            {
                let entry = store.get_mut(&id).unwrap();
                let ws = entry.data.worksheet(0).unwrap();
                let _ = ws.write(0, 0, original.as_str());
            }

            let input = SplitColumnInput {
                workbook_id: id.clone(),
                sheet_name: "Sheet1".into(),
                column: "A".into(),
                start_row: 1,
                end_row: 1,
                delimiter: delimiter.into(),
                has_header: false,
            };
            let result = split_column(&mut store, input).unwrap();
            prop_assert!(result.contains("\"status\":\"success\""));

            // Read back split parts and rejoin
            let entry = store.get_mut(&id).unwrap();
            let ws = entry.data.worksheet(0).unwrap();
            let mut read_parts: Vec<String> = Vec::new();
            for ci in 0..parts.len() {
                let val = cell_display_value(&ws.read_cell(0, 1 + ci as u16));
                if !val.is_empty() {
                    read_parts.push(val);
                }
            }
            let rejoined = read_parts.join(delimiter);
            let original_clone = original.clone();
            prop_assert_eq!(&rejoined, &original_clone,
                "Round-trip failed: original='{}', rejoined='{}'", original_clone, rejoined);
        }
    }
}

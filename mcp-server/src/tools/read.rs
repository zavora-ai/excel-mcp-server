use super::common::workbook_not_found;
use crate::engines::zavora;
use crate::store::WorkbookStore;
use crate::types::inputs::*;
use crate::types::responses::*;

pub fn read_sheet(store: &mut WorkbookStore, input: ReadSheetInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)),
    };
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    let (r1, c1, r2, c2) = if let Some(ref range_str) = input.range {
        zavora_xlsx::utility::parse_range_ref(range_str).map_err(|e| anyhow::anyhow!("{e}"))?
    } else {
        match ws.used_range() {
            Some(r) => r,
            None => return Ok(success("Sheet is empty", ReadSheetData { rows: vec![], total_rows: 0, page_rows: 0, continuation_token: None })),
        }
    };
    let offset = input.continuation_token.as_ref().and_then(|t| serde_json::from_str::<ContinuationToken>(t).ok()).map(|t| t.offset).unwrap_or(0);
    let page_size = 100u32;
    let start = r1 + offset;
    let end = (start + page_size - 1).min(r2);
    let mut rows = Vec::new();
    for r in start..=end {
        let mut row = Vec::new();
        for c in c1..=c2 { row.push(zavora::cell_to_json(&ws.read_cell(r, c))); }
        rows.push(row);
    }
    let total = r2 - r1 + 1;
    let token = if end < r2 {
        Some(serde_json::to_string(&ContinuationToken { sheet: input.sheet_name.clone(), offset: end - r1 + 1, range: input.range.clone() })?)
    } else { None };
    Ok(success("Sheet data read", ReadSheetData { rows, total_rows: total, page_rows: end - start + 1, continuation_token: token }))
}

pub fn read_cell(store: &mut WorkbookStore, input: ReadCellInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (row, col) = zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    let val = ws.read_cell(row, col);
    let formula = if let zavora_xlsx::CellValue::Formula { ref formula, .. } = val { Some(formula.clone()) } else { None };
    Ok(success("Cell read", CellData { cell: input.cell, value: zavora::cell_to_json(&val), value_type: zavora::cell_type_name(&val).to_string(), formula }))
}

pub fn search_cells(store: &mut WorkbookStore, input: SearchCellsInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let names: Vec<String> = match &input.sheet_name {
        Some(n) => vec![n.clone()],
        None => entry.data.sheet_names().iter().map(|s| s.to_string()).collect(),
    };
    let mut matches = Vec::new();
    let max = 200;
    let query_lower = input.query.to_lowercase();
    let is_exact = matches!(input.match_mode, crate::types::enums::MatchMode::Exact);
    for name in &names {
        let idx = match find_sheet(&entry.data, name) { Some(i) => i, None => continue };
        let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
        if let Some((r1, c1, r2, c2)) = ws.used_range() {
            for r in r1..=r2 {
                for c in c1..=c2 {
                    if matches.len() >= max { break; }
                    let val = ws.read_cell(r, c);
                    let s = match &val {
                        zavora_xlsx::CellValue::String(s) => s.clone(),
                        zavora_xlsx::CellValue::Number(n) => format!("{n}"),
                        zavora_xlsx::CellValue::Bool(b) => b.to_string(),
                        _ => continue,
                    };
                    let hit = if is_exact { s.to_lowercase() == query_lower } else { s.to_lowercase().contains(&query_lower) };
                    if hit {
                        matches.push(SearchMatch { sheet: name.clone(), cell: zavora_xlsx::utility::to_a1(r, c), value: zavora::cell_to_json(&val) });
                    }
                }
            }
        }
    }
    let truncated = matches.len() >= max;
    let total = matches.len();
    Ok(success("Search complete", SearchResult { matches, total_matches: total, truncated }))
}

pub fn sheet_to_csv(store: &mut WorkbookStore, input: SheetToCsvInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)),
    };
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;
    let delim = input.delimiter.as_bytes().first().copied().unwrap_or(b',');
    let mut buf = Vec::new();
    ws.to_csv(&mut buf, delim).map_err(|e| anyhow::anyhow!("{e}"))?;
    let csv = String::from_utf8(buf).unwrap_or_default();
    let rows = csv.lines().count() as u32;
    Ok(success("CSV exported", CsvExportData { csv, total_rows: rows, truncated: false }))
}

fn find_sheet(wb: &zavora_xlsx::Workbook, name: &str) -> Option<usize> {
    wb.sheet_names().iter().position(|n| *n == name)
}
fn sheet_err(name: &str) -> String {
    error(ErrorCategory::NotFound, &format!("Sheet '{}' not found", name), "Check sheet name.")
}

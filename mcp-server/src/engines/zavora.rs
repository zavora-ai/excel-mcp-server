//! Zavora-xlsx engine — single engine replacing rxw + umya + calamine.
//!
//! All tool operations go through this module.

use std::path::Path;
use zavora_xlsx::{CellValue, Workbook};

use crate::cell_ref;
use crate::types::responses::SheetSummary;

/// Create a new empty workbook.
pub fn create_workbook() -> Workbook {
    Workbook::new()
}

/// Open an existing xlsx file. `read_only` uses the fast read path.
pub fn open_workbook(path: &str, read_only: bool) -> Result<Workbook, String> {
    if read_only {
        Workbook::open_readonly(Path::new(path)).map_err(|e| e.to_string())
    } else {
        Workbook::open(Path::new(path)).map_err(|e| e.to_string())
    }
}

/// Save workbook to disk.
pub fn save_workbook(wb: &mut Workbook, path: &str) -> Result<(), String> {
    wb.save(Path::new(path)).map_err(|e| e.to_string())
}

/// Get sheet summaries for a workbook.
pub fn sheet_summaries(wb: &Workbook) -> Vec<SheetSummary> {
    let names = wb.sheet_names();
    names.iter().map(|name| SheetSummary {
        name: name.to_string(),
        dimensions: None,
        row_count: None,
        col_count: None,
    }).collect()
}

/// Convert a zavora CellValue to a serde_json::Value.
pub fn cell_to_json(val: &CellValue) -> serde_json::Value {
    match val {
        CellValue::Empty => serde_json::Value::Null,
        CellValue::String(s) => serde_json::Value::String(s.clone()),
        CellValue::Number(n) => serde_json::json!(*n),
        CellValue::Bool(b) => serde_json::Value::Bool(*b),
        CellValue::DateTime(dt) => serde_json::Value::String(dt.to_iso_string()),
        CellValue::Error(e) => serde_json::Value::String(format!("#ERR:{e}")),
        CellValue::Formula { formula, cached_value } => {
            let cached = cell_to_json(cached_value);
            serde_json::json!({ "formula": formula, "cached": cached })
        }
        CellValue::RichText(rt) => serde_json::Value::String(rt.plain_text()),
    }
}

/// Get the type name of a CellValue.
pub fn cell_type_name(val: &CellValue) -> &'static str {
    match val {
        CellValue::Empty => "empty",
        CellValue::String(_) => "string",
        CellValue::Number(_) => "number",
        CellValue::Bool(_) => "boolean",
        CellValue::DateTime(_) => "datetime",
        CellValue::Error(_) => "error",
        CellValue::Formula { .. } => "formula",
        CellValue::RichText(_) => "rich_text",
    }
}

/// Convert a JSON value to a cell write operation.
pub fn write_json_value(
    ws: &mut zavora_xlsx::Worksheet,
    row: u32,
    col: u16,
    val: &serde_json::Value,
) -> Result<(), String> {
    match val {
        serde_json::Value::Null => { /* skip empty */ Ok(()) }
        serde_json::Value::Bool(b) => ws.write(row, col, *b).map(|_| ()).map_err(|e| e.to_string()),
        serde_json::Value::Number(n) => {
            let f = n.as_f64().unwrap_or(0.0);
            ws.write(row, col, f).map(|_| ()).map_err(|e| e.to_string())
        }
        serde_json::Value::String(s) => {
            if s.starts_with('=') {
                ws.write_formula(row, col, &s[1..]).map(|_| ()).map_err(|e| e.to_string())
            } else if let Some(dt) = zavora_xlsx::ExcelDateTime::parse(s) {
                ws.write(row, col, dt).map(|_| ()).map_err(|e| e.to_string())
            } else {
                ws.write(row, col, s.as_str()).map(|_| ()).map_err(|e| e.to_string())
            }
        }
        _ => ws.write(row, col, val.to_string().as_str()).map(|_| ()).map_err(|e| e.to_string()),
    }
}

use super::common::workbook_not_found;
use crate::store::WorkbookStore;
use crate::types::inputs::*;
use crate::types::responses::*;

pub fn add_data_validation(store: &mut WorkbookStore, input: AddDataValidationInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (r1, c1, r2, c2) = zavora_xlsx::utility::parse_range_ref(&input.range).map_err(|e| anyhow::anyhow!("{e}"))?;

    let rule = match input.validation {
        crate::types::enums::ValidationRule::List { ref values } => {
            let strs: Vec<&str> = values.iter().map(|s| s.as_str()).collect();
            zavora_xlsx::ValidationRule::List(strs.iter().map(|s| s.to_string()).collect())
        }
        crate::types::enums::ValidationRule::ListRange { ref range } => {
            zavora_xlsx::ValidationRule::ListRange(range.clone())
        }
        crate::types::enums::ValidationRule::WholeNumber { min, max } => {
            zavora_xlsx::ValidationRule::WholeNumber { min, max }
        }
        crate::types::enums::ValidationRule::Decimal { min, max } => {
            zavora_xlsx::ValidationRule::Decimal { min, max }
        }
        crate::types::enums::ValidationRule::TextLength { min, max } => {
            zavora_xlsx::ValidationRule::TextLength { min, max }
        }
        crate::types::enums::ValidationRule::DateRange { ref min, ref max } => {
            zavora_xlsx::ValidationRule::DateRange { min: min.clone(), max: max.clone() }
        }
        crate::types::enums::ValidationRule::CustomFormula { ref formula } => {
            zavora_xlsx::ValidationRule::Custom(formula.clone())
        }
    };

    let mut dv = zavora_xlsx::DataValidation::new(rule);
    if let Some(ref msg) = input.input_message {
        dv.set_input_message(&msg.title, &msg.body);
    }
    if let Some(ref alert) = input.error_alert {
        let style = match alert.style {
            crate::types::enums::AlertStyle::Stop => zavora_xlsx::ErrorStyle::Stop,
            crate::types::enums::AlertStyle::Warning => zavora_xlsx::ErrorStyle::Warning,
            crate::types::enums::AlertStyle::Information => zavora_xlsx::ErrorStyle::Information,
        };
        dv.set_error_message(style, &alert.title, &alert.message);
    }

    entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?.add_data_validation(r1, c1, r2, c2, &dv)?;
    Ok(success_no_data(&format!("Data validation applied to {}", input.range)))
}

fn find_sheet(wb: &zavora_xlsx::Workbook, name: &str) -> Option<usize> { wb.sheet_names().iter().position(|n| *n == name) }
fn sheet_err(name: &str) -> String { error(ErrorCategory::NotFound, &format!("Sheet '{}' not found", name), "Check sheet name.") }

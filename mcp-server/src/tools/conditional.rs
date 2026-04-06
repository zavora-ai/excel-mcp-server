use super::common::workbook_not_found;
use crate::store::WorkbookStore;
use crate::types::inputs::*;
use crate::types::responses::*;

pub fn add_conditional_format(store: &mut WorkbookStore, input: AddConditionalFormatInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e, None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) { Some(i) => i, None => return Ok(sheet_err(&input.sheet_name)) };
    let (r1, c1, r2, c2) = zavora_xlsx::utility::parse_range_ref(&input.range).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry.data.worksheet(idx).map_err(|e| anyhow::anyhow!("{e}"))?;

    match input.rule {
        crate::types::enums::ConditionalFormatRule::CellValue { ref operator, value, value2 } => {
            let op = match operator {
                crate::types::enums::ComparisonOperator::GreaterThan => zavora_xlsx::CfOperator::GreaterThan,
                crate::types::enums::ComparisonOperator::LessThan => zavora_xlsx::CfOperator::LessThan,
                crate::types::enums::ComparisonOperator::EqualTo => zavora_xlsx::CfOperator::EqualTo,
                crate::types::enums::ComparisonOperator::NotEqualTo => zavora_xlsx::CfOperator::NotEqualTo,
                crate::types::enums::ComparisonOperator::GreaterThanOrEqual => zavora_xlsx::CfOperator::GreaterThanOrEqual,
                crate::types::enums::ComparisonOperator::LessThanOrEqual => zavora_xlsx::CfOperator::LessThanOrEqual,
                crate::types::enums::ComparisonOperator::Between => zavora_xlsx::CfOperator::Between,
            };
            let mut cf = zavora_xlsx::ConditionalFormatCell::new(op, value);
            if let Some(v2) = value2 { cf.set_value2(v2); }
            if let Some(ref style) = input.format {
                let mut fmt = zavora_xlsx::Format::new();
                if let Some(ref c) = style.font_color { fmt = fmt.font_color(c.as_str()); }
                if let Some(ref c) = style.background_color { fmt = fmt.background_color(c.as_str()); }
                if style.bold == Some(true) { fmt = fmt.bold(); }
                cf.set_format(&fmt);
            }
            ws.add_conditional_format(r1, c1, r2, c2, cf)?;
        }
        crate::types::enums::ConditionalFormatRule::ColorScale2 { ref min_color, ref max_color } => {
            ws.add_conditional_format(r1, c1, r2, c2,
                zavora_xlsx::ConditionalFormat2ColorScale::new(min_color.as_str(), max_color.as_str()))?;
        }
        crate::types::enums::ConditionalFormatRule::ColorScale3 { ref min_color, ref mid_color, ref max_color } => {
            ws.add_conditional_format(r1, c1, r2, c2,
                zavora_xlsx::ConditionalFormat3ColorScale::new(min_color.as_str(), mid_color.as_str(), max_color.as_str()))?;
        }
        crate::types::enums::ConditionalFormatRule::DataBar { ref color } => {
            ws.add_conditional_format(r1, c1, r2, c2,
                zavora_xlsx::ConditionalFormatDataBar::new(color.as_str()))?;
        }
        crate::types::enums::ConditionalFormatRule::IconSet { ref style } => {
            let is = match style {
                crate::types::enums::IconSetStyle::ThreeArrows => zavora_xlsx::IconSetType::ThreeArrows,
                crate::types::enums::IconSetStyle::ThreeTrafficLights => zavora_xlsx::IconSetType::ThreeTrafficLights,
                crate::types::enums::IconSetStyle::ThreeSymbols => zavora_xlsx::IconSetType::ThreeSymbols,
                crate::types::enums::IconSetStyle::FourArrows => zavora_xlsx::IconSetType::FourArrows,
                crate::types::enums::IconSetStyle::FiveArrows => zavora_xlsx::IconSetType::FiveArrows,
            };
            ws.add_conditional_format(r1, c1, r2, c2, zavora_xlsx::ConditionalFormatIconSet::new(is))?;
        }
    }
    Ok(success_no_data(&format!("Conditional format applied to {}", input.range)))
}

fn find_sheet(wb: &zavora_xlsx::Workbook, name: &str) -> Option<usize> { wb.sheet_names().iter().position(|n| *n == name) }
fn sheet_err(name: &str) -> String { error(ErrorCategory::NotFound, &format!("Sheet '{}' not found", name), "Check sheet name.") }

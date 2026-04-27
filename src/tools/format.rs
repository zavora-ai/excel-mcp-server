use super::common::{parse_multi_range, resolve_semantic_format, workbook_not_found};
use crate::store::WorkbookStore;
use crate::types::inputs::*;
use crate::types::responses::*;

pub fn set_cell_format(
    store: &mut WorkbookStore,
    input: SetCellFormatInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    // Build format first (before range parsing)
    let mut fmt = zavora_xlsx::Format::new();
    if input.bold == Some(true) {
        fmt = fmt.bold();
    }
    if input.italic == Some(true) {
        fmt = fmt.italic();
    }
    if input.underline == Some(true) {
        fmt = fmt.underline(zavora_xlsx::Underline::Single);
    }
    if let Some(size) = input.font_size {
        fmt = fmt.font_size(size);
    }
    if let Some(ref c) = input.font_color {
        fmt = fmt.font_color(c.as_str());
    }
    if let Some(ref c) = input.background_color {
        fmt = fmt.background_color(c.as_str());
    }
    if let Some(ref nf) = input.number_format {
        let resolved = resolve_semantic_format(nf);
        fmt = fmt.num_format(resolved);
    }
    if let Some(ref ha) = input.horizontal_alignment {
        fmt = fmt.align(match ha {
            crate::types::enums::HorizontalAlignment::Left => zavora_xlsx::Align::Left,
            crate::types::enums::HorizontalAlignment::Center => zavora_xlsx::Align::Center,
            crate::types::enums::HorizontalAlignment::Right => zavora_xlsx::Align::Right,
            crate::types::enums::HorizontalAlignment::Fill => zavora_xlsx::Align::Left,
            crate::types::enums::HorizontalAlignment::Justify => zavora_xlsx::Align::Left,
        });
    }
    if let Some(ref va) = input.vertical_alignment {
        fmt = fmt.align(match va {
            crate::types::enums::VerticalAlignment::Top => zavora_xlsx::Align::Top,
            crate::types::enums::VerticalAlignment::Center => zavora_xlsx::Align::VerticalCenter,
            crate::types::enums::VerticalAlignment::Bottom => zavora_xlsx::Align::Bottom,
            crate::types::enums::VerticalAlignment::Justify => zavora_xlsx::Align::Bottom,
        });
    }
    if let Some(ref bs) = input.border_style {
        let style = match bs {
            crate::types::enums::BorderStyle::Thin => zavora_xlsx::BorderStyle::Thin,
            crate::types::enums::BorderStyle::Medium => zavora_xlsx::BorderStyle::Medium,
            crate::types::enums::BorderStyle::Thick => zavora_xlsx::BorderStyle::Thick,
            crate::types::enums::BorderStyle::Dashed => zavora_xlsx::BorderStyle::Dashed,
            crate::types::enums::BorderStyle::Dotted => zavora_xlsx::BorderStyle::Dotted,
            crate::types::enums::BorderStyle::Double => zavora_xlsx::BorderStyle::Double,
            crate::types::enums::BorderStyle::None => zavora_xlsx::BorderStyle::None,
        };
        fmt = fmt.border(style);
    }
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    // Parse all ranges upfront for atomic validation (Requirement 1.6):
    // if any segment is invalid, no formatting is applied to any range.
    let ranges = parse_multi_range(&input.range)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    for (r1, c1, r2, c2) in &ranges {
        ws.set_range_format(*r1, *c1, *r2, *c2, &fmt)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
    }
    Ok(success_no_data(&format!(
        "Format applied to {}",
        input.range
    )))
}

pub fn merge_cells(
    store: &mut WorkbookStore,
    input: MergeCellsInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    // Parse all ranges upfront for atomic validation (Requirement 1.6):
    // if any segment is invalid, no merging is applied to any range.
    let ranges =
        parse_multi_range(&input.range).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    for (r1, c1, r2, c2) in &ranges {
        ws.merge_range(*r1, *c1, *r2, *c2, "", &zavora_xlsx::Format::new())
            .map_err(|e| anyhow::anyhow!("{e}"))?;
    }
    Ok(success_no_data(&format!("Cells merged: {}", input.range)))
}

pub fn batch_format(
    store: &mut WorkbookStore,
    input: BatchFormatInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };

    let mut operations_applied: usize = 0;
    let mut failures: Vec<BatchFormatFailure> = Vec::new();

    for (op_index, op) in input.operations.iter().enumerate() {
        // Build a Format object from the operation (same logic as set_cell_format)
        let mut fmt = zavora_xlsx::Format::new();
        if op.bold == Some(true) {
            fmt = fmt.bold();
        }
        if op.italic == Some(true) {
            fmt = fmt.italic();
        }
        if op.underline == Some(true) {
            fmt = fmt.underline(zavora_xlsx::Underline::Single);
        }
        if let Some(size) = op.font_size {
            fmt = fmt.font_size(size);
        }
        if let Some(ref c) = op.font_color {
            fmt = fmt.font_color(c.as_str());
        }
        if let Some(ref c) = op.background_color {
            fmt = fmt.background_color(c.as_str());
        }
        if let Some(ref nf) = op.number_format {
            let resolved = resolve_semantic_format(nf);
            fmt = fmt.num_format(resolved);
        }
        if let Some(ref ha) = op.horizontal_alignment {
            fmt = fmt.align(match ha {
                crate::types::enums::HorizontalAlignment::Left => zavora_xlsx::Align::Left,
                crate::types::enums::HorizontalAlignment::Center => zavora_xlsx::Align::Center,
                crate::types::enums::HorizontalAlignment::Right => zavora_xlsx::Align::Right,
                crate::types::enums::HorizontalAlignment::Fill => zavora_xlsx::Align::Left,
                crate::types::enums::HorizontalAlignment::Justify => zavora_xlsx::Align::Left,
            });
        }
        if let Some(ref va) = op.vertical_alignment {
            fmt = fmt.align(match va {
                crate::types::enums::VerticalAlignment::Top => zavora_xlsx::Align::Top,
                crate::types::enums::VerticalAlignment::Center => {
                    zavora_xlsx::Align::VerticalCenter
                }
                crate::types::enums::VerticalAlignment::Bottom => zavora_xlsx::Align::Bottom,
                crate::types::enums::VerticalAlignment::Justify => zavora_xlsx::Align::Bottom,
            });
        }
        if let Some(ref bs) = op.border_style {
            let style = match bs {
                crate::types::enums::BorderStyle::Thin => zavora_xlsx::BorderStyle::Thin,
                crate::types::enums::BorderStyle::Medium => zavora_xlsx::BorderStyle::Medium,
                crate::types::enums::BorderStyle::Thick => zavora_xlsx::BorderStyle::Thick,
                crate::types::enums::BorderStyle::Dashed => zavora_xlsx::BorderStyle::Dashed,
                crate::types::enums::BorderStyle::Dotted => zavora_xlsx::BorderStyle::Dotted,
                crate::types::enums::BorderStyle::Double => zavora_xlsx::BorderStyle::Double,
                crate::types::enums::BorderStyle::None => zavora_xlsx::BorderStyle::None,
            };
            fmt = fmt.border(style);
        }

        // Parse comma-separated ranges for this operation
        let ranges = match parse_multi_range(&op.range) {
            Ok(r) => r,
            Err(e) => {
                failures.push(BatchFormatFailure {
                    operation_index: op_index,
                    range: op.range.clone(),
                    error: e,
                });
                continue;
            }
        };

        // Apply format to each range segment
        let mut op_failed = false;
        let ws = entry
            .data
            .worksheet(idx)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
        for (r1, c1, r2, c2) in &ranges {
            if let Err(e) = ws.set_range_format(*r1, *c1, *r2, *c2, &fmt) {
                failures.push(BatchFormatFailure {
                    operation_index: op_index,
                    range: op.range.clone(),
                    error: e.to_string(),
                });
                op_failed = true;
                break;
            }
        }
        if !op_failed {
            operations_applied += 1;
        }
    }

    let result = BatchFormatResult {
        operations_applied,
        failures,
    };
    Ok(success("Batch format complete", result))
}

pub fn apply_theme(
    store: &mut WorkbookStore,
    input: ApplyThemeInput,
) -> Result<String, anyhow::Error> {
    // Validate theme name (Requirement 3.8)
    let valid_themes = ["financial_professional", "corporate", "minimal"];
    if !valid_themes.contains(&input.theme.as_str()) {
        return Ok(error(
            ErrorCategory::InvalidInput,
            &format!("Unknown theme '{}'", input.theme),
            &format!(
                "Valid themes: {}",
                valid_themes.join(", ")
            ),
        ));
    }

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

    // Determine the used range to know column extent
    let (used_r1, used_c1, used_r2, used_c2) = match ws.used_range() {
        Some(r) => r,
        None => {
            // No data — nothing to theme, but not an error
            return Ok(success_no_data(&format!(
                "Theme '{}' applied (sheet is empty, no styling needed)",
                input.theme
            )));
        }
    };

    // Convert 1-based header/total rows to 0-based row indices
    let header_rows_0: Vec<u32> = input.header_rows.iter().map(|r| r.saturating_sub(1)).collect();
    let total_rows_0: Vec<u32> = input.total_rows.iter().map(|r| r.saturating_sub(1)).collect();

    match input.theme.as_str() {
        "financial_professional" => {
            apply_financial_professional(ws, used_c1, used_c2, used_r1, used_r2, &header_rows_0, &total_rows_0)?;
        }
        "corporate" => {
            apply_corporate(ws, used_c1, used_c2, used_r1, used_r2, &header_rows_0, &total_rows_0)?;
        }
        "minimal" => {
            apply_minimal(ws, used_c1, used_c2, used_r1, used_r2, &header_rows_0, &total_rows_0)?;
        }
        _ => unreachable!(), // Already validated above
    }

    // Auto-detect currency columns if opted in (Requirement 3.7)
    if input.auto_detect_formats {
        auto_detect_currency_columns(ws, used_r1, used_c1, used_r2, used_c2, &header_rows_0, &total_rows_0)?;
    }

    // Autofit columns (all themes)
    ws.autofit().map_err(|e| anyhow::anyhow!("{e}"))?;

    Ok(success_no_data(&format!(
        "Theme '{}' applied to sheet",
        input.theme
    )))
}

/// Apply the "financial_professional" theme:
/// - Headers: Bold, white font (#FFFFFF), dark blue bg (#1F3864), center align
/// - Totals: Bold, top border (medium)
/// - Data rows: Alternating light blue (#D6E4F0) / white
fn apply_financial_professional(
    ws: &mut zavora_xlsx::Worksheet,
    c1: u16,
    c2: u16,
    used_r1: u32,
    used_r2: u32,
    header_rows: &[u32],
    total_rows: &[u32],
) -> Result<(), anyhow::Error> {
    // Apply header styling
    let header_fmt = zavora_xlsx::Format::new()
        .bold()
        .font_color("#FFFFFF")
        .background_color("#1F3864")
        .align(zavora_xlsx::Align::Center);
    for &hr in header_rows {
        ws.set_range_format(hr, c1, hr, c2, &header_fmt)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
    }

    // Apply total styling
    let total_fmt = zavora_xlsx::Format::new()
        .bold()
        .border_top(zavora_xlsx::BorderStyle::Medium);
    for &tr in total_rows {
        ws.set_range_format(tr, c1, tr, c2, &total_fmt)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
    }

    // Apply alternating row shading to data rows (rows not in header or total sets)
    let light_blue_fmt = zavora_xlsx::Format::new().background_color("#D6E4F0");
    let white_fmt = zavora_xlsx::Format::new().background_color("#FFFFFF");

    let mut shade_index: usize = 0;
    for r in used_r1..=used_r2 {
        if header_rows.contains(&r) || total_rows.contains(&r) {
            continue;
        }
        let fmt = if shade_index % 2 == 0 {
            &light_blue_fmt
        } else {
            &white_fmt
        };
        ws.set_range_format(r, c1, r, c2, fmt)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
        shade_index += 1;
    }

    Ok(())
}

/// Apply the "corporate" theme:
/// - Headers: Bold, dark font (#333333), light gray bg (#E0E0E0), thin bottom border
/// - Totals: Bold, thin top border
/// - Data rows: No alternating colors, subtle thin borders on all cells
fn apply_corporate(
    ws: &mut zavora_xlsx::Worksheet,
    c1: u16,
    c2: u16,
    used_r1: u32,
    used_r2: u32,
    header_rows: &[u32],
    total_rows: &[u32],
) -> Result<(), anyhow::Error> {
    // Apply header styling
    let header_fmt = zavora_xlsx::Format::new()
        .bold()
        .font_color("#333333")
        .background_color("#E0E0E0")
        .border_bottom(zavora_xlsx::BorderStyle::Thin);
    for &hr in header_rows {
        ws.set_range_format(hr, c1, hr, c2, &header_fmt)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
    }

    // Apply total styling
    let total_fmt = zavora_xlsx::Format::new()
        .bold()
        .border_top(zavora_xlsx::BorderStyle::Thin);
    for &tr in total_rows {
        ws.set_range_format(tr, c1, tr, c2, &total_fmt)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
    }

    // Apply subtle thin borders to all data rows (not headers/totals)
    let data_fmt = zavora_xlsx::Format::new()
        .border(zavora_xlsx::BorderStyle::Thin);
    for r in used_r1..=used_r2 {
        if header_rows.contains(&r) || total_rows.contains(&r) {
            continue;
        }
        ws.set_range_format(r, c1, r, c2, &data_fmt)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
    }

    Ok(())
}

/// Apply the "minimal" theme:
/// - Headers: Bold, thin bottom border
/// - Totals: Bold, thin bottom border
/// - Data rows: No background colors, no borders
fn apply_minimal(
    ws: &mut zavora_xlsx::Worksheet,
    c1: u16,
    c2: u16,
    _used_r1: u32,
    _used_r2: u32,
    header_rows: &[u32],
    total_rows: &[u32],
) -> Result<(), anyhow::Error> {
    // Apply header styling
    let header_fmt = zavora_xlsx::Format::new()
        .bold()
        .border_bottom(zavora_xlsx::BorderStyle::Thin);
    for &hr in header_rows {
        ws.set_range_format(hr, c1, hr, c2, &header_fmt)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
    }

    // Apply total styling
    let total_fmt = zavora_xlsx::Format::new()
        .bold()
        .border_bottom(zavora_xlsx::BorderStyle::Thin);
    for &tr in total_rows {
        ws.set_range_format(tr, c1, tr, c2, &total_fmt)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
    }

    // Data rows: no formatting applied (no colors, no borders)

    Ok(())
}

/// Auto-detect columns where most values are numeric and apply currency format.
/// Scans data rows (excluding header and total rows) and applies "$#,##0.00"
/// to columns where >50% of non-empty cells are numeric.
fn auto_detect_currency_columns(
    ws: &mut zavora_xlsx::Worksheet,
    used_r1: u32,
    used_c1: u16,
    used_r2: u32,
    used_c2: u16,
    header_rows: &[u32],
    total_rows: &[u32],
) -> Result<(), anyhow::Error> {
    let currency_fmt = zavora_xlsx::Format::new().num_format("$#,##0.00");

    for col in used_c1..=used_c2 {
        let mut numeric_count: usize = 0;
        let mut non_empty_count: usize = 0;

        for row in used_r1..=used_r2 {
            if header_rows.contains(&row) || total_rows.contains(&row) {
                continue;
            }
            let val = ws.read_cell(row, col);
            match val {
                zavora_xlsx::CellValue::Number(_) | zavora_xlsx::CellValue::DateTime(_) => {
                    numeric_count += 1;
                    non_empty_count += 1;
                }
                zavora_xlsx::CellValue::Formula { .. } => {
                    // Count formulas as numeric (they typically produce numbers)
                    numeric_count += 1;
                    non_empty_count += 1;
                }
                zavora_xlsx::CellValue::Empty => {
                    // Skip empty cells
                }
                _ => {
                    non_empty_count += 1;
                }
            }
        }

        // Apply currency format if >50% of non-empty cells are numeric
        if non_empty_count > 0 && numeric_count * 2 > non_empty_count {
            for row in used_r1..=used_r2 {
                if header_rows.contains(&row) || total_rows.contains(&row) {
                    continue;
                }
                ws.set_range_format(row, col, row, col, &currency_fmt)
                    .map_err(|e| anyhow::anyhow!("{e}"))?;
            }
        }
    }

    Ok(())
}

pub fn copy_format(
    store: &mut WorkbookStore,
    input: CopyFormatInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };

    // Parse source range
    let (sr1, sc1, sr2, sc2) = zavora_xlsx::utility::parse_range_ref(&input.source_range)
        .map_err(|e| anyhow::anyhow!("Invalid source range '{}': {}", input.source_range, e))?;

    let src_rows = (sr2 - sr1 + 1) as usize;
    let src_cols = (sc2 - sc1 + 1) as usize;

    // Parse all target ranges upfront for validation
    let mut target_ranges = Vec::with_capacity(input.target_ranges.len());
    for tr in &input.target_ranges {
        let parsed = zavora_xlsx::utility::parse_range_ref(tr)
            .map_err(|e| anyhow::anyhow!("Invalid target range '{}': {}", tr, e))?;
        target_ranges.push(parsed);
    }

    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    // Read formatting from each cell in the source range into a 2D grid
    let mut source_formats: Vec<Vec<Option<zavora_xlsx::Format>>> =
        Vec::with_capacity(src_rows);
    let mut has_any_format = false;

    for r in sr1..=sr2 {
        let mut row_formats = Vec::with_capacity(src_cols);
        for c in sc1..=sc2 {
            let fmt = ws.cell_format(r, c);
            if fmt.is_some() {
                has_any_format = true;
            }
            row_formats.push(fmt);
        }
        source_formats.push(row_formats);
    }

    // If source has no formatting, return success with a note (Requirement 4.5)
    if !has_any_format {
        let result = CopyFormatResult {
            targets_formatted: 0,
            note: Some("Source range has no formatting to copy".to_string()),
        };
        return Ok(success("Copy format complete", result));
    }

    // For each target range, tile the source formatting to fill it
    let mut targets_formatted: usize = 0;

    for (tr1, tc1, tr2, tc2) in &target_ranges {
        let tgt_rows = (*tr2 - *tr1 + 1) as usize;
        let tgt_cols = (*tc2 - *tc1 + 1) as usize;

        for tr_offset in 0..tgt_rows {
            for tc_offset in 0..tgt_cols {
                // Tile: map target position to source position using modulo
                let src_r_idx = tr_offset % src_rows;
                let src_c_idx = tc_offset % src_cols;

                if let Some(ref src_fmt) = source_formats[src_r_idx][src_c_idx] {
                    // Reconstruct a Format from the source format's properties
                    let fmt = rebuild_format(src_fmt);
                    let target_row = *tr1 + tr_offset as u32;
                    let target_col = *tc1 + tc_offset as u16;
                    ws.set_range_format(target_row, target_col, target_row, target_col, &fmt)
                        .map_err(|e| anyhow::anyhow!("{e}"))?;
                }
            }
        }
        targets_formatted += 1;
    }

    let result = CopyFormatResult {
        targets_formatted,
        note: None,
    };
    Ok(success("Copy format complete", result))
}

/// Reconstruct a new `Format` from the properties read from an existing `Format`.
/// This copies: bold, italic, underline, font size, font color, background color,
/// number format, horizontal/vertical alignment, and border styles.
fn rebuild_format(src: &zavora_xlsx::Format) -> zavora_xlsx::Format {
    let mut fmt = zavora_xlsx::Format::new();

    // Bold
    if src.is_bold() {
        fmt = fmt.bold();
    }

    // Italic
    if src.is_italic() {
        fmt = fmt.italic();
    }

    // Underline
    let underline = src.get_underline();
    if underline != zavora_xlsx::Underline::None {
        fmt = fmt.underline(underline);
    }

    // Font size (default is 11.0 in Excel; only set if non-default)
    let font_size = src.get_font_size();
    if font_size != 0.0 {
        fmt = fmt.font_size(font_size);
    }

    // Font name
    let font_name = src.get_font_name();
    if !font_name.is_empty() {
        fmt = fmt.font_name(font_name);
    }

    // Font color
    if let Some(rgb) = src.get_font_color() {
        let hex = format!("#{:02X}{:02X}{:02X}", rgb[0], rgb[1], rgb[2]);
        fmt = fmt.font_color(hex.as_str());
    }

    // Background color
    if let Some(rgb) = src.get_bg_color() {
        let hex = format!("#{:02X}{:02X}{:02X}", rgb[0], rgb[1], rgb[2]);
        fmt = fmt.background_color(hex.as_str());
    }

    // Number format
    let num_format = src.get_num_format();
    if !num_format.is_empty() {
        fmt = fmt.num_format(num_format);
    }

    // Horizontal alignment (0 = none/default)
    match src.get_h_align() {
        1 => fmt = fmt.align(zavora_xlsx::Align::Left),
        2 => fmt = fmt.align(zavora_xlsx::Align::Center),
        3 => fmt = fmt.align(zavora_xlsx::Align::Right),
        4 => fmt = fmt.align(zavora_xlsx::Align::Fill),
        5 => fmt = fmt.align(zavora_xlsx::Align::Justify),
        _ => {} // 0 = no alignment set
    }

    // Vertical alignment (0 = Top/default)
    match src.get_v_align() {
        1 => fmt = fmt.align(zavora_xlsx::Align::VerticalCenter),
        2 => fmt = fmt.align(zavora_xlsx::Align::Bottom),
        _ => {} // 0 = Top (default)
    }

    // Border styles (individual sides)
    let border_top = src.get_border_top();
    if border_top != zavora_xlsx::BorderStyle::None {
        fmt = fmt.border_top(border_top);
    }
    let border_bottom = src.get_border_bottom();
    if border_bottom != zavora_xlsx::BorderStyle::None {
        fmt = fmt.border_bottom(border_bottom);
    }
    let border_left = src.get_border_left();
    if border_left != zavora_xlsx::BorderStyle::None {
        fmt = fmt.border_left(border_left);
    }
    let border_right = src.get_border_right();
    if border_right != zavora_xlsx::BorderStyle::None {
        fmt = fmt.border_right(border_right);
    }

    fmt
}

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


pub fn apply_style(
    store: &mut WorkbookStore,
    input: ApplyStyleInput,
) -> Result<String, anyhow::Error> {
    let valid_presets = [
        "header",
        "title",
        "currency",
        "percentage",
        "date",
        "number",
        "text",
        "accounting",
        "total",
    ];
    if !valid_presets.contains(&input.style.as_str()) {
        return Ok(error(
            ErrorCategory::InvalidInput,
            &format!("Unknown style preset '{}'", input.style),
            &format!("Valid presets: {}", valid_presets.join(", ")),
        ));
    }

    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };

    // Build the format for the given style preset
    let fmt = match input.style.as_str() {
        "header" => zavora_xlsx::Format::new()
            .bold()
            .font_color("#FFFFFF")
            .background_color("#4472C4")
            .align(zavora_xlsx::Align::Center),
        "title" => zavora_xlsx::Format::new()
            .bold()
            .font_size(14.0)
            .font_color("#1F3864"),
        "currency" => zavora_xlsx::Format::new().num_format("$#,##0.00"),
        "percentage" => zavora_xlsx::Format::new().num_format("0.0%"),
        "date" => zavora_xlsx::Format::new().num_format("yyyy-mm-dd"),
        "number" => zavora_xlsx::Format::new().num_format("#,##0"),
        "text" => zavora_xlsx::Format::new().num_format("@"),
        "accounting" => zavora_xlsx::Format::new()
            .num_format(r#"_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)"#),
        "total" => zavora_xlsx::Format::new()
            .bold()
            .border_top(zavora_xlsx::BorderStyle::Thin),
        _ => unreachable!(),
    };

    // Parse comma-separated ranges atomically
    let ranges = parse_multi_range(&input.range).map_err(|e| anyhow::anyhow!("{e}"))?;

    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    for (r1, c1, r2, c2) in &ranges {
        ws.set_range_format(*r1, *c1, *r2, *c2, &fmt)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
    }

    Ok(success_no_data(&format!(
        "Style '{}' applied to {}",
        input.style, input.range
    )))
}

pub fn format_as_table_header(
    store: &mut WorkbookStore,
    input: FormatAsTableHeaderInput,
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

    // Check if sheet has data
    let (_used_r1, used_c1, _used_r2, used_c2) = match ws.used_range() {
        Some(r) => r,
        None => {
            return Ok(error(
                ErrorCategory::InvalidInput,
                "Sheet is empty — no data to format",
                "Write data to the sheet before formatting as table header.",
            ));
        }
    };

    // Determine header row (1-based input, convert to 0-based)
    let header_row_1based = input.header_row.unwrap_or(1);
    let header_row = header_row_1based.saturating_sub(1);

    // Determine colors
    let bg_color = input
        .background_color
        .as_deref()
        .unwrap_or("#4472C4");
    let font_color = input
        .font_color
        .as_deref()
        .unwrap_or("#FFFFFF");

    // Apply header formatting: bold, font color, bg color, center align
    let header_fmt = zavora_xlsx::Format::new()
        .bold()
        .font_color(font_color)
        .background_color(bg_color)
        .align(zavora_xlsx::Align::Center);

    ws.set_range_format(header_row, used_c1, header_row, used_c2, &header_fmt)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    // Freeze panes at the row below the header
    let freeze_row = header_row + 1;
    ws.set_freeze_panes(freeze_row, 0)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    // Set autofilter spanning header row from column A to last used column
    ws.set_autofilter(header_row, used_c1, header_row, used_c2);

    Ok(success_no_data(&format!(
        "Table header formatted on row {} (columns {}:{})",
        header_row_1based,
        zavora_xlsx::utility::to_a1(header_row, used_c1),
        zavora_xlsx::utility::to_a1(header_row, used_c2)
    )))
}

pub fn format_as_table_range(
    store: &mut WorkbookStore,
    input: FormatAsTableRangeInput,
) -> Result<String, anyhow::Error> {
    // Determine color scheme
    let style_name = input.style.as_deref().unwrap_or("blue");
    let (header_bg, header_font, alt_row_bg) = match style_name {
        "blue" => ("#4472C4", "#FFFFFF", "#D6E4F0"),
        "green" => ("#548235", "#FFFFFF", "#E2EFDA"),
        "gray" => ("#808080", "#FFFFFF", "#F2F2F2"),
        "orange" => ("#ED7D31", "#FFFFFF", "#FCE4D6"),
        _ => {
            return Ok(error(
                ErrorCategory::InvalidInput,
                &format!("Unknown table style '{}'", style_name),
                "Valid styles: blue, green, gray, orange",
            ));
        }
    };

    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };

    // Parse the range
    let (r1, c1, r2, c2) = zavora_xlsx::utility::parse_range_ref(&input.range)
        .map_err(|e| anyhow::anyhow!("Invalid range '{}': {}", input.range, e))?;

    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    // Apply header styling to first row of range (bold, bg color, white font)
    let header_fmt = zavora_xlsx::Format::new()
        .bold()
        .font_color(header_font)
        .background_color(header_bg)
        .border(zavora_xlsx::BorderStyle::Thin);

    ws.set_range_format(r1, c1, r1, c2, &header_fmt)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    // Apply alternating row shading to data rows + thin borders
    let alt_fmt = zavora_xlsx::Format::new()
        .background_color(alt_row_bg)
        .border(zavora_xlsx::BorderStyle::Thin);
    let plain_fmt = zavora_xlsx::Format::new()
        .border(zavora_xlsx::BorderStyle::Thin);

    let mut shade_index: usize = 0;
    for r in (r1 + 1)..=r2 {
        let fmt = if shade_index % 2 == 0 {
            &alt_fmt
        } else {
            &plain_fmt
        };
        ws.set_range_format(r, c1, r, c2, fmt)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
        shade_index += 1;
    }

    // Autofit columns within range
    ws.autofit().map_err(|e| anyhow::anyhow!("{e}"))?;

    Ok(success_no_data(&format!(
        "Table range formatted with '{}' style on {}",
        style_name, input.range
    )))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::store::WorkbookStore;
    use proptest::prelude::*;

    // ── Property 4: Copy Format Fidelity ──
    // **Validates: Requirements 4.1, 4.3, 4.4**
    //
    // Set random formatting on source, copy to target, read back and compare cell-by-cell.

    proptest! {
        #[test]
        fn prop_copy_format_fidelity(
            bold in proptest::bool::ANY,
            italic in proptest::bool::ANY,
            font_size in prop_oneof![Just(11.0), Just(12.0), Just(14.0), Just(16.0)],
        ) {
            let mut store = WorkbookStore::new();
            let result = crate::tools::workbook::create_workbook(&mut store).unwrap();
            let v: serde_json::Value = serde_json::from_str(&result).unwrap();
            let wid = v["data"]["workbook_id"].as_str().unwrap().to_string();

            // Write some data so cells exist
            crate::tools::write::write_cells(
                &mut store,
                crate::types::inputs::WriteCellsInput {
                    workbook_id: wid.clone(),
                    sheet_name: "Sheet1".into(),
                    cells: vec![
                        crate::types::inputs::CellWrite { cell: "A1".into(), value: serde_json::json!("src") },
                        crate::types::inputs::CellWrite { cell: "B1".into(), value: serde_json::json!("src2") },
                        crate::types::inputs::CellWrite { cell: "D1".into(), value: serde_json::json!("tgt") },
                        crate::types::inputs::CellWrite { cell: "E1".into(), value: serde_json::json!("tgt2") },
                    ],
                },
            ).unwrap();

            // Apply formatting to source range A1:B1
            let mut fmt = zavora_xlsx::Format::new();
            if bold { fmt = fmt.bold(); }
            if italic { fmt = fmt.italic(); }
            fmt = fmt.font_size(font_size);

            {
                let entry = store.get_mut(&wid).unwrap();
                let ws = entry.data.worksheet(0).unwrap();
                ws.set_range_format(0, 0, 0, 1, &fmt).unwrap();
            }

            // Copy format from A1:B1 to D1:E1
            let copy_result = copy_format(
                &mut store,
                CopyFormatInput {
                    workbook_id: wid.clone(),
                    sheet_name: "Sheet1".into(),
                    source_range: "A1:B1".into(),
                    target_ranges: vec!["D1:E1".into()],
                },
            ).unwrap();
            let v: serde_json::Value = serde_json::from_str(&copy_result).unwrap();
            prop_assert_eq!(v["status"].as_str(), Some("success"));

            // Read back and compare cell-by-cell
            let entry = store.get_mut(&wid).unwrap();
            let ws = entry.data.worksheet(0).unwrap();

            for col_offset in 0u16..=1 {
                let src_fmt = ws.cell_format(0, col_offset);
                let tgt_fmt = ws.cell_format(0, 3 + col_offset); // D=3, E=4

                let src_present = src_fmt.is_some();
                let tgt_present = tgt_fmt.is_some();

                match (src_fmt, tgt_fmt) {
                    (Some(s), Some(t)) => {
                        prop_assert_eq!(s.is_bold(), t.is_bold(),
                            "Bold mismatch at col offset {}", col_offset);
                        prop_assert_eq!(s.is_italic(), t.is_italic(),
                            "Italic mismatch at col offset {}", col_offset);
                        prop_assert_eq!(s.get_font_size(), t.get_font_size(),
                            "Font size mismatch at col offset {}", col_offset);
                    }
                    (None, None) => { /* both unformatted, ok */ }
                    _ => {
                        prop_assert!(false,
                            "Format presence mismatch at col offset {}: src={}, tgt={}",
                            col_offset, src_present, tgt_present);
                    }
                }
            }
        }
    }

    // ── Helper functions for tests ──

    fn helper_create_workbook(store: &mut WorkbookStore) -> String {
        let result = crate::tools::workbook::create_workbook(store).unwrap();
        let v: serde_json::Value = serde_json::from_str(&result).unwrap();
        v["data"]["workbook_id"].as_str().unwrap().to_string()
    }

    fn helper_write_cells(store: &mut WorkbookStore, wid: &str, cells: Vec<(&str, serde_json::Value)>) {
        let cell_writes: Vec<crate::types::inputs::CellWrite> = cells
            .into_iter()
            .map(|(cell, value)| crate::types::inputs::CellWrite {
                cell: cell.to_string(),
                value,
            })
            .collect();
        crate::tools::write::write_cells(
            store,
            crate::types::inputs::WriteCellsInput {
                workbook_id: wid.to_string(),
                sheet_name: "Sheet1".into(),
                cells: cell_writes,
            },
        )
        .unwrap();
    }

    /// Save workbook to a temp file and reopen it so that cell_format() works.
    /// Returns the new workbook_id in the same store.
    fn helper_save_reopen(store: &mut WorkbookStore, wid: &str) -> String {
        use std::sync::atomic::{AtomicU64, Ordering};
        static COUNTER: AtomicU64 = AtomicU64::new(0);
        let n = COUNTER.fetch_add(1, Ordering::Relaxed);
        let tmp = format!("/tmp/test_fmt_{}.xlsx", n);
        crate::tools::workbook::save_workbook(store,
            crate::types::inputs::SaveWorkbookInput {
                workbook_id: wid.to_string(), file_path: tmp.clone(),
            }).unwrap();
        let open_result = crate::tools::workbook::open_workbook(store,
            crate::types::inputs::OpenWorkbookInput {
                file_path: tmp, read_only: false,
            }).unwrap();
        let ov: serde_json::Value = serde_json::from_str(&open_result).unwrap();
        ov["data"]["workbook_id"].as_str().unwrap().to_string()
    }

    fn helper_parse_response(json: &str) -> serde_json::Value {
        serde_json::from_str(json).unwrap()
    }

    // ── Property 15: Format as Table Range Consistency ──
    // **Validates: Requirements 7.1, 7.2, 7.3**
    proptest! {
        #[test]
        fn prop_table_range_consistency(
            num_rows in 2u32..=10u32,
            num_cols in 1u16..=5u16,
            style_idx in 0u8..4u8,
        ) {
            let style_name = match style_idx {
                0 => "blue",
                1 => "green",
                2 => "gray",
                _ => "orange",
            };

            let mut store = WorkbookStore::new();
            let wid = helper_create_workbook(&mut store);

            // Write data to fill the range
            let end_col_letter = zavora_xlsx::utility::to_a1(0, num_cols - 1);
            let end_col_letter: String = end_col_letter.chars().take_while(|c| c.is_alphabetic()).collect();
            let range_str = format!("A1:{}{}", end_col_letter, num_rows);

            let mut cell_writes = Vec::new();
            for r in 1..=num_rows {
                for c in 0..num_cols {
                    let cell_ref = zavora_xlsx::utility::to_a1(r - 1, c);
                    cell_writes.push(crate::types::inputs::CellWrite {
                        cell: cell_ref,
                        value: serde_json::json!(format!("d{}_{}", r, c)),
                    });
                }
            }
            crate::tools::write::write_cells(
                &mut store,
                crate::types::inputs::WriteCellsInput {
                    workbook_id: wid.clone(),
                    sheet_name: "Sheet1".into(),
                    cells: cell_writes,
                },
            ).unwrap();

            let result = format_as_table_range(
                &mut store,
                FormatAsTableRangeInput {
                    workbook_id: wid.clone(),
                    sheet_name: "Sheet1".into(),
                    range: range_str,
                    style: Some(style_name.to_string()),
                },
            ).unwrap();
            let v: serde_json::Value = serde_json::from_str(&result).unwrap();
            prop_assert_eq!(v["status"].as_str(), Some("success"));

            // Save/reopen so cell_format works
            let wid2 = helper_save_reopen(&mut store, &wid);
            let entry = store.get_mut(&wid2).unwrap();
            let ws = entry.data.worksheet(0).unwrap();

            // 1. First row has bold + background color
            for c in 0..num_cols {
                let fmt = ws.cell_format(0, c);
                prop_assert!(fmt.is_some(), "Header ({}, {}) needs format", 0, c);
                let fmt = fmt.unwrap();
                prop_assert!(fmt.is_bold(), "Header ({}, {}) needs bold", 0, c);
                prop_assert!(fmt.get_bg_color().is_some(), "Header ({}, {}) needs bg", 0, c);
            }

            // 2. All cells have border styling
            for r in 0..num_rows {
                for c in 0..num_cols {
                    let fmt = ws.cell_format(r, c);
                    prop_assert!(fmt.is_some(), "Cell ({}, {}) needs format", r, c);
                    let fmt = fmt.unwrap();
                    let has_border =
                        fmt.get_border_top() != zavora_xlsx::BorderStyle::None
                        || fmt.get_border_bottom() != zavora_xlsx::BorderStyle::None
                        || fmt.get_border_left() != zavora_xlsx::BorderStyle::None
                        || fmt.get_border_right() != zavora_xlsx::BorderStyle::None;
                    prop_assert!(has_border, "Cell ({}, {}) needs borders", r, c);
                }
            }

            // 3. Data rows alternate shading
            if num_rows > 1 {
                let mut shade_idx: usize = 0;
                for r in 1..num_rows {
                    let fmt = ws.cell_format(r, 0).unwrap();
                    if shade_idx % 2 == 0 {
                        prop_assert!(fmt.get_bg_color().is_some(),
                            "Row {} (shade {}) needs bg", r, shade_idx);
                    }
                    shade_idx += 1;
                }
            }
        }
    }

    // ── Property 2: Batch Format Equivalence ──
    // **Validates: Requirements 2.1, 2.4**
    proptest! {
        #[test]
        fn prop_batch_format_equivalence(
            num_ops in 1usize..=3usize,
            bold_flags in proptest::collection::vec(proptest::bool::ANY, 1..=3),
            italic_flags in proptest::collection::vec(proptest::bool::ANY, 1..=3),
            font_sizes in proptest::collection::vec(
                prop_oneof![Just(11.0), Just(12.0), Just(14.0), Just(16.0)],
                1..=3
            ),
        ) {
            let n = num_ops.min(bold_flags.len()).min(italic_flags.len()).min(font_sizes.len());

            let mut store1 = WorkbookStore::new();
            let wid1 = helper_create_workbook(&mut store1);
            let mut store2 = WorkbookStore::new();
            let wid2 = helper_create_workbook(&mut store2);

            // Write identical data to both stores
            for s in [(&mut store1, &wid1), (&mut store2, &wid2)] {
                let mut cells = Vec::new();
                for i in 0..n {
                    cells.push(crate::types::inputs::CellWrite {
                        cell: format!("A{}", i + 1),
                        value: serde_json::json!(format!("d{}", i)),
                    });
                }
                crate::tools::write::write_cells(s.0,
                    crate::types::inputs::WriteCellsInput {
                        workbook_id: s.1.clone(), sheet_name: "Sheet1".into(), cells,
                    }).unwrap();
            }

            // Build operations helper
            fn make_ops(n: usize, bold_flags: &[bool], italic_flags: &[bool], font_sizes: &[f64])
                -> Vec<crate::types::inputs::FormatOperation>
            {
                let mut ops = Vec::new();
                for i in 0..n {
                    ops.push(crate::types::inputs::FormatOperation {
                        range: format!("A{}:A{}", i + 1, i + 1),
                        bold: if bold_flags[i] { Some(true) } else { None },
                        italic: if italic_flags[i] { Some(true) } else { None },
                        underline: None, font_size: Some(font_sizes[i]),
                        font_color: None, background_color: None, number_format: None,
                        horizontal_alignment: None, vertical_alignment: None, border_style: None,
                    });
                }
                ops
            }

            let ops1 = make_ops(n, &bold_flags, &italic_flags, &font_sizes);
            let ops2 = make_ops(n, &bold_flags, &italic_flags, &font_sizes);

            batch_format(&mut store1, BatchFormatInput {
                workbook_id: wid1.clone(), sheet_name: "Sheet1".into(), operations: ops1,
            }).unwrap();

            for op in &ops2 {
                set_cell_format(&mut store2, crate::types::inputs::SetCellFormatInput {
                    workbook_id: wid2.clone(), sheet_name: "Sheet1".into(),
                    range: op.range.clone(), bold: op.bold, italic: op.italic,
                    underline: op.underline, font_size: op.font_size,
                    font_color: op.font_color.clone(), background_color: op.background_color.clone(),
                    number_format: op.number_format.clone(),
                    horizontal_alignment: None, vertical_alignment: None, border_style: None,
                }).unwrap();
            }

            // Compare formatting cell-by-cell (save/reopen both)
            let wid1r = helper_save_reopen(&mut store1, &wid1);
            let e1 = store1.get_mut(&wid1r).unwrap();
            let ws1 = e1.data.worksheet(0).unwrap();
            let mut f1s = Vec::new();
            for i in 0..n { f1s.push(ws1.cell_format(i as u32, 0)); }

            let wid2r = helper_save_reopen(&mut store2, &wid2);
            let e2 = store2.get_mut(&wid2r).unwrap();
            let ws2 = e2.data.worksheet(0).unwrap();

            for i in 0..n {
                let f2 = ws2.cell_format(i as u32, 0);
                match (&f1s[i], &f2) {
                    (Some(a), Some(b)) => {
                        prop_assert_eq!(a.is_bold(), b.is_bold(), "Bold row {}", i);
                        prop_assert_eq!(a.is_italic(), b.is_italic(), "Italic row {}", i);
                        prop_assert_eq!(a.get_font_size(), b.get_font_size(), "Size row {}", i);
                    }
                    (None, None) => {}
                    _ => { prop_assert!(false, "Presence mismatch row {}", i); }
                }
            }
        }
    }

    // ── Unit Tests for Tier 1 Formatting Tools ──

    #[test]
    fn test_batch_format_empty_operations() {
        let mut store = WorkbookStore::new();
        let wid = helper_create_workbook(&mut store);
        let result = batch_format(&mut store, BatchFormatInput {
            workbook_id: wid, sheet_name: "Sheet1".into(), operations: vec![],
        }).unwrap();
        let v = helper_parse_response(&result);
        assert_eq!(v["status"].as_str(), Some("success"));
        assert_eq!(v["data"]["operations_applied"].as_u64(), Some(0));
        assert_eq!(v["data"]["failures"].as_array().unwrap().len(), 0);
    }

    #[test]
    fn test_batch_format_partial_failure() {
        let mut store = WorkbookStore::new();
        let wid = helper_create_workbook(&mut store);
        helper_write_cells(&mut store, &wid, vec![("A1", serde_json::json!("hi"))]);
        let result = batch_format(&mut store, BatchFormatInput {
            workbook_id: wid, sheet_name: "Sheet1".into(),
            operations: vec![
                crate::types::inputs::FormatOperation {
                    range: "A1:A1".into(), bold: Some(true),
                    italic: None, underline: None, font_size: None, font_color: None,
                    background_color: None, number_format: None, horizontal_alignment: None,
                    vertical_alignment: None, border_style: None,
                },
                crate::types::inputs::FormatOperation {
                    range: "INVALID".into(), bold: Some(true),
                    italic: None, underline: None, font_size: None, font_color: None,
                    background_color: None, number_format: None, horizontal_alignment: None,
                    vertical_alignment: None, border_style: None,
                },
                crate::types::inputs::FormatOperation {
                    range: "A1:A1".into(), italic: Some(true),
                    bold: None, underline: None, font_size: None, font_color: None,
                    background_color: None, number_format: None, horizontal_alignment: None,
                    vertical_alignment: None, border_style: None,
                },
            ],
        }).unwrap();
        let v = helper_parse_response(&result);
        assert_eq!(v["status"].as_str(), Some("success"));
        assert_eq!(v["data"]["operations_applied"].as_u64(), Some(2));
        assert_eq!(v["data"]["failures"].as_array().unwrap().len(), 1);
        assert_eq!(v["data"]["failures"][0]["operation_index"].as_u64(), Some(1));
    }

    #[test]
    fn test_apply_theme_financial() {
        let mut store = WorkbookStore::new();
        let wid = helper_create_workbook(&mut store);
        helper_write_cells(&mut store, &wid, vec![
            ("A1", serde_json::json!("Name")), ("B1", serde_json::json!("Amt")),
            ("A2", serde_json::json!("Alice")), ("B2", serde_json::json!(100)),
            ("A3", serde_json::json!("Bob")), ("B3", serde_json::json!(200)),
            ("A4", serde_json::json!("Total")), ("B4", serde_json::json!(300)),
        ]);
        let result = apply_theme(&mut store, ApplyThemeInput {
            workbook_id: wid.clone(), sheet_name: "Sheet1".into(),
            theme: "financial_professional".into(),
            header_rows: vec![1], total_rows: vec![4], auto_detect_formats: false,
        }).unwrap();
        let v = helper_parse_response(&result);
        assert_eq!(v["status"].as_str(), Some("success"));
        // Save/reopen to verify formatting
        let wid2 = helper_save_reopen(&mut store, &wid);
        let entry = store.get_mut(&wid2).unwrap();
        let ws = entry.data.worksheet(0).unwrap();
        let hdr = ws.cell_format(0, 0).expect("Header format");
        assert!(hdr.is_bold());
        assert!(hdr.get_bg_color().is_some());
    }

    #[test]
    fn test_apply_theme_corporate() {
        let mut store = WorkbookStore::new();
        let wid = helper_create_workbook(&mut store);
        helper_write_cells(&mut store, &wid, vec![
            ("A1", serde_json::json!("Name")), ("A2", serde_json::json!("Alice")),
        ]);
        let result = apply_theme(&mut store, ApplyThemeInput {
            workbook_id: wid.clone(), sheet_name: "Sheet1".into(),
            theme: "corporate".into(),
            header_rows: vec![1], total_rows: vec![], auto_detect_formats: false,
        }).unwrap();
        assert_eq!(helper_parse_response(&result)["status"].as_str(), Some("success"));
        let wid2 = helper_save_reopen(&mut store, &wid);
        let entry = store.get_mut(&wid2).unwrap();
        let ws = entry.data.worksheet(0).unwrap();
        let hdr = ws.cell_format(0, 0).expect("Corporate header");
        assert!(hdr.is_bold());
        assert!(hdr.get_bg_color().is_some());
    }

    #[test]
    fn test_apply_theme_minimal() {
        let mut store = WorkbookStore::new();
        let wid = helper_create_workbook(&mut store);
        helper_write_cells(&mut store, &wid, vec![
            ("A1", serde_json::json!("Name")), ("A2", serde_json::json!("Alice")),
        ]);
        let result = apply_theme(&mut store, ApplyThemeInput {
            workbook_id: wid.clone(), sheet_name: "Sheet1".into(),
            theme: "minimal".into(),
            header_rows: vec![1], total_rows: vec![], auto_detect_formats: false,
        }).unwrap();
        assert_eq!(helper_parse_response(&result)["status"].as_str(), Some("success"));
        let wid2 = helper_save_reopen(&mut store, &wid);
        let entry = store.get_mut(&wid2).unwrap();
        let ws = entry.data.worksheet(0).unwrap();
        let hdr = ws.cell_format(0, 0).expect("Minimal header");
        assert!(hdr.is_bold());
    }

    #[test]
    fn test_apply_theme_invalid_name() {
        let mut store = WorkbookStore::new();
        let wid = helper_create_workbook(&mut store);
        helper_write_cells(&mut store, &wid, vec![("A1", serde_json::json!("d"))]);
        let result = apply_theme(&mut store, ApplyThemeInput {
            workbook_id: wid, sheet_name: "Sheet1".into(),
            theme: "nonexistent".into(),
            header_rows: vec![], total_rows: vec![], auto_detect_formats: false,
        }).unwrap();
        let v = helper_parse_response(&result);
        assert_eq!(v["status"].as_str(), Some("error"));
        assert!(v["message"].as_str().unwrap().contains("nonexistent"));
    }

    #[test]
    fn test_copy_format_no_formatting() {
        let mut store = WorkbookStore::new();
        let wid = helper_create_workbook(&mut store);
        helper_write_cells(&mut store, &wid, vec![
            ("A1", serde_json::json!("s")), ("C1", serde_json::json!("t")),
        ]);
        let result = copy_format(&mut store, CopyFormatInput {
            workbook_id: wid, sheet_name: "Sheet1".into(),
            source_range: "A1:A1".into(), target_ranges: vec!["C1:C1".into()],
        }).unwrap();
        let v = helper_parse_response(&result);
        assert_eq!(v["status"].as_str(), Some("success"));
        assert!(v["data"]["note"].as_str().is_some());
    }

    #[test]
    fn test_copy_format_tiling() {
        let mut store = WorkbookStore::new();
        let wid = helper_create_workbook(&mut store);
        helper_write_cells(&mut store, &wid, vec![
            ("A1", serde_json::json!("s")),
            ("C1", serde_json::json!("t1")), ("D1", serde_json::json!("t2")),
            ("C2", serde_json::json!("t3")), ("D2", serde_json::json!("t4")),
        ]);
        {
            let entry = store.get_mut(&wid).unwrap();
            let ws = entry.data.worksheet(0).unwrap();
            ws.set_range_format(0, 0, 0, 0, &zavora_xlsx::Format::new().bold()).unwrap();
        }
        // Save/reopen so cell_format works (copy_format reads source via cell_format)
        let wid2 = helper_save_reopen(&mut store, &wid);
        let result = copy_format(&mut store, CopyFormatInput {
            workbook_id: wid2.clone(), sheet_name: "Sheet1".into(),
            source_range: "A1:A1".into(), target_ranges: vec!["C1:D2".into()],
        }).unwrap();
        let v = helper_parse_response(&result);
        assert_eq!(v["status"].as_str(), Some("success"));
        assert!(v["data"]["note"].is_null(), "Should not have 'no formatting' note");
        // Save/reopen again to verify target formatting
        let wid3 = helper_save_reopen(&mut store, &wid2);
        let entry = store.get_mut(&wid3).unwrap();
        let ws = entry.data.worksheet(0).unwrap();
        for r in 0..=1u32 {
            for c in 2..=3u16 {
                let fmt = ws.cell_format(r, c);
                assert!(fmt.is_some(), "({},{}) should have format", r, c);
                assert!(fmt.unwrap().is_bold(), "({},{}) should be bold", r, c);
            }
        }
    }

    #[test]
    fn test_apply_style_each_preset() {
        for preset in &["header","title","currency","percentage","date","number","text","accounting","total"] {
            let mut store = WorkbookStore::new();
            let wid = helper_create_workbook(&mut store);
            helper_write_cells(&mut store, &wid, vec![("A1", serde_json::json!("d"))]);
            let result = apply_style(&mut store, ApplyStyleInput {
                workbook_id: wid.clone(), sheet_name: "Sheet1".into(),
                range: "A1:A1".into(), style: preset.to_string(),
            }).unwrap();
            assert_eq!(helper_parse_response(&result)["status"].as_str(), Some("success"),
                "preset '{}' failed", preset);
            // Save/reopen to verify formatting was applied
            let wid2 = helper_save_reopen(&mut store, &wid);
            let entry = store.get_mut(&wid2).unwrap();
            let ws = entry.data.worksheet(0).unwrap();
            assert!(ws.cell_format(0, 0).is_some(), "'{}' should apply format", preset);
        }
    }

    #[test]
    fn test_apply_style_invalid_name() {
        let mut store = WorkbookStore::new();
        let wid = helper_create_workbook(&mut store);
        helper_write_cells(&mut store, &wid, vec![("A1", serde_json::json!("d"))]);
        let result = apply_style(&mut store, ApplyStyleInput {
            workbook_id: wid, sheet_name: "Sheet1".into(),
            range: "A1:A1".into(), style: "bad_style".into(),
        }).unwrap();
        let v = helper_parse_response(&result);
        assert_eq!(v["status"].as_str(), Some("error"));
        assert!(v["message"].as_str().unwrap().contains("bad_style"));
    }

    #[test]
    fn test_format_table_header_defaults() {
        let mut store = WorkbookStore::new();
        let wid = helper_create_workbook(&mut store);
        helper_write_cells(&mut store, &wid, vec![
            ("A1", serde_json::json!("Name")), ("B1", serde_json::json!("Val")),
            ("A2", serde_json::json!("Alice")), ("B2", serde_json::json!(100)),
        ]);
        let result = format_as_table_header(&mut store, FormatAsTableHeaderInput {
            workbook_id: wid.clone(), sheet_name: "Sheet1".into(),
            header_row: None, background_color: None, font_color: None,
        }).unwrap();
        assert_eq!(helper_parse_response(&result)["status"].as_str(), Some("success"));
        let wid2 = helper_save_reopen(&mut store, &wid);
        let entry = store.get_mut(&wid2).unwrap();
        let ws = entry.data.worksheet(0).unwrap();
        let fmt = ws.cell_format(0, 0).expect("Header format");
        assert!(fmt.is_bold());
        assert!(fmt.get_bg_color().is_some());
        assert!(fmt.get_font_color().is_some());
    }

    #[test]
    fn test_format_table_header_custom_row() {
        let mut store = WorkbookStore::new();
        let wid = helper_create_workbook(&mut store);
        helper_write_cells(&mut store, &wid, vec![
            ("A1", serde_json::json!("Title")),
            ("A2", serde_json::json!("Name")), ("B2", serde_json::json!("Val")),
            ("A3", serde_json::json!("Alice")),
        ]);
        let result = format_as_table_header(&mut store, FormatAsTableHeaderInput {
            workbook_id: wid.clone(), sheet_name: "Sheet1".into(),
            header_row: Some(2), background_color: None, font_color: None,
        }).unwrap();
        assert_eq!(helper_parse_response(&result)["status"].as_str(), Some("success"));
        let wid2 = helper_save_reopen(&mut store, &wid);
        let entry = store.get_mut(&wid2).unwrap();
        let ws = entry.data.worksheet(0).unwrap();
        let fmt = ws.cell_format(1, 0).expect("Row 2 format");
        assert!(fmt.is_bold());
    }

    #[test]
    fn test_format_table_header_empty_sheet() {
        let mut store = WorkbookStore::new();
        let wid = helper_create_workbook(&mut store);
        let result = format_as_table_header(&mut store, FormatAsTableHeaderInput {
            workbook_id: wid, sheet_name: "Sheet1".into(),
            header_row: None, background_color: None, font_color: None,
        }).unwrap();
        let v = helper_parse_response(&result);
        assert_eq!(v["status"].as_str(), Some("error"));
        assert!(v["message"].as_str().unwrap().to_lowercase().contains("empty"));
    }

    #[test]
    fn test_format_table_range_default_blue() {
        let mut store = WorkbookStore::new();
        let wid = helper_create_workbook(&mut store);
        helper_write_cells(&mut store, &wid, vec![
            ("A1", serde_json::json!("Hdr")),
            ("A2", serde_json::json!("D1")),
            ("A3", serde_json::json!("D2")),
        ]);
        let result = format_as_table_range(&mut store, FormatAsTableRangeInput {
            workbook_id: wid.clone(), sheet_name: "Sheet1".into(),
            range: "A1:A3".into(), style: None,
        }).unwrap();
        assert_eq!(helper_parse_response(&result)["status"].as_str(), Some("success"));
        let wid2 = helper_save_reopen(&mut store, &wid);
        let entry = store.get_mut(&wid2).unwrap();
        let ws = entry.data.worksheet(0).unwrap();
        let hdr = ws.cell_format(0, 0).expect("Header format");
        assert!(hdr.is_bold());
        assert!(hdr.get_bg_color().is_some());
        let d1 = ws.cell_format(1, 0).expect("Data row format");
        let has_border = d1.get_border_top() != zavora_xlsx::BorderStyle::None
            || d1.get_border_bottom() != zavora_xlsx::BorderStyle::None;
        assert!(has_border, "Data row should have borders");
    }

    #[test]
    fn test_format_table_range_each_style() {
        for style in &["blue", "green", "gray", "orange"] {
            let mut store = WorkbookStore::new();
            let wid = helper_create_workbook(&mut store);
            helper_write_cells(&mut store, &wid, vec![
                ("A1", serde_json::json!("H")), ("A2", serde_json::json!("D")),
            ]);
            let result = format_as_table_range(&mut store, FormatAsTableRangeInput {
                workbook_id: wid.clone(), sheet_name: "Sheet1".into(),
                range: "A1:A2".into(), style: Some(style.to_string()),
            }).unwrap();
            assert_eq!(helper_parse_response(&result)["status"].as_str(), Some("success"),
                "style '{}' failed", style);
            let wid2 = helper_save_reopen(&mut store, &wid);
            let entry = store.get_mut(&wid2).unwrap();
            let ws = entry.data.worksheet(0).unwrap();
            assert!(ws.cell_format(0, 0).unwrap().is_bold(), "'{}' header bold", style);
        }
    }

    #[test]
    fn test_describe_formatting_empty() {
        let mut store = WorkbookStore::new();
        let wid = helper_create_workbook(&mut store);
        helper_write_cells(&mut store, &wid, vec![("A1", serde_json::json!("d"))]);
        let result = crate::tools::read::describe_formatting(&mut store,
            crate::types::inputs::DescribeFormattingInput {
                workbook_id: wid, sheet_name: "Sheet1".into(), range: "A1:A1".into(),
            }).unwrap();
        let v = helper_parse_response(&result);
        assert_eq!(v["status"].as_str(), Some("success"));
        assert_eq!(v["data"]["format_groups"].as_array().unwrap().len(), 0);
    }

    #[test]
    fn test_describe_formatting_grouped() {
        let mut store = WorkbookStore::new();
        let wid = helper_create_workbook(&mut store);
        helper_write_cells(&mut store, &wid, vec![
            ("A1", serde_json::json!("d1")),
            ("A2", serde_json::json!("d2")),
            ("A3", serde_json::json!("d3")),
        ]);
        {
            let entry = store.get_mut(&wid).unwrap();
            let ws = entry.data.worksheet(0).unwrap();
            ws.set_range_format(0, 0, 1, 0, &zavora_xlsx::Format::new().bold()).unwrap();
            ws.set_range_format(2, 0, 2, 0, &zavora_xlsx::Format::new().italic()).unwrap();
        }
        // Save/reopen so cell_format works (used by describe_formatting)
        let wid2 = helper_save_reopen(&mut store, &wid);
        let result = crate::tools::read::describe_formatting(&mut store,
            crate::types::inputs::DescribeFormattingInput {
                workbook_id: wid2, sheet_name: "Sheet1".into(), range: "A1:A3".into(),
            }).unwrap();
        let v = helper_parse_response(&result);
        assert_eq!(v["status"].as_str(), Some("success"));
        let groups = v["data"]["format_groups"].as_array().unwrap();
        assert_eq!(groups.len(), 2, "Should have 2 format groups");
        let bold_group = groups.iter().find(|g| g["bold"] == serde_json::json!(true));
        assert!(bold_group.is_some(), "Should have bold group");
        assert_eq!(bold_group.unwrap()["ranges"].as_array().unwrap().len(), 2);
    }
}

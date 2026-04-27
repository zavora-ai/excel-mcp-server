use super::common::workbook_not_found;
use crate::engines::zavora;
use crate::store::WorkbookStore;
use crate::types::inputs::*;
use crate::types::responses::*;

pub fn read_sheet(
    store: &mut WorkbookStore,
    input: ReadSheetInput,
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
    let (r1, c1, r2, c2) = if let Some(ref range_str) = input.range {
        zavora_xlsx::utility::parse_range_ref(range_str).map_err(|e| anyhow::anyhow!("{e}"))?
    } else {
        match ws.used_range() {
            Some(r) => r,
            None => {
                return Ok(success(
                    "Sheet is empty",
                    ReadSheetData {
                        rows: vec![],
                        total_rows: 0,
                        page_rows: 0,
                        continuation_token: None,
                    },
                ))
            }
        }
    };
    let offset = input
        .continuation_token
        .as_ref()
        .and_then(|t| serde_json::from_str::<ContinuationToken>(t).ok())
        .map(|t| t.offset)
        .unwrap_or(0);
    let page_size = 100u32;
    let start = r1 + offset;
    let end = (start + page_size - 1).min(r2);
    let mut rows = Vec::new();
    for r in start..=end {
        let mut row = Vec::new();
        for c in c1..=c2 {
            row.push(zavora::cell_to_json(&ws.read_cell(r, c)));
        }
        rows.push(row);
    }
    let total = r2 - r1 + 1;
    let token = if end < r2 {
        Some(serde_json::to_string(&ContinuationToken {
            sheet: input.sheet_name.clone(),
            offset: end - r1 + 1,
            range: input.range.clone(),
        })?)
    } else {
        None
    };
    Ok(success(
        "Sheet data read",
        ReadSheetData {
            rows,
            total_rows: total,
            page_rows: end - start + 1,
            continuation_token: token,
        },
    ))
}

pub fn read_cell(store: &mut WorkbookStore, input: ReadCellInput) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let idx = match find_sheet(&entry.data, &input.sheet_name) {
        Some(i) => i,
        None => return Ok(sheet_err(&input.sheet_name)),
    };
    let (row, col) =
        zavora_xlsx::utility::parse_cell_ref(&input.cell).map_err(|e| anyhow::anyhow!("{e}"))?;
    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    let val = ws.read_cell(row, col);
    let formula = if let zavora_xlsx::CellValue::Formula { ref formula, .. } = val {
        Some(formula.clone())
    } else {
        None
    };
    Ok(success(
        "Cell read",
        CellData {
            cell: input.cell,
            value: zavora::cell_to_json(&val),
            value_type: zavora::cell_type_name(&val).to_string(),
            formula,
        },
    ))
}

pub fn search_cells(
    store: &mut WorkbookStore,
    input: SearchCellsInput,
) -> Result<String, anyhow::Error> {
    let entry = match store.get_mut(&input.workbook_id) {
        Some(e) => e,
        None => return Ok(workbook_not_found(store, &input.workbook_id)),
    };
    let names: Vec<String> = match &input.sheet_name {
        Some(n) => vec![n.clone()],
        None => entry
            .data
            .sheet_names()
            .iter()
            .map(|s| s.to_string())
            .collect(),
    };
    let mut matches = Vec::new();
    let max = 200;
    let query_lower = input.query.to_lowercase();
    let is_exact = matches!(input.match_mode, crate::types::enums::MatchMode::Exact);
    for name in &names {
        let idx = match find_sheet(&entry.data, name) {
            Some(i) => i,
            None => continue,
        };
        let ws = entry
            .data
            .worksheet(idx)
            .map_err(|e| anyhow::anyhow!("{e}"))?;
        if let Some((r1, c1, r2, c2)) = ws.used_range() {
            for r in r1..=r2 {
                for c in c1..=c2 {
                    if matches.len() >= max {
                        break;
                    }
                    let val = ws.read_cell(r, c);
                    let s = match &val {
                        zavora_xlsx::CellValue::String(s) => s.clone(),
                        zavora_xlsx::CellValue::Number(n) => format!("{n}"),
                        zavora_xlsx::CellValue::Bool(b) => b.to_string(),
                        _ => continue,
                    };
                    let hit = if is_exact {
                        s.to_lowercase() == query_lower
                    } else {
                        s.to_lowercase().contains(&query_lower)
                    };
                    if hit {
                        matches.push(SearchMatch {
                            sheet: name.clone(),
                            cell: zavora_xlsx::utility::to_a1(r, c),
                            value: zavora::cell_to_json(&val),
                        });
                    }
                }
            }
        }
    }
    let truncated = matches.len() >= max;
    let total = matches.len();
    Ok(success(
        "Search complete",
        SearchResult {
            matches,
            total_matches: total,
            truncated,
        },
    ))
}

pub fn sheet_to_csv(
    store: &mut WorkbookStore,
    input: SheetToCsvInput,
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
    let delim = input.delimiter.as_bytes().first().copied().unwrap_or(b',');
    let mut buf = Vec::new();
    ws.to_csv(&mut buf, delim)
        .map_err(|e| anyhow::anyhow!("{e}"))?;
    let csv = String::from_utf8(buf).unwrap_or_default();
    let rows = csv.lines().count() as u32;
    Ok(success(
        "CSV exported",
        CsvExportData {
            csv,
            total_rows: rows,
            truncated: false,
        },
    ))
}

pub fn describe_formatting(
    store: &mut WorkbookStore,
    input: DescribeFormattingInput,
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
        .map_err(|e| anyhow::anyhow!("Invalid range '{}': {}", input.range, e))?;

    let ws = entry
        .data
        .worksheet(idx)
        .map_err(|e| anyhow::anyhow!("{e}"))?;

    // Collect formatting for each cell, grouping identical formats together.
    // We use a serializable key to group cells with the same formatting.
    use std::collections::BTreeMap;

    /// A hashable/comparable representation of a cell's formatting.
    #[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
    struct FormatKey {
        bold: Option<bool>,
        italic: Option<bool>,
        underline: Option<bool>,
        font_size_bits: Option<u64>, // f64 bits for Ord
        font_color: Option<String>,
        background_color: Option<String>,
        number_format: Option<String>,
        h_align: Option<String>,
        v_align: Option<String>,
        border_style: Option<String>,
    }

    fn extract_format_key(fmt: &zavora_xlsx::Format) -> FormatKey {
        let bold = if fmt.is_bold() { Some(true) } else { None };
        let italic = if fmt.is_italic() { Some(true) } else { None };
        let underline = if fmt.get_underline() != zavora_xlsx::Underline::None {
            Some(true)
        } else {
            None
        };
        let fs = fmt.get_font_size();
        let font_size_bits = if fs != 0.0 {
            Some(fs.to_bits())
        } else {
            None
        };
        let font_color = fmt.get_font_color().map(|rgb| {
            format!("#{:02X}{:02X}{:02X}", rgb[0], rgb[1], rgb[2])
        });
        let background_color = fmt.get_bg_color().map(|rgb| {
            format!("#{:02X}{:02X}{:02X}", rgb[0], rgb[1], rgb[2])
        });
        let nf = fmt.get_num_format();
        let number_format = if nf.is_empty() { None } else { Some(nf.to_string()) };
        let h_align = match fmt.get_h_align() {
            1 => Some("left".to_string()),
            2 => Some("center".to_string()),
            3 => Some("right".to_string()),
            4 => Some("fill".to_string()),
            5 => Some("justify".to_string()),
            _ => None,
        };
        let v_align = match fmt.get_v_align() {
            1 => Some("center".to_string()),
            2 => Some("bottom".to_string()),
            _ => None,
        };
        // Summarize border style
        let bt = fmt.get_border_top();
        let bb = fmt.get_border_bottom();
        let bl = fmt.get_border_left();
        let br = fmt.get_border_right();
        let border_style = if bt != zavora_xlsx::BorderStyle::None
            || bb != zavora_xlsx::BorderStyle::None
            || bl != zavora_xlsx::BorderStyle::None
            || br != zavora_xlsx::BorderStyle::None
        {
            Some(format!("top:{:?},bottom:{:?},left:{:?},right:{:?}", bt, bb, bl, br))
        } else {
            None
        };

        FormatKey {
            bold,
            italic,
            underline,
            font_size_bits,
            font_color,
            background_color,
            number_format,
            h_align,
            v_align,
            border_style,
        }
    }

    let mut groups: BTreeMap<FormatKey, Vec<String>> = BTreeMap::new();

    for r in r1..=r2 {
        for c in c1..=c2 {
            if let Some(fmt) = ws.cell_format(r, c) {
                let key = extract_format_key(&fmt);
                let cell_ref = zavora_xlsx::utility::to_a1(r, c);
                groups.entry(key).or_default().push(cell_ref);
            }
        }
    }

    if groups.is_empty() {
        return Ok(success(
            "No formatting found in range",
            DescribeFormattingResult {
                format_groups: vec![],
            },
        ));
    }

    let format_groups: Vec<FormatGroup> = groups
        .into_iter()
        .map(|(key, ranges)| FormatGroup {
            ranges,
            bold: key.bold,
            italic: key.italic,
            underline: key.underline,
            font_size: key.font_size_bits.map(f64::from_bits),
            font_color: key.font_color,
            background_color: key.background_color,
            number_format: key.number_format,
            horizontal_alignment: key.h_align,
            vertical_alignment: key.v_align,
            border_style: key.border_style,
        })
        .collect();

    Ok(success(
        "Formatting described",
        DescribeFormattingResult { format_groups },
    ))
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

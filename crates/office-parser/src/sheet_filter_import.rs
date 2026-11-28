//! Typed project criteria and a conservative projection of standard AutoFilter.
//! The complete resolved XML tree is retained separately for unprojected Office
//! criteria; a partial projection is never returned as a complete filter.

use std::collections::{BTreeMap, BTreeSet};

use office_io_common::OfficeFilterNode;
use serde::{Deserialize, Deserializer, Serialize};
use serde_json::Value;

use crate::model::CellPayload;

const MAX_ROW_INDEX: u32 = 1_048_575;
const MAX_COLUMN_INDEX: u32 = 16_383;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct AutoFilterRange {
    pub start_row: u32,
    pub start_col: u32,
    pub end_row: u32,
    pub end_col: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct AutoFilterState {
    pub range: AutoFilterRange,
    pub columns: BTreeMap<u32, FilterCriterion>,
    pub filtered_rows: Vec<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", deny_unknown_fields)]
pub enum FilterCriterion {
    Values { #[serde(rename = "selectedItems")] selected_items: Vec<Value> },
    Color {
        #[serde(default, rename = "cellFillColors", skip_serializing_if = "Option::is_none")]
        cell_fill_colors: Option<Vec<Option<String>>>,
        #[serde(default, rename = "cellTextColors", skip_serializing_if = "Option::is_none")]
        cell_text_colors: Option<Vec<String>>,
    },
    Condition {
        operator: String,
        #[serde(default, skip_serializing_if = "Option::is_none", deserialize_with = "present_value")]
        value: Option<Value>,
        #[serde(default, rename = "secondValue", skip_serializing_if = "Option::is_none", deserialize_with = "present_value")]
        second_value: Option<Value>,
        #[serde(default, rename = "caseSensitive", skip_serializing_if = "Option::is_none")]
        case_sensitive: Option<bool>,
    },
}

fn present_value<'de, D: Deserializer<'de>>(deserializer: D) -> Result<Option<Value>, D::Error> {
    Value::deserialize(deserializer).map(Some)
}
fn scalar(value: &Value) -> bool { value.is_null() || value.is_string() || value.is_number() || value.is_boolean() }

pub(crate) fn validate(state: &AutoFilterState) -> Result<(), String> {
    let range = &state.range;
    if range.start_row >= range.end_row || range.start_col > range.end_col || range.end_row > MAX_ROW_INDEX || range.end_col > MAX_COLUMN_INDEX {
        return Err("invalid typed AutoFilter range".into());
    }
    let mut rows = BTreeSet::new();
    if state.filtered_rows.iter().any(|row| *row <= range.start_row || *row > MAX_ROW_INDEX || !rows.insert(*row)) {
        return Err("invalid or duplicate typed AutoFilter hidden row".into());
    }
    for (column, criterion) in &state.columns {
        if *column < range.start_col || *column > range.end_col { return Err("typed AutoFilter column is outside its range".into()); }
        match criterion {
            FilterCriterion::Color { cell_fill_colors, cell_text_colors } => {
                let invalid = |color: &String| color.trim().is_empty() || color.len() > 256;
                if cell_fill_colors.is_none() && cell_text_colors.is_none()
                    || cell_fill_colors.as_ref().is_some_and(|colors| colors.len() > 4096 || colors.iter().flatten().any(invalid))
                    || cell_text_colors.as_ref().is_some_and(|colors| colors.len() > 4096 || colors.iter().any(invalid)) {
                    return Err("invalid typed AutoFilter color predicate".into());
                }
            }
            FilterCriterion::Values { selected_items } if selected_items.iter().all(scalar) => {}
            FilterCriterion::Condition { operator, value, second_value, .. } => {
                let no_operand = matches!(operator.as_str(), "isBlank" | "isNotBlank" | "isError");
                let supported = no_operand || matches!(operator.as_str(), "equals" | "doesNotEqual" | "greaterThan" | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual" | "between" | "notBetween" | "contains" | "doesNotContain" | "beginsWith" | "endsWith");
                if !supported || (!no_operand && value.is_none()) || (matches!(operator.as_str(), "between" | "notBetween") && second_value.is_none())
                    || value.as_ref().is_some_and(|value| !scalar(value)) || second_value.as_ref().is_some_and(|value| !scalar(value)) {
                    return Err("invalid typed AutoFilter predicate".into());
                }
            }
            _ => return Err("typed AutoFilter members must be scalar values".into()),
        }
    }
    Ok(())
}

fn coordinate(value: &str) -> Option<(u32, u32)> {
    let value = value.replace('$', "");
    let split = value.find(|character: char| character.is_ascii_digit())?;
    let (letters, digits) = value.split_at(split);
    if letters.is_empty() || digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) { return None; }
    let mut column = 0u32;
    for letter in letters.bytes() {
        if !letter.is_ascii_alphabetic() { return None; }
        column = column.checked_mul(26)?.checked_add(u32::from(letter.to_ascii_uppercase() - b'A' + 1))?;
    }
    let row = digits.parse::<u32>().ok()?.checked_sub(1)?;
    let column = column.checked_sub(1)?;
    (row <= MAX_ROW_INDEX && column <= MAX_COLUMN_INDEX).then_some((row, column))
}

fn range(value: &str) -> Option<AutoFilterRange> {
    let (start, end) = value.split_once(':')?;
    let (start_row, start_col) = coordinate(start)?;
    let (end_row, end_col) = coordinate(end)?;
    (start_row < end_row && start_col <= end_col).then_some(AutoFilterRange { start_row, start_col, end_row, end_col })
}

fn flag(node: &OfficeFilterNode, key: &str, default: bool) -> Result<bool, String> {
    match node.attributes.get(key).map(String::as_str) {
        None => Ok(default), Some("1" | "true") => Ok(true), Some("0" | "false") => Ok(false), _ => Err(format!("invalid AutoFilter {key} flag")),
    }
}

fn cells_in_column<'a>(cells: &'a [CellPayload], range: &AutoFilterRange, column: u32) -> impl Iterator<Item = &'a CellPayload> {
    let first = range.start_row + 1;
    let last = range.end_row;
    cells.iter().filter(move |cell| u32::from(cell.col) == column && first <= cell.row && cell.row <= last)
}

fn operand(value: &str, cells: &[CellPayload], range: &AutoFilterRange, column: u32) -> Value {
    let populated = cells_in_column(cells, range, column).filter(|cell| !cell.value.is_null() && cell.value.as_str() != Some(""));
    let mut any_number = false;
    let mut only_numbers = true;
    for cell in populated { any_number |= cell.value.is_number(); only_numbers &= cell.value.is_number(); }
    if only_numbers && any_number {
        if let Ok(number) = value.parse::<f64>() {
            if let Some(number) = serde_json::Number::from_f64(number) { return Value::Number(number); }
        }
    }
    Value::String(value.into())
}

fn members(node: &OfficeFilterNode, cells: &[CellPayload], range: &AutoFilterRange, column: u32) -> Result<FilterCriterion, String> {
    let mut selected_items = Vec::new();
    if flag(node, "blank", false)? { selected_items.extend([Value::Null, Value::String(String::new())]); }
    for member in &node.children {
        if !member.is_standard("filter") { return Err("date-group or extended AutoFilter members require a separate Office criterion adapter".into()); }
        let text = member.attributes.get("val").ok_or_else(|| "AutoFilter member is missing val".to_owned())?;
        let mut found = false;
        for cell in cells_in_column(cells, range, column) {
            let raw = match &cell.value { Value::String(value) => value.clone(), Value::Null => String::new(), value => value.to_string() };
            if scalar(&cell.value) && (raw == *text || cell.display == *text) { selected_items.push(cell.value.clone()); found = true; }
        }
        if !found { selected_items.push(Value::String(text.clone())); }
    }
    let mut seen = BTreeSet::new();
    selected_items.retain(|value| seen.insert(value.to_string()));
    Ok(FilterCriterion::Values { selected_items })
}

fn unescape_pattern(value: &str) -> Result<(String, Vec<usize>), String> {
    let mut result = String::new();
    let mut stars = Vec::new();
    let mut characters = value.chars();
    while let Some(character) = characters.next() {
        match character {
            '~' => result.push(characters.next().ok_or_else(|| "incomplete AutoFilter wildcard escape".to_owned())?),
            '*' => { stars.push(result.len()); result.push('*'); }
            '?' => return Err("arbitrary AutoFilter wildcard patterns require a separate predicate adapter".into()),
            value => result.push(value),
        }
    }
    Ok((result, stars))
}

fn single_custom(node: &OfficeFilterNode, cells: &[CellPayload], range: &AutoFilterRange, column: u32) -> Result<FilterCriterion, String> {
    if !node.is_standard("customFilter") { return Err("extended AutoFilter predicate requires a separate Office adapter".into()); }
    let operation = node.attributes.get("operator").map(String::as_str).unwrap_or("equal");
    let value = node.attributes.get("val").ok_or_else(|| "AutoFilter predicate is missing val".to_owned())?;
    let mut operator = match operation {
        "equal" => "equals", "notEqual" => "doesNotEqual", "greaterThan" => "greaterThan", "greaterThanOrEqual" => "greaterThanOrEqual",
        "lessThan" => "lessThan", "lessThanOrEqual" => "lessThanOrEqual", _ => return Err("unsupported standard AutoFilter comparison".into()),
    };
    if value == " " && matches!(operation, "equal" | "notEqual") {
        return Ok(FilterCriterion::Condition { operator: if operation == "equal" { "isBlank" } else { "isNotBlank" }.into(), value: None, second_value: None, case_sensitive: None });
    }
    let (literal, stars) = unescape_pattern(value)?;
    let literal = if stars.is_empty() { literal } else if matches!(operation, "equal" | "notEqual") {
        if stars.len() == 2 && stars[0] == 0 && stars[1] + 1 == literal.len() {
            operator = if operation == "equal" { "contains" } else { "doesNotContain" };
            literal[1..literal.len() - 1].into()
        } else if stars.len() == 1 && operation == "equal" && stars[0] == 0 {
            operator = "endsWith"; literal[1..].into()
        } else if stars.len() == 1 && operation == "equal" && stars[0] + 1 == literal.len() {
            operator = "beginsWith"; literal[..literal.len() - 1].into()
        } else { return Err("arbitrary AutoFilter wildcard patterns require a separate predicate adapter".into()); }
    } else { return Err("AutoFilter ordering predicate contains a wildcard".into()); };
    Ok(FilterCriterion::Condition { operator: operator.into(), value: Some(operand(&literal, cells, range, column)), second_value: None, case_sensitive: None })
}

fn custom(node: &OfficeFilterNode, cells: &[CellPayload], range: &AutoFilterRange, column: u32) -> Result<FilterCriterion, String> {
    if node.children.len() == 1 { return single_custom(&node.children[0], cells, range, column); }
    if node.children.len() != 2 { return Err("AutoFilter custom predicate must contain one or two operands".into()); }
    let and = flag(node, "and", false)?;
    let first = single_custom(&node.children[0], cells, range, column)?;
    let second = single_custom(&node.children[1], cells, range, column)?;
    let (FilterCriterion::Condition { operator: first_op, value: first_value, .. }, FilterCriterion::Condition { operator: second_op, value: second_value, .. }) = (first, second) else {
        return Err("compound AutoFilter predicate requires a separate predicate adapter".into());
    };
    if and && ((first_op == "isBlank" && second_op == "isNotBlank") || (first_op == "isNotBlank" && second_op == "isBlank")) {
        return Ok(FilterCriterion::Values { selected_items: Vec::new() });
    }
    let (operator, lower, upper) = if and && first_op == "greaterThanOrEqual" && second_op == "lessThanOrEqual" { ("between", first_value, second_value) }
        else if and && second_op == "greaterThanOrEqual" && first_op == "lessThanOrEqual" { ("between", second_value, first_value) }
        else if !and && first_op == "lessThan" && second_op == "greaterThan" { ("notBetween", first_value, second_value) }
        else if !and && second_op == "lessThan" && first_op == "greaterThan" { ("notBetween", second_value, first_value) }
        else { return Err("compound AutoFilter predicate requires a separate predicate adapter".into()); };
    Ok(FilterCriterion::Condition { operator: operator.into(), value: lower, second_value: upper, case_sensitive: None })
}

pub(crate) fn standard_filter(node: &OfficeFilterNode, cells: &[CellPayload], hidden_rows: &[u32]) -> Result<AutoFilterState, String> {
    standard_filter_with_colors(node, cells, hidden_rows, &crate::sheet_filter_color::FilterColors::default())
}

pub(crate) fn standard_filter_with_colors(node: &OfficeFilterNode, cells: &[CellPayload], hidden_rows: &[u32], colors: &crate::sheet_filter_color::FilterColors) -> Result<AutoFilterState, String> {
    if !node.is_standard("autoFilter") { return Err("invalid standard AutoFilter namespace".into()); }
    let range = range(node.attributes.get("ref").ok_or_else(|| "AutoFilter is missing ref".to_owned())?).ok_or_else(|| "invalid standard AutoFilter range".to_owned())?;
    let mut columns = BTreeMap::new();
    for column in &node.children {
        if column.is_standard("sortState") { continue; }
        if !column.is_standard("filterColumn") { return Err("extended AutoFilter child requires a separate Office adapter".into()); }
        let index = column.attributes.get("colId").and_then(|value| value.parse::<u32>().ok()).and_then(|value| range.start_col.checked_add(value)).ok_or_else(|| "invalid standard AutoFilter column".to_owned())?;
        if index > range.end_col || columns.contains_key(&index) { return Err("invalid or duplicate standard AutoFilter column".into()); }
        if column.children.is_empty() { continue; }
        if column.children.len() != 1 { return Err("extended AutoFilter column requires a separate Office adapter".into()); }
        let criterion = &column.children[0];
        let criterion = if criterion.is_standard("filters") { members(criterion, cells, &range, index)? }
            else if criterion.is_standard("customFilters") { custom(criterion, cells, &range, index)? }
            else if criterion.is_standard("colorFilter") {
                match colors.decode(criterion)? {
                    office_io_common::filter_color::FilterColor::Fill(color) => FilterCriterion::Color { cell_fill_colors: Some(vec![color]), cell_text_colors: None },
                    office_io_common::filter_color::FilterColor::Text(color) => FilterCriterion::Color { cell_fill_colors: None, cell_text_colors: Some(vec![color]) },
                }
            }
            else { return Err(format!("AutoFilter {} criteria are retained in sourceAutoFilter and need their Office predicate adapter", criterion.name)); };
        columns.insert(index, criterion);
    }
    // Standard XLSX does not label manual and filtered hidden rows separately.
    // This initial cache is replaced by the application engine, while raw source
    // geometry remains available. Exact project snapshots restore each reason.
    let filtered_rows = hidden_rows.iter().copied().filter(|row| *row > range.start_row && *row <= range.end_row).collect();
    let result = AutoFilterState { range, columns, filtered_rows };
    validate(&result)?;
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn typed_snapshots_preserve_explicit_null_and_validate_columns_and_data_rows() {
        let state: AutoFilterState = serde_json::from_value(json!({ "range": { "startRow": 0, "startCol": 1, "endRow": 9, "endCol": 2 }, "columns": { "1": { "kind": "condition", "operator": "equals", "value": null } }, "filteredRows": [3] })).unwrap();
        validate(&state).unwrap();
        let FilterCriterion::Condition { value, .. } = &state.columns[&1] else { panic!("condition missing"); };
        assert_eq!(value.as_ref(), Some(&Value::Null));
        let mut invalid = state.clone(); invalid.filtered_rows.push(0);
        assert!(validate(&invalid).is_err());
        let mut invalid = state; invalid.columns.insert(3, FilterCriterion::Values { selected_items: Vec::new() });
        assert!(validate(&invalid).is_err());
    }
}

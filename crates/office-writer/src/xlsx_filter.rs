//! Canonical worksheet filters mapped to standard SpreadsheetML AutoFilter.
//! Filtering remains owned by the application engine: filteredRows is the visible
//! snapshot, while the XML criteria allow Office clients to reapply the filter.

use super::{
    MAX_COLUMN_INDEX, MAX_ROW_INDEX, RangeInput, SheetInput, cell_from_entry, column_name, esc,
    range_bounds, range_ref, value_text,
};
use crate::{ExportError, ExportWarning};
use serde::{Deserialize, Deserializer};
use serde_json::Value;
use std::collections::{BTreeMap, BTreeSet};

pub(super) fn snapshot_value(sheet: &SheetInput) -> Result<Value, ExportError> {
    let Some(filter) = &sheet.auto_filter else { return Ok(Value::Null); };
    let (start_row, start_col, end_row, end_col) = range_bounds(&filter.range).ok_or_else(|| ExportError::invalid("invalid typed AutoFilter snapshot range"))?;
    let mut columns = serde_json::Map::new();
    for (column, criterion) in &filter.columns {
        let value = match criterion {
            Criterion::Values { selected_items } => serde_json::json!({ "kind": "values", "selectedItems": selected_items }),
            Criterion::Condition { operator, value, second_value, case_sensitive } => {
                let mut criterion = serde_json::Map::new();
                criterion.insert("kind".into(), Value::String("condition".into()));
                criterion.insert("operator".into(), Value::String(operator.clone()));
                if let Some(value) = value { criterion.insert("value".into(), value.clone()); }
                if let Some(value) = second_value { criterion.insert("secondValue".into(), value.clone()); }
                criterion.insert("caseSensitive".into(), Value::Bool(*case_sensitive));
                Value::Object(criterion)
            }
        };
        columns.insert(column.to_string(), value);
    }
    let filtered_rows = filter.filtered_rows.iter().copied().collect::<BTreeSet<_>>();
    Ok(serde_json::json!({ "range": { "startRow": start_row, "startCol": start_col, "endRow": end_row, "endCol": end_col }, "columns": columns, "filteredRows": filtered_rows }))
}

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(super) struct AutoFilterInput {
    pub(super) range: RangeInput,
    #[serde(default)]
    pub(super) columns: BTreeMap<u32, Criterion>,
    #[serde(default)]
    pub(super) filtered_rows: Vec<u32>,
}

#[derive(Debug, Deserialize, Clone)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub(super) enum Criterion {
    Values {
        #[serde(rename = "selectedItems")]
        selected_items: Vec<Value>,
    },
    Condition {
        operator: String,
        #[serde(default, deserialize_with = "present_value")]
        value: Option<Value>,
        #[serde(default, rename = "secondValue", deserialize_with = "present_value")]
        second_value: Option<Value>,
        #[serde(default, rename = "caseSensitive")]
        case_sensitive: bool,
    },
}

// Preserve an explicit null separately from a missing comparison operand.
fn present_value<'de, D: Deserializer<'de>>(deserializer: D) -> Result<Option<Value>, D::Error> {
    Value::deserialize(deserializer).map(Some)
}

fn scalar(value: &Value) -> bool {
    value.is_null() || value.is_string() || value.is_number() || value.is_boolean()
}

pub(super) fn validate(sheet: &SheetInput) -> Result<(), ExportError> {
    let Some(filter) = &sheet.auto_filter else {
        return Ok(());
    };
    validate_filter(filter, sheet.row_count, sheet.column_count, true)
}

pub(super) fn validate_filter(filter: &AutoFilterInput, row_count: Option<u32>, column_count: Option<u32>, require_data: bool) -> Result<(), ExportError> {
    let (first_row, first_col, last_row, last_col) = range_bounds(&filter.range)
        .ok_or_else(|| ExportError::invalid("autoFilter must have a complete worksheet range"))?;
    if last_row < first_row || require_data && last_row == first_row
        || last_col < first_col
        || last_row > MAX_ROW_INDEX
        || last_col > MAX_COLUMN_INDEX
        || row_count.is_some_and(|count| last_row >= count)
        || column_count.is_some_and(|count| last_col >= count)
    {
        return Err(ExportError::invalid(
            "autoFilter range must include a header and data within worksheet bounds",
        ));
    }
    if filter
        .filtered_rows
        .iter()
        .any(|row| *row <= first_row || *row > last_row)
    {
        return Err(ExportError::invalid(
            "autoFilter filteredRows must be data rows inside its range",
        ));
    }
    for (column, criterion) in &filter.columns {
        if *column < first_col || *column > last_col {
            return Err(ExportError::invalid(
                "autoFilter column must be inside its range",
            ));
        }
        match criterion {
            Criterion::Values { selected_items } => {
                if selected_items.iter().any(|value| !scalar(value)) {
                    return Err(ExportError::invalid(
                        "autoFilter values must be JSON scalars",
                    ));
                }
            }
            Criterion::Condition {
                operator,
                value,
                second_value,
                ..
            } => {
                let no_value = matches!(operator.as_str(), "isBlank" | "isNotBlank" | "isError");
                let supported = no_value
                    || matches!(
                        operator.as_str(),
                        "equals"
                            | "doesNotEqual"
                            | "greaterThan"
                            | "greaterThanOrEqual"
                            | "lessThan"
                            | "lessThanOrEqual"
                            | "between"
                            | "notBetween"
                            | "contains"
                            | "doesNotContain"
                            | "beginsWith"
                            | "endsWith"
                    );
                if !supported
                    || (!no_value && value.is_none())
                    || (matches!(operator.as_str(), "between" | "notBetween")
                        && second_value.is_none())
                    || value.as_ref().is_some_and(|value| !scalar(value))
                    || second_value.as_ref().is_some_and(|value| !scalar(value))
                {
                    return Err(ExportError::invalid(
                        "autoFilter condition has an unsupported operator or invalid operands",
                    ));
                }
            }
        }
    }
    Ok(())
}

pub(super) fn hidden_rows(sheet: &SheetInput) -> BTreeSet<u32> {
    sheet
        .hidden_rows
        .iter()
        .copied()
        .chain(
            sheet
                .auto_filter
                .iter()
                .flat_map(|filter| filter.filtered_rows.iter().copied()),
        )
        .collect()
}

fn warning(
    warnings: &mut Vec<ExportWarning>,
    code: &'static str,
    sheet: &SheetInput,
    column: u32,
    detail: &str,
) {
    warnings.push(ExportWarning::new(
        code,
        format!(
            "AutoFilter on {}!{}: {detail}",
            sheet.name,
            column_name(column)
        ),
    ));
}

fn custom(operator: &str, value: &str) -> String {
    format!(
        "<customFilter operator=\"{operator}\" val=\"{}\"/>",
        esc(value)
    )
}

fn none_match() -> String {
    // A valid contradictory pair, including for blanks. An empty <filters/> may
    // be interpreted as 'show all' by Office clients instead of the intended [].
    format!(
        "<customFilters and=\"1\">{}{}</customFilters>",
        custom("equal", " "),
        custom("notEqual", " ")
    )
}

fn values_xml(values: &[Value]) -> String {
    if values.is_empty() {
        return none_match();
    }
    let blank = values
        .iter()
        .any(|value| value.is_null() || value.as_str() == Some(""));
    let unique: BTreeSet<String> = values
        .iter()
        .filter(|value| !value.is_null() && value.as_str() != Some(""))
        .filter_map(value_text)
        .collect();
    let mut out = format!("<filters blank=\"{}\">", u8::from(blank));
    for value in unique {
        out.push_str(&format!("<filter val=\"{}\"/>", esc(&value)));
    }
    out.push_str("</filters>");
    out
}

fn wildcard_literal(value: &str) -> String {
    let mut result = String::new();
    for character in value.chars() {
        if matches!(character, '*' | '?' | '~') {
            result.push('~');
        }
        result.push(character);
    }
    result
}

fn operand(value: Option<&Value>) -> String {
    match value {
        Some(Value::Null) | None => String::new(),
        Some(value) => value_text(value).unwrap_or_default(),
    }
}

fn scalar_type(value: &Value) -> u8 {
    if value.is_number() {
        1
    } else if value.is_boolean() {
        2
    } else if value.is_string() {
        4
    } else {
        0
    }
}

fn column_values(sheet: &SheetInput, filter: &AutoFilterInput, column: u32) -> Vec<Value> {
    let Some((first, _, last, _)) = range_bounds(&filter.range) else {
        return Vec::new();
    };
    sheet
        .cells
        .iter()
        .filter_map(|(key, raw)| {
            let cell = cell_from_entry(key, raw)?;
            if cell.col != Some(column)
                || cell.row.is_none_or(|row| row <= first || row > last)
                || matches!(cell.data_type.as_deref(), Some("e" | "error"))
            {
                return None;
            }
            cell.value
        })
        .collect()
}

fn condition_xml(
    sheet: &SheetInput,
    column: u32,
    filter: &AutoFilterInput,
    criterion: &Criterion,
    warnings: &mut Vec<ExportWarning>,
) -> Result<String, ExportError> {
    let Criterion::Condition {
        operator,
        value,
        second_value,
        case_sensitive,
    } = criterion
    else {
        return Err(ExportError::invalid("expected an AutoFilter condition"));
    };
    if *case_sensitive {
        warning(
            warnings,
            "AUTOFILTER_CASE_SENSITIVE",
            sheet,
            column,
            "standard XLSX filters do not preserve case-sensitive comparison; current hidden rows are retained, reapplying in Office may differ",
        );
    }
    if operator == "isBlank" {
        return Ok("<filters blank=\"1\"/>".into());
    }
    if operator == "isNotBlank" {
        return Ok(format!(
            "<customFilters>{}</customFilters>",
            custom("notEqual", " ")
        ));
    }
    if operator == "isError" {
        let (first_row, _, last_row, _) = range_bounds(&filter.range)
            .ok_or_else(|| ExportError::invalid("invalid AutoFilter range"))?;
        let mut errors = BTreeSet::new();
        for (key, raw) in &sheet.cells {
            let Some(cell) = cell_from_entry(key, raw) else {
                continue;
            };
            if cell.col != Some(column)
                || cell
                    .row
                    .is_none_or(|row| row <= first_row || row > last_row)
            {
                continue;
            }
            let metadata_error = cell
                .metadata
                .as_ref()
                .and_then(|metadata| metadata.get("error"))
                .and_then(Value::as_str);
            if let Some(error) = metadata_error {
                errors.insert(error.to_owned());
            } else if matches!(cell.data_type.as_deref(), Some("e" | "error")) {
                if let Some(error) = cell.value.as_ref().and_then(value_text) {
                    errors.insert(error);
                }
            }
        }
        warning(
            warnings,
            "AUTOFILTER_ERROR_SNAPSHOT",
            sheet,
            column,
            "isError was exported as the current error-code member list; standard AutoFilter has no dynamic error-type predicate, so reapplying after errors change may differ",
        );
        return Ok(values_xml(
            &errors.into_iter().map(Value::String).collect::<Vec<_>>(),
        ));
    }
    let first = operand(value.as_ref());
    let second = operand(second_value.as_ref());
    let first_literal = wildcard_literal(&first);
    let second_literal = wildcard_literal(&second);
    if value.as_ref().is_some_and(|expected| {
        !expected.is_null()
            && column_values(sheet, filter, column).iter().any(|actual| {
                !actual.is_null()
                    && actual.as_str() != Some("")
                    && scalar_type(actual) != scalar_type(expected)
            })
    }) {
        warning(
            warnings,
            "AUTOFILTER_MIXED_TYPES",
            sheet,
            column,
            "Office may coerce mixed scalar types differently from canonical typed comparison when the filter is reapplied",
        );
    }
    if matches!(operator.as_str(), "equals" | "doesNotEqual")
        && matches!(value.as_ref().and_then(Value::as_str), Some("" | " "))
    {
        warning(
            warnings,
            "AUTOFILTER_BLANK_TYPES",
            sheet,
            column,
            "Office custom blank operands cannot preserve every distinction between empty, whitespace and missing cells",
        );
    }
    let (and, conditions) = match operator.as_str() {
        "equals" => (None, custom("equal", &first_literal)),
        "doesNotEqual" => (None, custom("notEqual", &first_literal)),
        "greaterThan" | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual" => {
            (None, custom(operator, &first_literal))
        }
        "contains" => (None, custom("equal", &format!("*{first_literal}*"))),
        "doesNotContain" => (None, custom("notEqual", &format!("*{first_literal}*"))),
        "beginsWith" => (None, custom("equal", &format!("{first_literal}*"))),
        "endsWith" => (None, custom("equal", &format!("*{first_literal}"))),
        "between" => (
            Some(true),
            format!(
                "{}{}",
                custom("greaterThanOrEqual", &first_literal),
                custom("lessThanOrEqual", &second_literal)
            ),
        ),
        "notBetween" => (
            Some(false),
            format!(
                "{}{}",
                custom("lessThan", &first_literal),
                custom("greaterThan", &second_literal)
            ),
        ),
        _ => return Err(ExportError::invalid("unsupported AutoFilter condition")),
    };
    if value.as_ref().is_some_and(Value::is_null)
        || second_value.as_ref().is_some_and(Value::is_null)
    {
        warning(
            warnings,
            "AUTOFILTER_NULL_COMPARISON",
            sheet,
            column,
            "null comparison uses an empty Office operand; canonical typed/null comparison may differ when the filter is reapplied",
        );
    }
    if value.as_ref().is_some_and(Value::is_string)
        && matches!(
            operator.as_str(),
            "greaterThan"
                | "greaterThanOrEqual"
                | "lessThan"
                | "lessThanOrEqual"
                | "between"
                | "notBetween"
        )
    {
        warning(
            warnings,
            "AUTOFILTER_TEXT_COLLATION",
            sheet,
            column,
            "Office text ordering depends on its locale and may differ from the application numeric zh-CN collation",
        );
    }
    Ok(format!(
        "<customFilters{}>{conditions}</customFilters>",
        and.map(|value| format!(" and=\"{}\"", u8::from(value)))
            .unwrap_or_default()
    ))
}

pub(super) fn write_xml(
    sheet: &SheetInput,
    warnings: &mut Vec<ExportWarning>,
) -> Result<String, ExportError> {
    let Some(filter) = &sheet.auto_filter else {
        return Ok(String::new());
    };
    write_filter_xml(sheet, filter, warnings)
}

pub(super) fn write_filter_xml(sheet: &SheetInput, filter: &AutoFilterInput, warnings: &mut Vec<ExportWarning>) -> Result<String, ExportError> {
    let reference =
        range_ref(&filter.range).ok_or_else(|| ExportError::invalid("invalid AutoFilter range"))?;
    let (_, first_col, _, _) = range_bounds(&filter.range)
        .ok_or_else(|| ExportError::invalid("invalid AutoFilter range"))?;
    let mut out = format!("<autoFilter ref=\"{reference}\">");
    for (column, criterion) in &filter.columns {
        out.push_str(&format!("<filterColumn colId=\"{}\">", column - first_col));
        match criterion {
            Criterion::Values { selected_items } => {
                let null = selected_items.iter().any(Value::is_null);
                let empty = selected_items
                    .iter()
                    .any(|value| value.as_str() == Some(""));
                if null != empty {
                    warning(
                        warnings,
                        "AUTOFILTER_BLANK_TYPES",
                        sheet,
                        *column,
                        "XLSX blank filtering does not distinguish missing cells from empty text; current hidden rows are retained",
                    );
                }
                let mut types = BTreeMap::<String, u8>::new();
                for value in selected_items {
                    let mask = scalar_type(value);
                    if let Some(text) = value_text(value) {
                        *types.entry(text).or_default() |= mask;
                    }
                }
                for value in column_values(sheet, filter, *column) {
                    if let Some(text) = value_text(&value) {
                        if let Some(mask) = types.get_mut(&text) {
                            *mask |= scalar_type(&value);
                        }
                    }
                }
                if types.values().any(|mask| mask.count_ones() > 1) {
                    warning(
                        warnings,
                        "AUTOFILTER_VALUE_TYPES",
                        sheet,
                        *column,
                        "XLSX member values cannot preserve distinctions between scalar types with the same text",
                    );
                }
                out.push_str(&values_xml(selected_items));
            }
            Criterion::Condition { .. } => {
                out.push_str(&condition_xml(sheet, *column, filter, criterion, warnings)?)
            }
        }
        out.push_str("</filterColumn>");
    }
    out.push_str("</autoFilter>");
    Ok(out)
}

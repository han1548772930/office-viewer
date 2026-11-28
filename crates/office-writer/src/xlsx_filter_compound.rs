//! Native customFilters contain at most two predicates. Keep larger local
//! expressions in the typed extension and report the native reapply boundary.
use std::collections::{BTreeMap, BTreeSet};
use serde_json::Value;
use super::{AutoFilterInput, ConditionProjection, Criterion, SheetInput, cell_from_entry,
    condition_projection, custom, range_bounds, values_xml, warning};
use crate::{ExportError, ExportWarning};

pub(super) fn write_xml(sheet: &SheetInput, column: u32, filter: &AutoFilterInput, join: &str,
    conditions: &[Criterion], warnings: &mut Vec<ExportWarning>) -> Result<String, ExportError> {
    if !matches!(join, "and" | "or") || !(1..=2).contains(&conditions.len()) {
        return Err(ExportError::invalid("invalid AutoFilter compound condition"));
    }
    if conditions.len() == 1 { return Ok(condition_projection(sheet, column, filter, &conditions[0], warnings)?.xml()); }
    let mut predicates = String::new();
    let mut native = true;
    for condition in conditions {
        let projection = condition_projection(sheet, column, filter, condition, warnings)?;
        if matches!(condition, Criterion::Condition { operator, .. } if operator == "isBlank") {
            predicates.push_str(&custom("equal", " "));
        } else if let ConditionProjection::Custom { conditions, count: 1, .. } = projection {
            predicates.push_str(&conditions);
        } else { native = false; }
    }
    if native { return Ok(format!("<customFilters and=\"{}\">{predicates}</customFilters>", u8::from(join == "and"))); }
    warning(warnings, "AUTOFILTER_COMPOUND_SNAPSHOT", sheet, column,
        "this compound contains more than two native predicates or a dynamic error-type rule; the complete expression is retained in the project extension and current visible members are written for Office, whose later reapply may differ");
    visible_members(sheet, column, filter)
}

fn visible_members(sheet: &SheetInput, column: u32, filter: &AutoFilterInput) -> Result<String, ExportError> {
    let (first, _, last, _) = range_bounds(&filter.range).ok_or_else(|| ExportError::invalid("invalid AutoFilter snapshot range"))?;
    let hidden = filter.filtered_rows.iter().copied().filter(|row| *row > first && *row <= last).collect::<BTreeSet<_>>();
    let mut rows = BTreeMap::new();
    for (key, raw) in &sheet.cells {
        let Some(cell) = cell_from_entry(key, raw) else { continue; };
        if cell.col != Some(column) { continue; }
        let Some(row) = cell.row.filter(|row| *row > first && *row <= last && !hidden.contains(row)) else { continue; };
        let value = cell.metadata.as_ref().and_then(|metadata| metadata.get("error")).cloned().or(cell.value).unwrap_or(Value::Null);
        rows.insert(row, value);
    }
    let visible_count = (last - first) as usize - hidden.len();
    let implicit_blank = rows.len() < visible_count;
    let mut values = rows.into_values().collect::<Vec<_>>();
    if implicit_blank { values.push(Value::Null); }
    Ok(values_xml(&values))
}

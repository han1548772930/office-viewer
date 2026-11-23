//! Lossless projection of range-rule formula origins into standard SpreadsheetML.
//! A safe sqref origin keeps ordinary A1 formulas. A rule with no such cell uses
//! reference-returning INDEX expressions, preserving one rule's full statistical scope.

use super::{ConditionalFormatInput, RangeInput, range_bounds};
use crate::ExportError;
use serde::{Deserialize, Serialize};
use serde_json::Value;

const MAX_ROW: u32 = 1_048_575;
const MAX_COLUMN: u32 = 16_383;

#[derive(Debug, Deserialize, Serialize, Clone, Copy, PartialEq, Eq)]
pub(super) struct FormulaAnchor {
    pub(super) row: u32,
    pub(super) col: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Bounds { start_row: u32, start_col: u32, end_row: u32, end_col: u32 }

impl Bounds {
    fn origin(self) -> FormulaAnchor { FormulaAnchor { row: self.start_row, col: self.start_col } }
    fn contains(self, origin: FormulaAnchor) -> bool {
        origin.row >= self.start_row && origin.row <= self.end_row && origin.col >= self.start_col && origin.col <= self.end_col
    }
    fn intersect(self, other: Self) -> Option<Self> {
        let result = Self { start_row: self.start_row.max(other.start_row), start_col: self.start_col.max(other.start_col),
            end_row: self.end_row.min(other.end_row), end_col: self.end_col.min(other.end_col) };
        (result.start_row <= result.end_row && result.start_col <= result.end_col).then_some(result)
    }
    fn subtract(self, removed: Self) -> Vec<Self> {
        let Some(cut) = self.intersect(removed) else { return vec![self] };
        let mut parts = Vec::new();
        if self.start_row < cut.start_row { parts.push(Self { end_row: cut.start_row - 1, ..self }); }
        if cut.end_row < self.end_row { parts.push(Self { start_row: cut.end_row + 1, ..self }); }
        if self.start_col < cut.start_col { parts.push(Self { start_row: cut.start_row, end_row: cut.end_row, end_col: cut.start_col - 1, ..self }); }
        if cut.end_col < self.end_col { parts.push(Self { start_row: cut.start_row, end_row: cut.end_row, start_col: cut.end_col + 1, ..self }); }
        parts
    }
    fn input(self) -> RangeInput {
        RangeInput { start_row: Some(self.start_row), start_col: Some(self.start_col), end_row: Some(self.end_row), end_col: Some(self.end_col),
            first_row: None, first_col: None, last_row: None, last_col: None }
    }
}

fn checked_range(range: &RangeInput) -> Result<Bounds, ExportError> {
    let (start_row, start_col, end_row, end_col) = range_bounds(range).ok_or_else(|| ExportError::invalid("conditional format range has missing coordinates"))?;
    if start_row > end_row || start_col > end_col || end_row > MAX_ROW || end_col > MAX_COLUMN {
        return Err(ExportError::invalid("conditional format range is outside the worksheet"));
    }
    Ok(Bounds { start_row, start_col, end_row, end_col })
}

fn formula_operands(rule: &ConditionalFormatInput) -> Vec<&str> {
    if rule.rule_type == "office" { return Vec::new() }
    let expression = rule.rule_type == "expression" || rule.sub_type.as_deref() == Some("formula");
    let numeric = rule.rule_type == "cellIs" || rule.rule_type == "highlightCell" && rule.sub_type.as_deref().unwrap_or("number") == "number";
    let mut values = Vec::new();
    if expression || numeric || has_standard_predicate(rule) {
        values.push(rule.formula1.as_str());
        if let Some(second) = rule.formula2.as_deref() { values.push(second); }
    }
    if expression && let Some(value) = rule.value.as_ref().and_then(Value::as_str) { values.push(value); }
    if let Some(config) = rule.config.as_ref() {
        for key in ["min", "max"] {
            if let Some(value) = config.get(key).and_then(formula_threshold) { values.push(value); }
        }
        for key in ["thresholds", "icons"] {
            for entry in config.get(key).and_then(Value::as_array).into_iter().flatten() {
                if let Some(value) = entry.get("value").and_then(formula_threshold) { values.push(value); }
            }
        }
    }
    values
}

fn has_standard_predicate(rule: &ConditionalFormatInput) -> bool {
    if rule.rule_type == "office" || rule.formula1.trim().is_empty() { return false }
    let text = matches!(rule.rule_type.as_str(), "containsText" | "notContainsText" | "beginsWith" | "endsWith")
        || rule.sub_type.as_deref() == Some("text");
    if !text { return true }
    if matches!(rule.operator.as_str(), "containsBlanks" | "notContainsBlanks" | "containsErrors" | "notContainsErrors") { return true }
    rule.value.as_ref().and_then(super::json_scalar_string).is_some_and(|literal| literal != rule.formula1)
}

fn formula_threshold(value: &Value) -> Option<&str> {
    (value.get("type").and_then(Value::as_str) == Some("formula")).then(|| value.get("value").and_then(Value::as_str)).flatten()
}

pub(super) fn project_rule(rule: &ConditionalFormatInput) -> Result<ConditionalFormatInput, ExportError> {
    let ranges = rule.ranges.iter().map(checked_range).collect::<Result<Vec<_>, _>>()?;
    let first = ranges.first().ok_or_else(|| ExportError::invalid("conditional format has no ranges"))?;
    let origin = rule.formula_anchor.unwrap_or_else(|| ranges.iter().min_by_key(|range| (range.start_row, range.start_col)).unwrap().origin());
    if origin.row > MAX_ROW || origin.col > MAX_COLUMN { return Err(ExportError::invalid("conditional formula origin is outside the worksheet")); }
    if rule.rule_type == "office" { return Ok(rule.clone()) }
    let mut safe = Bounds { start_row: 0, start_col: 0, end_row: MAX_ROW, end_col: MAX_COLUMN };
    for formula in formula_operands(rule) {
        for reference in references(formula) {
            if !reference.absolute_row {
                let delta = i64::from(origin.row) - i64::from(reference.row);
                safe.start_row = safe.start_row.max(delta.max(0) as u32);
                safe.end_row = safe.end_row.min((i64::from(MAX_ROW) + delta).min(i64::from(MAX_ROW)) as u32);
            }
            if !reference.absolute_col {
                let delta = i64::from(origin.col) - i64::from(reference.col);
                safe.start_col = safe.start_col.max(delta.max(0) as u32);
                safe.end_col = safe.end_col.min((i64::from(MAX_COLUMN) + delta).min(i64::from(MAX_COLUMN)) as u32);
            }
        }
    }
    let mut result = rule.clone();
    let target = if safe.contains(first.origin()) { first.origin() }
    else if let Some(anchor_range) = ranges.iter().filter_map(|range| range.intersect(safe)).min_by_key(|range| (range.start_row, range.start_col)) {
        // Reorder disjoint sqref pieces, not rules: aggregations see exactly the same cell union.
        result.ranges = std::iter::once(anchor_range).chain(ranges.iter().flat_map(|range| range.subtract(anchor_range))).map(Bounds::input).collect();
        anchor_range.origin()
    } else { first.origin() };
    let expression = rule.rule_type == "expression" || rule.sub_type.as_deref() == Some("formula");
    let numeric = rule.rule_type == "cellIs" || rule.rule_type == "highlightCell" && rule.sub_type.as_deref().unwrap_or("number") == "number";
    if expression || numeric || has_standard_predicate(rule) {
        result.formula1 = project_formula(&rule.formula1, origin, target);
        result.formula2 = rule.formula2.as_ref().map(|value| project_formula(value, origin, target));
    }
    if expression && let Some(value) = rule.value.as_ref().and_then(Value::as_str) {
        let formula = project_formula(value, origin, target);
        if result.formula1.trim().is_empty() { result.formula1 = formula.clone(); }
        result.value = Some(Value::String(formula));
    }
    if let Some(config) = result.config.as_mut() {
        for key in ["min", "max"] {
            if let Some(value) = config.get_mut(key) { project_threshold(value, origin, target); }
        }
        for key in ["thresholds", "icons"] {
            for entry in config.get_mut(key).and_then(Value::as_array_mut).into_iter().flatten() {
                if let Some(value) = entry.get_mut("value") { project_threshold(value, origin, target); }
            }
        }
    }
    result.formula_anchor = Some(target);
    Ok(result)
}

fn project_threshold(value: &mut Value, origin: FormulaAnchor, target: FormulaAnchor) {
    let Some(formula) = formula_threshold(value) else { return };
    let projected = project_formula(formula, origin, target).trim().trim_start_matches('=').to_owned();
    value["value"] = Value::String(projected);
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ReferenceKind { Cell, WholeRow, WholeColumn }

#[derive(Debug, Clone)]
struct Reference {
    start: usize, end: usize, prefix: String, kind: ReferenceKind,
    row: u32, col: u32, absolute_row: bool, absolute_col: bool,
}

fn shifted(value: u32, absolute: bool, before: u32, after: u32, maximum: u32) -> Option<u32> {
    let next = i64::from(value) + if absolute { 0 } else { i64::from(after) - i64::from(before) };
    (next >= 0 && next <= i64::from(maximum)).then_some(next as u32)
}

pub(super) fn project_formula(formula: &str, origin: FormulaAnchor, target: FormulaAnchor) -> String {
    let spans = references(formula);
    let ordinary = spans.iter().all(|span| shifted(span.row, span.absolute_row, origin.row, target.row, MAX_ROW).is_some()
        && shifted(span.col, span.absolute_col, origin.col, target.col, MAX_COLUMN).is_some());
    let mut output = String::new();
    let mut cursor = 0;
    for (index, span) in spans.iter().enumerate() {
        output.push_str(&formula[cursor..span.start]);
        if ordinary {
            output.push_str(&span.prefix);
            output.push_str(&reference_address(span,
                shifted(span.row, span.absolute_row, origin.row, target.row, MAX_ROW).unwrap(),
                shifted(span.col, span.absolute_col, origin.col, target.col, MAX_COLUMN).unwrap()));
        } else {
            let prefix = if span.prefix.is_empty() && index > 0 && formula[spans[index - 1].end..span.start].trim() == ":" { &spans[index - 1].prefix } else { &span.prefix };
            if span.absolute_row && span.absolute_col { output.push_str(&format!("{}{}", prefix, reference_address(span, span.row, span.col))); }
            else {
                output.push_str(&format!("INDEX({prefix}$A$1:$XFD$1048576,{},{})",
                    if span.kind == ReferenceKind::WholeColumn { "0".into() } else { index_axis(span.row, span.absolute_row, origin.row, true) },
                    if span.kind == ReferenceKind::WholeRow { "0".into() } else { index_axis(span.col, span.absolute_col, origin.col, false) }));
            }
        }
        cursor = span.end;
    }
    output.push_str(&formula[cursor..]);
    output
}

fn index_axis(value: u32, absolute: bool, origin: u32, row: bool) -> String {
    if absolute { return (value + 1).to_string() }
    let delta = i64::from(value) - i64::from(origin);
    let function = if row { "ROW()" } else { "COLUMN()" };
    let coordinate = if delta == 0 { function.to_owned() } else if delta > 0 { format!("{function}+{delta}") } else { format!("{function}{delta}") };
    // INDEX row/column zero means a whole axis. Map lower overflow to a positive out-of-range index instead.
    if delta < 0 { format!("IF({coordinate}<1,{},{coordinate})", if row { MAX_ROW + 2 } else { MAX_COLUMN + 2 }) } else { coordinate }
}

fn address(row: u32, col: u32, absolute_row: bool, absolute_col: bool) -> String {
    format!("{}{}{}{}", if absolute_col { "$" } else { "" }, super::column_name(col), if absolute_row { "$" } else { "" }, row + 1)
}

fn reference_address(reference: &Reference, row: u32, col: u32) -> String {
    match reference.kind {
        ReferenceKind::Cell => address(row, col, reference.absolute_row, reference.absolute_col),
        ReferenceKind::WholeRow => format!("{}{}", if reference.absolute_row { "$" } else { "" }, row + 1),
        ReferenceKind::WholeColumn => format!("{}{}", if reference.absolute_col { "$" } else { "" }, super::column_name(col)),
    }
}

fn identifier(character: char) -> bool { character.is_alphanumeric() || matches!(character, '_' | '.' | '\\') }
fn qualifier_end(source: &str, start: usize) -> Option<usize> {
    let first = source[start..].chars().next()?;
    if !first.is_alphabetic() && !matches!(first, '_' | '\\') { return None }
    let mut cursor = start;
    while let Some(character) = source[cursor..].chars().next() {
        if !identifier(character) { break }
        cursor += character.len_utf8();
    }
    (source.as_bytes().get(cursor) == Some(&b'!')).then_some(cursor + 1)
}

fn read_address(source: &str, start: usize) -> Option<Reference> {
    let mut cursor = start;
    let bytes = source.as_bytes();
    let absolute_first = bytes.get(cursor) == Some(&b'$');
    if absolute_first { cursor += 1; }
    let letters = cursor;
    while bytes.get(cursor).is_some_and(u8::is_ascii_alphabetic) { cursor += 1; }
    if cursor - letters > 3 { return None }
    let has_column = cursor > letters;
    let col = if has_column {
        source[letters..cursor].bytes().fold(0u32, |value, letter| value * 26 + u32::from(letter.to_ascii_uppercase() - b'A') + 1).checked_sub(1)?
    } else { 0 };
    let absolute_row = if has_column { bytes.get(cursor) == Some(&b'$') } else { absolute_first };
    if has_column && absolute_row { cursor += 1; }
    let digits = cursor;
    while bytes.get(cursor).is_some_and(u8::is_ascii_digit) { cursor += 1; }
    let has_row = cursor > digits;
    if !has_row && (!has_column || absolute_row) { return None }
    if source[cursor..].chars().next().is_some_and(identifier) || source[cursor..].trim_start().starts_with('(') { return None }
    let row = if has_row { source[digits..cursor].parse::<u32>().ok()?.checked_sub(1)? } else { 0 };
    if row > MAX_ROW || col > MAX_COLUMN { return None }
    Some(Reference { start, end: cursor, prefix: String::new(),
        kind: if !has_column { ReferenceKind::WholeRow } else if !has_row { ReferenceKind::WholeColumn } else { ReferenceKind::Cell },
        row, col, absolute_row: !has_row || absolute_row, absolute_col: !has_column || absolute_first })
}

fn references(source: &str) -> Vec<Reference> {
    let mut result = Vec::new();
    let mut cursor = 0;
    while cursor < source.len() {
        let character = source[cursor..].chars().next().unwrap();
        let start = cursor;
        if character == '"' {
            cursor += 1;
            while cursor < source.len() {
                if source.as_bytes()[cursor] == b'"' {
                    cursor += 1;
                    if source.as_bytes().get(cursor) == Some(&b'"') { cursor += 1; } else { break }
                } else { cursor += source[cursor..].chars().next().unwrap().len_utf8(); }
            }
            continue;
        }
        if character == '\'' {
            cursor += 1;
            while cursor < source.len() {
                if source.as_bytes()[cursor] == b'\'' {
                    cursor += 1;
                    if source.as_bytes().get(cursor) == Some(&b'\'') { cursor += 1; continue }
                    if source.as_bytes().get(cursor) == Some(&b'!') {
                        if let Some(mut reference) = read_address(source, cursor + 1) {
                            reference.start = start; reference.prefix = source[start..cursor + 1].to_owned(); cursor = reference.end; result.push(reference);
                        } else { cursor += 1; }
                    }
                    break;
                }
                cursor += source[cursor..].chars().next().unwrap().len_utf8();
            }
            continue;
        }
        if character == '[' {
            let mut nesting = 1;
            cursor += 1;
            while cursor < source.len() && nesting > 0 {
                let next = source[cursor..].chars().next().unwrap();
                if next == '[' { nesting += 1; } else if next == ']' { nesting -= 1; }
                cursor += next.len_utf8();
            }
            if !source[..start].chars().next_back().is_some_and(identifier) {
                if let Some(address_start) = qualifier_end(source, cursor) {
                    if let Some(mut reference) = read_address(source, address_start) {
                        reference.start = start; reference.prefix = source[start..address_start].to_owned(); cursor = reference.end; result.push(reference);
                    }
                }
            }
            continue;
        }
        if source[..cursor].chars().next_back().is_some_and(identifier) { cursor += character.len_utf8(); continue }
        let address_start = qualifier_end(source, cursor).unwrap_or(cursor);
        if let Some(mut reference) = read_address(source, address_start) {
            reference.start = cursor; reference.prefix = source[cursor..address_start].to_owned(); cursor = reference.end; result.push(reference);
        } else { cursor += character.len_utf8(); }
    }
    // A bare name or numeric literal is not an axis reference. Both endpoints
    // must describe the same whole axis and be joined by the range operator.
    let axis_ranges = result.iter().enumerate().map(|(index, reference)| {
        reference.kind == ReferenceKind::Cell
            || index > 0 && result[index - 1].kind == reference.kind && source[result[index - 1].end..reference.start].trim() == ":"
            || index + 1 < result.len() && result[index + 1].kind == reference.kind && source[reference.end..result[index + 1].start].trim() == ":"
    }).collect::<Vec<_>>();
    result.into_iter().zip(axis_ranges).filter_map(|(reference, keep)| keep.then_some(reference)).collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn shared_projection_fixtures_preserve_literals_and_relative_reference_boundaries() {
        let cases: Vec<Value> = serde_json::from_str(include_str!("../../../tests/fixtures/conditional-formula-projection.json")).unwrap();
        for case in cases {
            let origin = serde_json::from_value(case["origin"].clone()).unwrap();
            let target = serde_json::from_value(case["target"].clone()).unwrap();
            assert_eq!(project_formula(case["formula"].as_str().unwrap(), origin, target), case["projected"].as_str().unwrap(), "{}", case["name"]);
        }
    }

    fn rule(ranges: Value, formula: &str) -> ConditionalFormatInput {
        serde_json::from_value(json!({ "id": "formula", "type": "expression", "ranges": ranges, "formulaAnchor": { "row": 1, "col": 0 }, "formula1": formula, "priority": 1 })).unwrap()
    }

    #[test]
    fn earlier_target_retains_one_rule_and_reorders_the_same_coverage_to_a_safe_origin() {
        let original = rule(json!([{ "startRow": 0, "startCol": 0, "endRow": 3, "endCol": 0 }]), "=A1>5");
        let projected = project_rule(&original).unwrap();
        assert_eq!(projected.formula1, "=A1>5");
        assert_eq!(projected.formula_anchor, Some(FormulaAnchor { row: 1, col: 0 }));
        assert_eq!(projected.ranges.iter().map(|range| checked_range(range).unwrap()).collect::<Vec<_>>(), vec![
            Bounds { start_row: 1, start_col: 0, end_row: 3, end_col: 0 }, Bounds { start_row: 0, start_col: 0, end_row: 0, end_col: 0 }
        ]);
        assert_eq!(original.ranges.len(), 1);
    }

    #[test]
    fn no_common_valid_cell_uses_standard_per_cell_projection_without_splitting_rule_scope() {
        let original = rule(json!([{ "startRow": 0, "startCol": 0, "endRow": 0, "endCol": 0 }]), "=IFERROR(A1,0)+IFERROR(A1048576,0)");
        let projected = project_rule(&original).unwrap();
        assert_eq!(projected.ranges.len(), 1);
        assert!(projected.formula1.contains("IF(ROW()-1<1,1048577,ROW()-1)"));
        assert!(projected.formula1.contains("ROW()+1048574"));
    }

    #[test]
    fn all_visual_formula_thresholds_share_the_same_lossless_origin_and_scope() {
        let mut original = rule(json!([{ "startRow": 0, "startCol": 0, "endRow": 3, "endCol": 0 }]), "");
        original.rule_type = "iconSet".into();
        original.config = Some(json!({ "min": { "type": "formula", "value": "=A1" }, "max": { "type": "num", "value": 100 },
            "thresholds": [{ "value": { "type": "formula", "value": "=$B$1+A1" }, "color": "#123456" }],
            "icons": [{ "value": { "type": "formula", "value": "=A1" }, "iconType": "3Arrows", "iconId": "0" }], "custom": { "keep": "A1" } }));
        let projected = project_rule(&original).unwrap();
        assert_eq!(projected.config.as_ref().unwrap()["min"]["value"], "A1");
        assert_eq!(projected.config.as_ref().unwrap()["thresholds"][0]["value"]["value"], "$B$1+A1");
        assert_eq!(projected.config.as_ref().unwrap()["icons"][0]["value"]["value"], "A1");
        assert_eq!(projected.config.as_ref().unwrap()["custom"]["keep"], "A1");
        assert_eq!(projected.ranges.len(), 2);
    }

    #[test]
    fn invalid_origins_and_ranges_reject_before_rule_projection() {
        let mut invalid = rule(json!([{ "startRow": 0, "startCol": 0, "endRow": 1, "endCol": 0 }]), "=A1");
        invalid.formula_anchor = Some(FormulaAnchor { row: MAX_ROW + 1, col: 0 });
        assert_eq!(project_rule(&invalid).unwrap_err().code, "INVALID_MODEL");
        invalid.formula_anchor = None;
        invalid.ranges[0].end_row = Some(MAX_ROW + 1);
        assert_eq!(project_rule(&invalid).unwrap_err().code, "INVALID_MODEL");
    }

    #[test]
    fn whole_axis_references_constrain_only_their_selected_axis() {
        let mut original = rule(json!([{ "startRow": 0, "startCol": 0, "endRow": 3, "endCol": 0 }]), "=SUM(1:2)>0");
        original.formula_anchor = Some(FormulaAnchor { row: 1, col: 100 });
        let projected = project_rule(&original).unwrap();
        assert_eq!(projected.formula_anchor, Some(FormulaAnchor { row: 1, col: 0 }));
        assert_eq!(projected.formula1, original.formula1);
        assert_eq!(projected.ranges.len(), 2);

        original.formula1 = "=SUM(A:B)>0".into();
        original.formula_anchor = Some(FormulaAnchor { row: 100, col: 0 });
        let projected = project_rule(&original).unwrap();
        assert_eq!(projected.formula_anchor, Some(FormulaAnchor { row: 0, col: 0 }));
        assert_eq!(projected.formula1, original.formula1);
        assert_eq!(projected.ranges.len(), 1);
    }

    #[test]
    fn text_predicates_move_references_while_the_separate_matching_literal_remains_unchanged() {
        let mut original = rule(json!([{ "startRow": 2, "startCol": 2, "endRow": 5, "endCol": 2 }]), "=ISNUMBER(SEARCH(\"A1\",B2))");
        original.rule_type = "containsText".into(); original.sub_type = Some("text".into()); original.value = Some(json!("A1"));
        original.formula_anchor = Some(FormulaAnchor { row: 1, col: 1 });
        let projected = project_rule(&original).unwrap();
        assert_eq!(projected.formula1, "=ISNUMBER(SEARCH(\"A1\",C3))"); assert_eq!(projected.value, Some(json!("A1")));
        original.formula1 = "A1".into(); assert_eq!(project_rule(&original).unwrap().formula1, "A1");
    }

    #[test]
    fn office_model_fields_never_enter_the_known_expression_or_threshold_projection_path() {
        let mut original = rule(json!([{ "startRow": 2, "startCol": 2, "endRow": 5, "endCol": 2 }]), "=A1>0");
        original.rule_type = "office".into(); original.sub_type = Some("formula".into());
        original.config = Some(json!({ "min": { "type": "formula", "value": "=A1" }, "vendor": { "ref": "A1" } }));
        let projected = project_rule(&original).unwrap();
        assert_eq!(projected.formula1, original.formula1); assert_eq!(projected.config, original.config);
        assert_eq!(projected.formula_anchor, original.formula_anchor); assert!(formula_operands(&original).is_empty());
    }
}

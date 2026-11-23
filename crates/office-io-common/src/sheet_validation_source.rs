//! Worksheet-level metadata for preserving the editable validation model.
//! The standard SpreadsheetML projection remains authoritative after outside edits.

use std::collections::HashSet;
use std::io::{self, Write};

use serde::Serialize;
use serde_json::Value;

use crate::conditional_format_xml::{Element, MAIN_NAMESPACE, STRICT_MAIN_NAMESPACE, XmlSelection};

pub const MODEL_NAMESPACE: &str = "urn:office-studio:data-validation:1";
pub const MODEL_EXTENSION_URI: &str = "{0DF8C2D3-6701-4F72-A3C8-B52703ED40A1}";
pub const MODEL_BYTE_LIMIT: usize = 16 * 1024 * 1024;
pub const WORKBOOK_BYTE_LIMIT: usize = 64 * 1024 * 1024;
pub const WORKBOOK_ELEMENT_LIMIT: usize = 1_048_576;
pub const RANGE_LIMIT: usize = 262_144;

fn normalize_namespaces(node: &mut Element) {
    if node.namespace.is_empty() || node.namespace == STRICT_MAIN_NAMESPACE { node.namespace = MAIN_NAMESPACE.into(); }
    for child in &mut node.children {
        if child.is_main("dataValidation") || child.is_main("formula1") || child.is_main("formula2") { normalize_namespaces(child); }
    }
}

/// This hash detects stale model metadata; it is not an authentication signature.
pub fn projection_fingerprint(projection: Option<&Element>) -> String {
    let mut projection = projection.cloned().unwrap_or_else(|| Element::new(MAIN_NAMESPACE, "dataValidations"));
    normalize_namespaces(&mut projection);
    projection.fingerprint()
}

struct BoundedJson { bytes: Vec<u8>, limit: usize }
impl Write for BoundedJson {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if input.len() > self.limit.saturating_sub(self.bytes.len()) {
            return Err(io::Error::new(io::ErrorKind::InvalidData, "DV_SOURCE_MODEL_LIMIT: validation model exceeds its byte budget"));
        }
        self.bytes.extend_from_slice(input);
        Ok(input.len())
    }
    fn flush(&mut self) -> io::Result<()> { Ok(()) }
}

pub fn encode_model(value: &impl Serialize) -> Result<String, String> {
    let mut output = BoundedJson { bytes: Vec::new(), limit: MODEL_BYTE_LIMIT };
    serde_json::to_writer(&mut output, value).map_err(|error| error.to_string())?;
    String::from_utf8(output.bytes).map_err(|error| error.to_string())
}

#[derive(Default)]
pub struct ValidationSourceBudget { bytes: usize, elements: usize, model_bytes: usize }
impl ValidationSourceBudget {
    pub fn include_xml(&mut self, selection: &XmlSelection) -> Result<(), String> {
        self.bytes = self.bytes.checked_add(selection.selected_bytes).ok_or("DV_SOURCE_TOTAL_LIMIT: XML byte counter overflow")?;
        self.elements = self.elements.checked_add(selection.selected_elements).ok_or("DV_SOURCE_TOTAL_LIMIT: XML element counter overflow")?;
        if self.bytes > WORKBOOK_BYTE_LIMIT || self.elements > WORKBOOK_ELEMENT_LIMIT {
            return Err("DV_SOURCE_TOTAL_LIMIT: workbook validation XML exceeds its resource budget".into());
        }
        Ok(())
    }
    pub fn include_model(&mut self, bytes: usize) -> Result<(), String> {
        if bytes > MODEL_BYTE_LIMIT { return Err("DV_SOURCE_MODEL_LIMIT: validation model exceeds its byte budget".into()); }
        self.model_bytes = self.model_bytes.checked_add(bytes).ok_or("DV_SOURCE_TOTAL_LIMIT: model byte counter overflow")?;
        if self.model_bytes > WORKBOOK_BYTE_LIMIT { return Err("DV_SOURCE_TOTAL_LIMIT: workbook validation models exceed their byte budget".into()); }
        Ok(())
    }
}

fn coordinate(value: Option<&Value>, maximum: u64) -> bool {
    value.and_then(Value::as_u64).is_some_and(|value| value <= maximum)
}
fn optional(value: Option<&Value>, valid: impl FnOnce(&Value) -> bool) -> bool { value.is_none_or(|value| value.is_null() || valid(value)) }

/// Validate every typed field before an extension is allowed to replace imported rules.
pub fn validate_saved_rules(value: &Value) -> Result<(), String> {
    let rules = value.as_array().ok_or("DV_SOURCE_MODEL_INVALID: validation state is not an array")?;
    if rules.len() > RANGE_LIMIT { return Err("DV_SOURCE_MODEL_LIMIT: too many validation rules".into()); }
    let mut ids = HashSet::new();
    let mut range_count = 0usize;
    for rule in rules {
        let rule = rule.as_object().ok_or("DV_SOURCE_MODEL_INVALID: validation rule is not an object")?;
        let id = rule.get("id").and_then(Value::as_str).filter(|id| !id.trim().is_empty()).ok_or("DV_SOURCE_MODEL_INVALID: missing validation ID")?;
        if !ids.insert(id) { return Err("DV_SOURCE_MODEL_INVALID: duplicate validation ID".into()); }
        if !rule.get("type").and_then(Value::as_str).is_some_and(|kind| matches!(kind, "whole" | "decimal" | "date" | "time" | "textLength" | "list" | "listMultiple" | "checkbox" | "custom" | "any")) {
            return Err("DV_SOURCE_MODEL_INVALID: unsupported validation type".into());
        }
        let ranges = rule.get("ranges").and_then(Value::as_array).filter(|ranges| !ranges.is_empty()).ok_or("DV_SOURCE_MODEL_INVALID: missing validation ranges")?;
        range_count = range_count.checked_add(ranges.len()).ok_or("DV_SOURCE_MODEL_LIMIT: validation range counter overflow")?;
        if range_count > RANGE_LIMIT { return Err("DV_SOURCE_MODEL_LIMIT: too many validation ranges".into()); }
        for range in ranges {
            let range = range.as_object().ok_or("DV_SOURCE_MODEL_INVALID: validation range is not an object")?;
            if !coordinate(range.get("startRow"), 1_048_575) || !coordinate(range.get("endRow"), 1_048_575)
                || !coordinate(range.get("startCol"), 16_383) || !coordinate(range.get("endCol"), 16_383)
                || range["startRow"].as_u64() > range["endRow"].as_u64() || range["startCol"].as_u64() > range["endCol"].as_u64() {
                return Err("DV_SOURCE_MODEL_INVALID: validation range exceeds worksheet bounds".into());
            }
        }
        if !optional(rule.get("formulaAnchor"), |anchor| anchor.as_object().is_some_and(|anchor| coordinate(anchor.get("row"), 1_048_575) && coordinate(anchor.get("col"), 16_383))) {
            return Err("DV_SOURCE_MODEL_INVALID: invalid validation formula anchor".into());
        }
        for field in ["formula1", "formula2", "errorTitle", "error", "promptTitle", "prompt"] {
            if !optional(rule.get(field), Value::is_string) { return Err(format!("DV_SOURCE_MODEL_INVALID: {field} is not a string")); }
        }
        for field in ["allowBlank", "showDropDown", "showErrorMessage", "showInputMessage"] {
            if !optional(rule.get(field), Value::is_boolean) { return Err(format!("DV_SOURCE_MODEL_INVALID: {field} is not a boolean")); }
        }
        for (field, values) in [
            ("operator", &["between", "notBetween", "equal", "notEqual", "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual"][..]),
            ("errorStyle", &["stop", "warning", "info"][..]),
            ("renderMode", &["text", "arrow", "custom"][..]),
        ] {
            if !optional(rule.get(field), |value| value.as_str().is_some_and(|value| values.contains(&value))) { return Err(format!("DV_SOURCE_MODEL_INVALID: invalid {field}")); }
        }
        if !optional(rule.get("metadata"), Value::is_object) { return Err("DV_SOURCE_MODEL_INVALID: metadata is not an object".into()); }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::conditional_format_xml::parse;
    use serde_json::json;

    #[test]
    fn namespace_prefix_and_strict_dialect_changes_preserve_the_same_projection() {
        let ordinary = parse(r#"<dataValidations xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1"><dataValidation type="custom" sqref="C3"><formula1>A1&gt;0</formula1></dataValidation></dataValidations>"#).unwrap();
        let strict = parse(r#"<s:dataValidations xmlns:s="http://purl.oclc.org/ooxml/spreadsheetml/main" count="1"><s:dataValidation sqref="C3" type="custom"><s:formula1>A1&gt;0</s:formula1></s:dataValidation></s:dataValidations>"#).unwrap();
        assert_eq!(projection_fingerprint(Some(&ordinary)), projection_fingerprint(Some(&strict)));
        let edited = parse(&ordinary.render().replace("A1&gt;0", "B1&gt;0")).unwrap();
        assert_ne!(projection_fingerprint(Some(&ordinary)), projection_fingerprint(Some(&edited)));
    }

    #[test]
    fn invalid_ranges_anchors_and_duplicate_ids_never_replace_the_standard_rules() {
        let rule = json!({"id":"dv","type":"custom","ranges":[{"startRow":0,"startCol":0,"endRow":4,"endCol":0}],"formulaAnchor":{"row":1,"col":0},"formula1":"=A1>0","metadata":{"keep":true}});
        validate_saved_rules(&json!([rule.clone()])).unwrap();
        let mut invalid = rule.clone(); invalid["formulaAnchor"]["row"] = json!(1_048_576);
        assert!(validate_saved_rules(&json!([invalid])).is_err());
        let mut invalid = rule.clone(); invalid["ranges"][0]["endCol"] = json!(0.5);
        assert!(validate_saved_rules(&json!([invalid])).is_err());
        assert!(validate_saved_rules(&json!([rule.clone(), rule])).is_err());
    }

    #[test]
    fn model_serialization_fails_before_crossing_the_declared_budget() {
        let mut output = BoundedJson { bytes: Vec::new(), limit: 4 };
        assert!(serde_json::to_writer(&mut output, &json!("oversized")).is_err());
        assert!(output.bytes.len() <= 4);
        let mut budget = ValidationSourceBudget::default();
        assert!(budget.include_model(MODEL_BYTE_LIMIT + 1).is_err());
    }
}

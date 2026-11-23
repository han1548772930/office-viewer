use crate::{ExportError, ExportWarning, xml};
use office_io_common::conditional_format_xml;
use office_io_common::sheet_validation_source::{MODEL_EXTENSION_URI, MODEL_NAMESPACE, ValidationSourceBudget, encode_model, projection_fingerprint, validate_saved_rules};
use serde::{Deserialize, Serialize};
use serde_json::{Map, Value};

use super::conditional_formula::{FormulaAnchor, project_formula};

#[derive(Debug, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct DataValidationInput {
    #[serde(default)]
    id: String,
    #[serde(rename = "type", default)]
    validation_type: String,
    #[serde(default)]
    ranges: Vec<ValidationRange>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    formula_anchor: Option<FormulaAnchor>,
    #[serde(default)]
    operator: Option<String>,
    #[serde(default)]
    formula1: Option<String>,
    #[serde(default)]
    formula2: Option<String>,
    #[serde(default)]
    allow_blank: Option<bool>,
    #[serde(default)]
    show_drop_down: Option<bool>,
    #[serde(default)]
    show_error_message: Option<bool>,
    #[serde(default)]
    error_style: Option<String>,
    #[serde(default)]
    error_title: Option<String>,
    #[serde(default)]
    error: Option<String>,
    #[serde(default)]
    show_input_message: Option<bool>,
    #[serde(default)]
    prompt_title: Option<String>,
    #[serde(default)]
    prompt: Option<String>,
    #[serde(flatten)]
    extensions: Map<String, Value>,
}

#[derive(Debug, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
struct ValidationRange {
    start_row: u32,
    start_col: u32,
    end_row: u32,
    end_col: u32,
}

impl ValidationRange {
    fn reference(&self) -> Result<String, ExportError> {
        if self.start_row > self.end_row || self.start_col > self.end_col {
            return Err(ExportError::invalid(
                "Data validation range end precedes its start",
            ));
        }
        if self.end_row > 1_048_575 || self.end_col > 16_383 {
            return Err(ExportError::invalid(
                "Data validation range exceeds XLSX worksheet limits",
            ));
        }
        let first = cell_ref(self.start_row, self.start_col);
        let last = cell_ref(self.end_row, self.end_col);
        Ok(if first == last {
            first
        } else {
            format!("{first}:{last}")
        })
    }
}

pub(super) fn validate(sheet: &super::SheetInput) -> Result<(), ExportError> {
    for rule in &sheet.data_validations {
        if rule.ranges.is_empty() {
            return Err(ExportError::invalid(format!(
                "Data validation rule {:?} has no ranges",
                rule.id
            )));
        }
        for range in &rule.ranges {
            range.reference()?;
        }
        if !supported_type(&rule.validation_type) {
            return Err(ExportError::invalid(format!("Unsupported data validation type {:?}", rule.validation_type)));
        }
        if rule.formula_anchor.is_some_and(|anchor| anchor.row > 1_048_575 || anchor.col > 16_383) {
            return Err(ExportError::invalid("Data validation formula anchor exceeds XLSX worksheet limits"));
        }
        if rule.operator.as_deref().is_some_and(|operator| !matches!(operator,
            "between" | "notBetween" | "equal" | "notEqual" | "greaterThan" | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual")) {
            return Err(ExportError::invalid("Unsupported data validation operator"));
        }
        if rule.error_style.as_deref().is_some_and(|style| !matches!(style, "stop" | "warning" | "info")) {
            return Err(ExportError::invalid("Unsupported data validation error style"));
        }
        if rule.validation_type == "list" || rule.validation_type == "listMultiple" {
            if rule.formula1.as_deref().unwrap_or("").is_empty()
                && rule.validation_type != "checkbox"
            {
                return Err(ExportError::invalid(format!(
                    "Data validation rule {:?} has no list formula",
                    rule.id
                )));
            }
        }
    }
    Ok(())
}

pub(super) fn validate_with_budget(sheet: &super::SheetInput, budget: &mut ValidationSourceBudget) -> Result<(), ExportError> {
    validate(sheet)?;
    if !sheet.data_validations.is_empty() {
        budget.include_model(snapshot_model(sheet)?.len()).map_err(ExportError::invalid)?;
    }
    Ok(())
}

pub(super) fn write_xml(
    sheet: &super::SheetInput,
    warnings: &mut Vec<ExportWarning>,
) -> Result<String, ExportError> {
    validate(sheet)?;
    let supported_count = sheet
        .data_validations
        .iter()
        .filter(|rule| supported_type(&rule.validation_type))
        .count();
    if supported_count == 0 {
        return Ok(String::new());
    }
    let mut out = format!("<dataValidations count=\"{}\">", supported_count);
    for rule in &sheet.data_validations {
        let kind = match rule.validation_type.as_str() {
            "whole" => "whole",
            "decimal" => "decimal",
            "date" => "date",
            "time" => "time",
            "textLength" => "textLength",
            "list" | "listMultiple" | "checkbox" => "list",
            "custom" => "custom",
            "any" => "none",
            other => {
                warnings.push(ExportWarning::new(
                    "UNSUPPORTED_DATA_VALIDATION",
                    format!(
                        "Unsupported data validation type {other:?} on worksheet {}",
                        sheet.name
                    ),
                ));
                continue;
            }
        };
        let ranges = rule
            .ranges
            .iter()
            .map(ValidationRange::reference)
            .collect::<Result<Vec<_>, _>>()?;
        let Some(sqref) = (!ranges.is_empty()).then(|| ranges.join(" ")) else {
            continue;
        };
        let mut attrs = format!(" type=\"{kind}\" sqref=\"{}\"", xml::escape(&sqref));
        if let Some(operator) = rule.operator.as_deref() {
            let operator = match operator {
                "between" => "between",
                "notBetween" => "notBetween",
                "equal" => "equal",
                "notEqual" => "notEqual",
                "greaterThan" => "greaterThan",
                "greaterThanOrEqual" => "greaterThanOrEqual",
                "lessThan" => "lessThan",
                "lessThanOrEqual" => "lessThanOrEqual",
                other => {
                    warnings.push(ExportWarning::new(
                        "UNSUPPORTED_DATA_VALIDATION_OPERATOR",
                        format!("Unsupported validation operator {other:?}"),
                    ));
                    continue;
                }
            };
            attrs.push_str(&format!(" operator=\"{operator}\""));
        }
        attrs.push_str(&format!(
            " allowBlank=\"{}\" showDropDown=\"{}\" showInputMessage=\"{}\" showErrorMessage=\"{}\"",
            flag(rule.allow_blank.unwrap_or(true)),
            // OOXML uses the inverse of the product-facing show-dropdown option.
            flag(!rule.show_drop_down.unwrap_or(true)),
            flag(rule.show_input_message.unwrap_or(false)),
            flag(rule.show_error_message.unwrap_or(true)),
        ));
        if let Some(style) = rule.error_style.as_deref() {
            let style = match style {
                "stop" => "stop",
                "warning" => "warning",
                "info" => "information",
                other => {
                    warnings.push(ExportWarning::new(
                        "UNSUPPORTED_DATA_VALIDATION_STYLE",
                        format!("Unsupported validation error style {other:?}"),
                    ));
                    "stop"
                }
            };
            attrs.push_str(&format!(" errorStyle=\"{style}\""));
        }
        append_attr(&mut attrs, "errorTitle", rule.error_title.as_deref());
        append_attr(&mut attrs, "error", rule.error.as_deref());
        append_attr(&mut attrs, "promptTitle", rule.prompt_title.as_deref());
        append_attr(&mut attrs, "prompt", rule.prompt.as_deref());
        out.push_str(&format!("<dataValidation{attrs}>"));
        let target = FormulaAnchor { row: rule.ranges[0].start_row, col: rule.ranges[0].start_col };
        let formula1 = if rule.validation_type == "checkbox" {
            Some("\"TRUE,FALSE\"".to_owned())
        } else {
            projected_formula(rule, rule.formula1.as_deref(), target)
        };
        let formula2 = projected_formula(rule, rule.formula2.as_deref(), target);
        append_formula(&mut out, "formula1", formula1.as_deref());
        append_formula(&mut out, "formula2", formula2.as_deref());
        out.push_str("</dataValidation>");
    }
    out.push_str("</dataValidations>");
    Ok(out)
}

fn projected_formula(rule: &DataValidationInput, formula: Option<&str>, target: FormulaAnchor) -> Option<String> {
    formula.map(|formula| {
        if matches!(rule.validation_type.as_str(), "list" | "listMultiple" | "checkbox" | "any") { return formula.to_owned(); }
        project_formula(formula, rule.formula_anchor.unwrap_or(target), target)
    })
}

fn snapshot_model(sheet: &super::SheetInput) -> Result<String, ExportError> {
    // Bound the first serialization before materializing an editable JSON copy.
    let encoded = encode_model(&sheet.data_validations).map_err(ExportError::invalid)?;
    let mut rules: Value = serde_json::from_str(&encoded).map_err(|error| ExportError::invalid(error.to_string()))?;
    for (ordinal, rule) in rules.as_array_mut().ok_or_else(|| ExportError::invalid("Data validation snapshot is not an array"))?.iter_mut().enumerate() {
        if rule["id"].as_str().is_none_or(|id| id.trim().is_empty()) {
            rule["id"] = Value::String(format!("validation-{}-{}", sheet.id, ordinal + 1));
        }
    }
    validate_saved_rules(&rules).map_err(ExportError::invalid)?;
    encode_model(&rules).map_err(ExportError::invalid)
}

/// CT_DataValidation has no general extension container; metadata belongs to the worksheet extLst.
pub(super) fn extension_xml(sheet: &super::SheetInput, projection: &str) -> Result<String, ExportError> {
    if sheet.data_validations.is_empty() { return Ok(String::new()); }
    let standard = conditional_format_xml::parse(projection).map_err(ExportError::invalid)?;
    let model = snapshot_model(sheet)?;
    let xml = format!("<ext uri=\"{MODEL_EXTENSION_URI}\"><dv:validationState xmlns:dv=\"{MODEL_NAMESPACE}\" version=\"1\" fingerprint=\"{}\">{}</dv:validationState></ext>",
        projection_fingerprint(Some(&standard)), xml::escape(&model));
    conditional_format_xml::parse(&xml).and_then(|node| node.validate_fragment())
        .map_err(|error| ExportError::invalid(error.replace("CF_SOURCE_", "DV_SOURCE_")))?;
    Ok(xml)
}

fn append_attr(attrs: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value.filter(|value| !value.is_empty()) {
        attrs.push_str(&xml::attr(name, value));
    }
}

fn append_formula(out: &mut String, name: &str, formula: Option<&str>) {
    if let Some(formula) = formula.filter(|formula| !formula.is_empty()) {
        out.push_str(&format!(
            "<{name}>{}</{name}>",
            xml::escape(formula.trim_start_matches('='))
        ));
    }
}

fn flag(value: bool) -> u8 {
    u8::from(value)
}

fn supported_type(value: &str) -> bool {
    matches!(
        value,
        "whole"
            | "decimal"
            | "date"
            | "time"
            | "textLength"
            | "list"
            | "listMultiple"
            | "checkbox"
            | "custom"
            | "any"
    )
}

fn cell_ref(row: u32, column: u32) -> String {
    let mut letters = String::new();
    let mut value = column + 1;
    while value > 0 {
        let index = (value - 1) % 26;
        letters.insert(0, (b'A' + index as u8) as char);
        value = (value - 1) / 26;
    }
    format!("{letters}{}", row + 1)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn writes_validation_rules_with_multiple_ranges_and_messages() {
        let sheet: super::super::SheetInput = serde_json::from_value(serde_json::json!({
            "name":"Data",
            "dataValidations":[{
                "id":"dv-1", "type":"whole", "operator":"between", "formula1":"=1", "formula2":"10",
                "ranges":[{"startRow":0,"startCol":0,"endRow":4,"endCol":0},{"startRow":6,"startCol":1,"endRow":6,"endCol":2}],
                "allowBlank":false,"showDropDown":true,"showErrorMessage":true,"showInputMessage":true,
                "errorStyle":"warning","errorTitle":"Invalid","error":"Enter 1 to 10","prompt":"Integer only"
            }]
        })).unwrap();
        let xml = write_xml(&sheet, &mut Vec::new()).unwrap();
        assert!(xml.contains("<dataValidations count=\"1\">"));
        assert!(xml.contains("sqref=\"A1:A5 B7:C7\""));
        assert!(xml.contains("operator=\"between\""));
        assert!(xml.contains("showDropDown=\"0\""));
        assert!(xml.contains("errorStyle=\"warning\""));
        assert!(xml.contains("<formula1>1</formula1><formula2>10</formula2>"));
        assert!(xml.contains("error=\"Enter 1 to 10\""));
    }

    #[test]
    fn writes_list_checkbox_and_formula_validation_types() {
        let sheet: super::super::SheetInput = serde_json::from_value(serde_json::json!({
            "name":"Data",
            "dataValidations":[
                {"id":"list","type":"list","formula1":"\"A,B\"","ranges":[{"startRow":0,"startCol":0,"endRow":0,"endCol":0}]},
                {"id":"checkbox","type":"checkbox","ranges":[{"startRow":0,"startCol":1,"endRow":0,"endCol":1}]},
                {"id":"custom","type":"custom","formula1":"=A1>0","ranges":[{"startRow":0,"startCol":2,"endRow":0,"endCol":2}]}
            ]
        })).unwrap();
        let xml = write_xml(&sheet, &mut Vec::new()).unwrap();
        assert_eq!(xml.matches("type=\"list\"").count(), 2);
        assert!(xml.contains("<formula1>&quot;A,B&quot;</formula1>"));
        assert!(xml.contains("<formula1>&quot;TRUE,FALSE&quot;</formula1>"));
        assert!(xml.contains("type=\"custom\""));
    }

    #[test]
    fn invalid_ranges_fail_export_validation() {
        let sheet: super::super::SheetInput = serde_json::from_value(serde_json::json!({
            "name":"Data",
            "dataValidations":[{"id":"bad","type":"whole","ranges":[{"startRow":1,"startCol":0,"endRow":0,"endCol":0}]}]
        })).unwrap();
        assert!(validate(&sheet).is_err());
    }

    #[test]
    fn projects_custom_and_whole_axis_operands_from_the_retained_origin() {
        let sheet: super::super::SheetInput = serde_json::from_value(serde_json::json!({
            "id":"data", "name":"Data", "dataValidations":[{
                "id":"anchored", "type":"custom", "formulaAnchor":{"row":1,"col":1},
                "formula1":"=SUM(A:C,1:3)+A1>0", "ranges":[{"startRow":0,"startCol":0,"endRow":2,"endCol":0}]
            }]
        })).unwrap();
        let projection = write_xml(&sheet, &mut Vec::new()).unwrap();
        assert!(projection.contains("INDEX($A$1:$XFD$1048576,0,IF(COLUMN()-1&lt;1,16385,COLUMN()-1))"));
        assert!(projection.contains("INDEX($A$1:$XFD$1048576,IF(ROW()-1&lt;1,1048577,ROW()-1),0)"));
        assert_eq!(projection.matches("<dataValidation ").count(), 1);
        let extension = extension_xml(&sheet, &projection).unwrap();
        assert!(extension.contains("dv:validationState"));
        assert!(!projection.contains("extLst"));
        let extension = conditional_format_xml::parse(&extension).unwrap();
        let state = extension.children.iter().find(|node| node.is(MODEL_NAMESPACE, "validationState")).unwrap();
        let restored: Value = serde_json::from_str(&state.text).unwrap();
        assert_eq!(restored[0]["formulaAnchor"], serde_json::json!({"row":1,"col":1}));
    }

    #[test]
    fn list_sources_do_not_offset_and_any_values_keep_a_standard_prompt_rule() {
        let sheet: super::super::SheetInput = serde_json::from_value(serde_json::json!({
            "name":"Data", "dataValidations":[
                {"id":"list","type":"list","formula1":"=A1:A3","formulaAnchor":{"row":2,"col":2},"ranges":[{"startRow":10,"startCol":10,"endRow":12,"endCol":10}]},
                {"id":"any","type":"any","prompt":"Use any value","showInputMessage":true,"ranges":[{"startRow":0,"startCol":0,"endRow":0,"endCol":0}]}
            ]
        })).unwrap();
        let projection = write_xml(&sheet, &mut Vec::new()).unwrap();
        assert!(projection.contains("<formula1>A1:A3</formula1>"));
        assert!(projection.contains("type=\"none\""));
        assert!(projection.contains("prompt=\"Use any value\""));
    }

    #[test]
    fn invalid_origins_and_duplicate_model_ids_fail_the_real_export_boundary() {
        let mut model = serde_json::json!({"kind":"workbook","sheets":[{"id":"data","name":"Data","cells":{},"dataValidations":[{
            "id":"dv","type":"custom","formula1":"=A1>0","formulaAnchor":{"row":1_048_576,"col":0},"ranges":[{"startRow":0,"startCol":0,"endRow":0,"endCol":0}]
        }]}]});
        assert!(crate::export_xlsx_bytes(&model.to_string()).is_err());
        model["sheets"][0]["dataValidations"][0]["formulaAnchor"] = serde_json::json!({"row":1,"col":0});
        let duplicate = model["sheets"][0]["dataValidations"][0].clone();
        model["sheets"][0]["dataValidations"].as_array_mut().unwrap().push(duplicate);
        assert!(crate::export_xlsx_bytes(&model.to_string()).is_err());
    }
}

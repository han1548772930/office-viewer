use crate::{ExportError, ExportWarning, xml};
use serde::Deserialize;

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct DataValidationInput {
    #[serde(default)]
    id: String,
    #[serde(rename = "type", default)]
    validation_type: String,
    #[serde(default)]
    ranges: Vec<ValidationRange>,
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
}

#[derive(Debug, Deserialize)]
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

pub(super) fn write_xml(
    sheet: &super::SheetInput,
    warnings: &mut Vec<ExportWarning>,
) -> Result<String, ExportError> {
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
        let formula1 = if rule.validation_type == "checkbox" {
            Some("\"TRUE,FALSE\"".to_owned())
        } else {
            rule.formula1.clone()
        };
        append_formula(&mut out, "formula1", formula1.as_deref());
        append_formula(&mut out, "formula2", rule.formula2.as_deref());
        out.push_str("</dataValidation>");
    }
    out.push_str("</dataValidations>");
    Ok(out)
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
}

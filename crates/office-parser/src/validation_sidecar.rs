use std::collections::{HashMap, VecDeque};
use std::io::{BufRead, BufReader};

use office_io_common::conditional_format_xml::{self as source_xml, Element, XmlSelection, XmlSelectionLimits};
use office_io_common::sheet_validation_source::{MODEL_EXTENSION_URI, MODEL_NAMESPACE, RANGE_LIMIT, ValidationSourceBudget, projection_fingerprint, validate_saved_rules};
use zip::ZipArchive;

type RangeKey = (u32, u16, u32, u16);

#[derive(Debug, Clone, Default)]
pub(crate) struct ValidationSupplement {
    pub(crate) show_drop_down: Option<bool>,
    pub(crate) error_style: Option<String>,
    pub(crate) formula_anchor: Option<(u32, u16)>,
}

#[derive(Debug, Default)]
pub(crate) struct ValidationSidecar {
    enabled: bool,
    by_sheet: HashMap<String, HashMap<RangeKey, VecDeque<ValidationSupplement>>>,
    restored: HashMap<String, Vec<crate::model::DataValidationPayload>>,
    additional: HashMap<String, Vec<crate::model::DataValidationPayload>>,
    ranges_seen: usize,
    pub(crate) warnings: Vec<String>,
    pub(crate) failure: Option<String>,
}

impl ValidationSidecar {
    pub(crate) fn read(bytes: &[u8]) -> Self {
        match read_sidecar(bytes) {
            Ok(sidecar) => sidecar,
            Err(warning) => Self {
                warnings: vec![warning.clone()],
                failure: Some(warning),
                ..Self::default()
            },
        }
    }

    pub(crate) fn take(
        &mut self,
        sheet_name: &str,
        range: RangeKey,
    ) -> Option<ValidationSupplement> {
        if self.restored.contains_key(sheet_name) { return None; }
        let supplement = self
            .by_sheet
            .get_mut(sheet_name)
            .and_then(|ranges| ranges.get_mut(&range))
            .and_then(VecDeque::pop_front);
        if self.enabled && supplement.is_none() {
            self.warnings.push(format!(
                "Could not match imported data validation range {range:?} to its XLSX metadata on worksheet {sheet_name}"
            ));
        }
        supplement
    }

    pub(crate) fn restore(&mut self, sheet: &mut crate::model::SheetPayload) {
        if let Some(rules) = self.restored.remove(&sheet.name) {
            sheet.data_validations = rules;
            self.by_sheet.remove(&sheet.name);
        }
        if let Some(rules) = self.additional.remove(&sheet.name) { sheet.data_validations.extend(rules); }
    }

    pub(crate) fn warn_unmatched_rules(&mut self) {
        let unmatched = self
            .by_sheet
            .values()
            .flat_map(HashMap::values)
            .map(VecDeque::len)
            .sum::<usize>();
        if unmatched > 0 {
            self.warnings.push(format!(
                "rxls did not retain {unmatched} data validation range(s), usually due to a missing formula or malformed rule"
            ));
        }
    }
}

fn read_sidecar(bytes: &[u8]) -> Result<ValidationSidecar, String> {
    let mut archive = ZipArchive::new(std::io::Cursor::new(bytes))
        .map_err(|error| format!("DV_SOURCE_PACKAGE_INVALID: invalid XLSX ZIP package: {error}"))?;
    let sheet_refs = crate::xlsx_parts::worksheet_parts(&mut archive).map_err(|error| format!("DV_SOURCE_PACKAGE_INVALID: {error}"))?;
    let mut budget = ValidationSourceBudget::default();
    let mut sidecar = ValidationSidecar {
        enabled: true,
        ..ValidationSidecar::default()
    };
    for sheet in sheet_refs {
        let entry = archive.by_name(&sheet.path).map_err(|error| format!("DV_SOURCE_PACKAGE_INVALID: missing worksheet {}: {error}", sheet.path))?;
        let selected = read_validation_selection(BufReader::new(entry)).map_err(|error| format!("{error} [XLSX part {}]", sheet.path))?;
        budget.include_xml(&selected)?;
        if let Some(rules) = saved_validation_rules(&selected, &mut budget, &mut sidecar.warnings)? {
            sidecar.restored.insert(sheet.name, rules);
        } else {
            for container in selected.fragments.iter().filter(|node| node.is_main("dataValidations")) {
                parse_validation_container(&sheet.name, container, &mut sidecar)?;
            }
        }
    }
    Ok(sidecar)
}

fn read_validation_selection<R: BufRead>(input: R) -> Result<XmlSelection, String> {
    let selected = source_xml::extract(input, XmlSelectionLimits::default(), |path, node| {
        Ok(path.len() == 2 && path[0].is_main("worksheet") && path[1].is_main("dataValidations")
            || path.len() == 3 && path[0].is_main("worksheet") && path[1].is_main("extLst") && path[2].is_main("ext")
                && node.attribute_ns("", "uri") == Some(MODEL_EXTENSION_URI))
    }).map_err(|error| error.replace("CF_SOURCE_", "DV_SOURCE_"))?;
    if !selected.root.is_main("worksheet") { return Err("DV_SOURCE_XML_INVALID: invalid worksheet root".into()); }
    if selected.fragments.iter().filter(|node| node.is_main("dataValidations")).count() > 1 {
        return Err("DV_SOURCE_XML_INVALID: duplicate worksheet dataValidations container".into());
    }
    Ok(selected)
}

fn saved_validation_rules(selected: &XmlSelection, budget: &mut ValidationSourceBudget, warnings: &mut Vec<String>) -> Result<Option<Vec<crate::model::DataValidationPayload>>, String> {
    let extensions = selected.fragments.iter().filter(|node| node.is_main("ext")).collect::<Vec<_>>();
    if extensions.len() > 1 { return Err("DV_SOURCE_MODEL_INVALID: duplicate validation state extensions".into()); }
    let Some(extension) = extensions.first() else { return Ok(None); };
    let states = extension.children.iter().filter(|node| node.is(MODEL_NAMESPACE, "validationState")).collect::<Vec<_>>();
    if states.len() != 1 || extension.children.len() != 1 { return Err("DV_SOURCE_MODEL_INVALID: missing or duplicate validation state".into()); }
    let state = states[0];
    if state.attribute_ns("", "version") != Some("1") || !state.children.is_empty() {
        return Err("DV_SOURCE_MODEL_INVALID: unsupported or nested validation state".into());
    }
    let expected = state.attribute_ns("", "fingerprint").filter(|value| value.len() == 16 && value.bytes().all(|byte| byte.is_ascii_hexdigit()))
        .ok_or("DV_SOURCE_MODEL_INVALID: invalid validation projection fingerprint")?;
    let actual = projection_fingerprint(selected.fragments.iter().find(|node| node.is_main("dataValidations")));
    if !expected.eq_ignore_ascii_case(&actual) {
        warnings.push("Ignored stale validation model metadata after external edits to the standard rules".into());
        return Ok(None);
    }
    budget.include_model(state.text.len())?;
    let mut model: serde_json::Value = serde_json::from_str(&state.text).map_err(|error| format!("DV_SOURCE_MODEL_INVALID: {error}"))?;
    validate_saved_rules(&model)?;
    // Existing DTO consumers receive explicit Office defaults. Optional model
    // fields remain optional in the worksheet extension and are never guessed.
    for rule in model.as_array_mut().ok_or("DV_SOURCE_MODEL_INVALID: invalid rule array")? {
        let rule = rule.as_object_mut().ok_or("DV_SOURCE_MODEL_INVALID: invalid rule object")?;
        for (key, default) in [("operator", serde_json::json!("between")), ("formula1", serde_json::json!("")),
            ("allowBlank", serde_json::json!(true)), ("showErrorMessage", serde_json::json!(true)), ("showInputMessage", serde_json::json!(false))] {
            if rule.get(key).is_none_or(serde_json::Value::is_null) { rule.insert(key.into(), default); }
        }
    }
    serde_json::from_value(model).map(Some).map_err(|error| format!("DV_SOURCE_MODEL_INVALID: {error}"))
}

#[cfg(test)]
fn parse_worksheet_validations(sheet_name: &str, xml: &str, sidecar: &mut ValidationSidecar) {
    let root = match source_xml::parse(xml) {
        Ok(root) => root,
        Err(error) => { sidecar.failure = Some(error.replace("CF_SOURCE_", "DV_SOURCE_")); return; }
    };
    let container = if root.is_main("dataValidations") { Some(&root) } else { root.child("dataValidations") };
    if let Some(container) = container {
        if let Err(error) = parse_validation_container(sheet_name, container, sidecar) { sidecar.failure = Some(error); }
    }
}

fn parse_validation_container(sheet_name: &str, container: &Element, sidecar: &mut ValidationSidecar) -> Result<(), String> {
        for node in container.children.iter().filter(|node| node.is_main("dataValidation")) {
                let attrs = node.attrs.iter().filter(|(name, _)| node.attribute_namespaces.get(*name).is_none_or(|namespace| namespace.is_empty()))
                    .map(|(name, value)| (name.clone(), value.clone())).collect::<HashMap<_, _>>();
                let kind = attrs.get("type").map(String::as_str).unwrap_or("none");
                if !matches!(
                    kind,
                    "none" | "list" | "whole" | "decimal" | "date" | "time" | "textLength" | "custom"
                ) {
                    sidecar.warnings.push(format!(
                        "Skipped unsupported or missing data validation type {kind:?} on worksheet {sheet_name}"
                    ));
                    continue;
                }
                if let Some(operator) = attrs.get("operator")
                    && !matches!(
                        operator.as_str(),
                        "between"
                            | "notBetween"
                            | "equal"
                            | "notEqual"
                            | "greaterThan"
                            | "lessThan"
                            | "greaterThanOrEqual"
                            | "lessThanOrEqual"
                    )
                {
                    sidecar.warnings.push(format!(
                        "Unsupported data validation operator {operator:?} on worksheet {sheet_name}"
                    ));
                }
                for attribute in [
                    "allowBlank",
                    "showDropDown",
                    "showInputMessage",
                    "showErrorMessage",
                ] {
                    if attrs
                        .get(attribute)
                        .is_some_and(|value| parse_ooxml_bool(value).is_none())
                    {
                        sidecar.warnings.push(format!(
                            "Invalid {attribute} boolean {value:?} on worksheet {sheet_name}",
                            value = attrs.get(attribute).expect("attribute was present")
                        ));
                    }
                }
                let Some(sqref) = attrs.get("sqref") else {
                    sidecar.warnings.push(format!(
                        "Skipped data validation without sqref on worksheet {sheet_name}"
                    ));
                    continue;
                };
                if sqref.trim().is_empty() {
                    sidecar.warnings.push(format!(
                        "Skipped data validation with empty sqref on worksheet {sheet_name}"
                    ));
                    continue;
                }
                let ranges = sqref.split_whitespace().take(RANGE_LIMIT + 1).map(|range| (range, parse_range(range))).collect::<Vec<_>>();
                sidecar.ranges_seen = sidecar.ranges_seen.checked_add(ranges.len()).ok_or("DV_SOURCE_MODEL_LIMIT: validation range counter overflow")?;
                if sidecar.ranges_seen > RANGE_LIMIT { return Err("DV_SOURCE_MODEL_LIMIT: too many worksheet validation ranges".into()); }
                if kind == "none" {
                    let valid_ranges = ranges.into_iter().filter_map(|(text, range)| {
                        if range.is_none() { sidecar.warnings.push(format!("Skipped malformed data validation range {text:?} on worksheet {sheet_name}")); }
                        range.map(|(start_row, start_col, end_row, end_col)| crate::model::DataValidationRangePayload { start_row, start_col: u32::from(start_col), end_row, end_col: u32::from(end_col) })
                    }).collect::<Vec<_>>();
                    if !valid_ranges.is_empty() {
                        let rules = sidecar.additional.entry(sheet_name.to_owned()).or_default();
                        rules.push(prompt_only_validation(node, &attrs, valid_ranges, format!("validation-any-{}", rules.len() + 1)));
                    }
                    continue;
                }
                let formula_anchor = ranges.iter().find_map(|(_, range)| range.map(|range| (range.0, range.1)));
                let supplement = ValidationSupplement {
                    // OOXML showDropDown=true suppresses the in-cell arrow.
                    show_drop_down: attrs
                        .get("showDropDown")
                        .and_then(|value| parse_ooxml_bool(value).map(|hidden| !hidden)),
                    error_style: attrs.get("errorStyle").map(|value| match value.as_str() {
                        "information" => "info".to_owned(),
                        other => other.to_owned(),
                    }),
                    formula_anchor,
                };
                if supplement
                    .error_style
                    .as_deref()
                    .is_some_and(|style| !matches!(style, "stop" | "warning" | "info"))
                {
                    sidecar.warnings.push(format!(
                        "Unsupported data validation error style {:?} on worksheet {sheet_name}",
                        supplement.error_style
                    ));
                }
                for (range_text, range) in ranges {
                    match range {
                        Some(range) => sidecar
                            .by_sheet
                            .entry(sheet_name.to_owned())
                            .or_default()
                            .entry(range)
                            .or_default()
                            .push_back(supplement.clone()),
                        None => sidecar.warnings.push(format!(
                            "Skipped malformed data validation range {range_text:?} on worksheet {sheet_name}"
                        )),
                    }
                }
        }
    Ok(())
}

fn prompt_only_validation(node: &Element, attributes: &HashMap<String, String>, ranges: Vec<crate::model::DataValidationRangePayload>, id: String) -> crate::model::DataValidationPayload {
    crate::model::DataValidationPayload {
        id, validation_type: "any".into(), ranges, formula_anchor: None,
        operator: attributes.get("operator").cloned().unwrap_or_else(|| "between".into()),
        formula1: node.child("formula1").map(|node| node.text.clone()).unwrap_or_default(),
        formula2: node.child("formula2").map(|node| node.text.clone()),
        allow_blank: attributes.get("allowBlank").and_then(|value| parse_ooxml_bool(value)).unwrap_or(true),
        show_drop_down: attributes.get("showDropDown").and_then(|value| parse_ooxml_bool(value)).map(|value| !value),
        show_error_message: attributes.get("showErrorMessage").and_then(|value| parse_ooxml_bool(value)).unwrap_or(false),
        error_style: attributes.get("errorStyle").map(|value| if value == "information" { "info".into() } else { value.clone() }),
        error_title: attributes.get("errorTitle").cloned(), error: attributes.get("error").cloned(),
        show_input_message: attributes.get("showInputMessage").and_then(|value| parse_ooxml_bool(value)).unwrap_or(false),
        prompt_title: attributes.get("promptTitle").cloned(), prompt: attributes.get("prompt").cloned(),
        extensions: serde_json::Map::new(),
    }
}

fn parse_ooxml_bool(value: &str) -> Option<bool> {
    match value {
        "1" | "true" => Some(true),
        "0" | "false" => Some(false),
        _ => None,
    }
}

fn parse_range(value: &str) -> Option<RangeKey> {
    let (first, last) = match value.split_once(':') {
        Some((first, last)) if !last.contains(':') => (parse_cell(first)?, parse_cell(last)?),
        Some(_) => return None,
        None => {
            let cell = parse_cell(value)?;
            (cell, cell)
        }
    };
    (first.0 <= last.0 && first.1 <= last.1).then_some((first.0, first.1, last.0, last.1))
}

fn parse_cell(value: &str) -> Option<(u32, u16)> {
    let value = value.replace('$', "");
    let split = value.find(|character: char| character.is_ascii_digit())?;
    let (column, row) = value.split_at(split);
    if column.is_empty() || row.is_empty() || !column.bytes().all(|byte| byte.is_ascii_alphabetic())
    {
        return None;
    }
    let column = column.bytes().try_fold(0_u32, |index, byte| {
        index
            .checked_mul(26)?
            .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))
    })?;
    let row = row.parse::<u32>().ok()?;
    if column == 0 || column > 16_384 || row == 0 || row > 1_048_576 {
        return None;
    }
    Some((row - 1, u16::try_from(column - 1).ok()?))
}

#[cfg(test)]
mod tests {
    use super::{ValidationSidecar, parse_range, parse_worksheet_validations, read_validation_selection};

    #[test]
    fn parses_absolute_and_multi_cell_validation_ranges() {
        assert_eq!(parse_range("$A$2:C4"), Some((1, 0, 3, 2)));
        assert_eq!(
            parse_range("XFD1048576"),
            Some((1_048_575, 16_383, 1_048_575, 16_383))
        );
        assert_eq!(parse_range("C4:A2"), None);
        assert_eq!(parse_range("A0"), None);
    }

    #[test]
    fn warns_for_unsupported_types_and_malformed_ranges() {
        let xml = r#"<worksheet><dataValidations>
            <dataValidation type="extensionRule" sqref="A1"/>
            <dataValidation type="whole" sqref="A0"/>
            <dataValidation type="whole" sqref="A1" operator="invalid" showDropDown="yes"/>
            <dataValidation type="whole"/>
        </dataValidations></worksheet>"#;
        let mut sidecar = ValidationSidecar::default();

        parse_worksheet_validations("Input", xml, &mut sidecar);

        assert_eq!(sidecar.warnings.len(), 5);
        assert!(sidecar.warnings[0].contains("unsupported or missing data validation type"));
        assert!(sidecar.warnings[1].contains("malformed data validation range"));
        assert!(sidecar.warnings[2].contains("Unsupported data validation operator"));
        assert!(sidecar.warnings[3].contains("Invalid showDropDown boolean"));
        assert!(sidecar.warnings[4].contains("without sqref"));
    }

    #[test]
    fn every_disjoint_range_keeps_the_first_valid_sqref_formula_origin() {
        let mut sidecar = ValidationSidecar::default();
        parse_worksheet_validations("Input", r#"<worksheet><dataValidations><dataValidation type="custom" sqref="A0 C3:C4 E8:E9"><formula1>A1&gt;0</formula1></dataValidation></dataValidations></worksheet>"#, &mut sidecar);
        assert_eq!(sidecar.take("Input", (2, 2, 3, 2)).unwrap().formula_anchor, Some((2, 2)));
        assert_eq!(sidecar.take("Input", (7, 4, 8, 4)).unwrap().formula_anchor, Some((2, 2)));
        assert_eq!(sidecar.warnings.len(), 1);
    }

    #[test]
    fn scans_large_worksheet_cells_without_spending_the_validation_fragment_budget() {
        let rows = "<row><c><v>1</v></c></row>".repeat(800_000);
        let xml = format!("<worksheet><sheetData>{rows}</sheetData><dataValidations count=\"1\"><dataValidation type=\"whole\" sqref=\"C3\"><formula1>1</formula1></dataValidation></dataValidations></worksheet>");
        let selected = read_validation_selection(std::io::Cursor::new(xml.as_bytes())).unwrap();
        assert_eq!(selected.fragments.len(), 1);
        assert!(selected.selected_bytes < 1000);
        assert_eq!(selected.selected_elements, 3);
    }

    #[test]
    fn foreign_namespace_rules_and_nested_lookalikes_do_not_supply_validation_origins() {
        let xml = r#"<s:worksheet xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:other="urn:other"><s:dataValidations><other:dataValidation type="custom" sqref="A1"/><s:extLst><s:dataValidation type="custom" sqref="B2"/></s:extLst><s:dataValidation type="custom" sqref="C3 E8"><s:formula1>A1&gt;0</s:formula1></s:dataValidation></s:dataValidations></s:worksheet>"#;
        let mut sidecar = ValidationSidecar::default();
        parse_worksheet_validations("Input", xml, &mut sidecar);
        assert!(sidecar.take("Input", (0, 0, 0, 0)).is_none());
        assert!(sidecar.take("Input", (1, 1, 1, 1)).is_none());
        assert_eq!(sidecar.take("Input", (7, 4, 7, 4)).unwrap().formula_anchor, Some((2, 2)));
    }
}

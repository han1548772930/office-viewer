use super::{WorkbookInput, esc};
use crate::ExportError;
use serde_json::{Map, Value};
use std::collections::HashMap;

#[derive(Debug, PartialEq)]
struct DefinedName {
    name: String,
    reference: String,
    scope: Option<usize>,
    hidden: bool,
    comment: Option<String>,
}

fn sheet_index(workbook: &WorkbookInput, safe_names: &[String], value: &Value) -> Option<usize> {
    if let Some(index) = value.as_u64() {
        return usize::try_from(index).ok().filter(|index| *index < workbook.sheets.len());
    }
    let text = value.as_str()?;
    workbook.sheets.iter().position(|sheet| sheet.id == text)
        .or_else(|| workbook.sheets.iter().position(|sheet| sheet.name.to_lowercase() == text.to_lowercase()))
        .or_else(|| safe_names.iter().position(|name| name.to_lowercase() == text.to_lowercase()))
}
pub(super) fn read_scope(workbook: &WorkbookInput, safe_names: &[String], map: &Map<String, Value>,
    sheet_scope: Option<usize>, local: bool) -> Result<Option<usize>, ExportError> {
    let declared = map.get("sheetId").or_else(|| map.get("subUnitId")).or_else(|| map.get("sheet"))
        .or_else(|| map.get("localSheetId").filter(|value| value.as_str() != Some("AllDefaultWorkbook")));
    if let Some(value) = declared.filter(|value| !value.is_null()) {
        let index = sheet_index(workbook, safe_names, value)
            .ok_or_else(|| ExportError::invalid("local defined name refers to a missing worksheet"))?;
        if sheet_scope.is_some_and(|owner| owner != index) {
            return Err(ExportError::invalid("worksheet-local defined name has a conflicting owner"));
        }
        return Ok(Some(index));
    }
    if map.get("localSheetId").and_then(Value::as_str) == Some("AllDefaultWorkbook") { return Ok(None); }
    if let Some(index) = sheet_scope { return Ok(Some(index)); }
    if local { return Err(ExportError::invalid("local defined name is missing its worksheet scope")); }
    Ok(None)
}

/// Native workbook arrays and worksheet-local imported names share one OOXML
/// projection. An unresolved local scope cannot become a global name.
pub(super) fn write(workbook: &WorkbookInput, safe_names: &[String]) -> Result<String, ExportError> {
    let mut names = Vec::<DefinedName>::new();
    let mut identities = HashMap::<(Option<usize>, String), usize>::new();
    let mut read = |metadata: Option<&Value>, sheet_scope: Option<usize>| -> Result<(), ExportError> {
        let Some(metadata) = metadata.and_then(Value::as_object) else { return Ok(()); };
        for (key, local) in [("definedNames", false), ("localDefinedNames", true)] {
            let Some(values) = metadata.get(key) else { continue; };
            let items = values.as_array().ok_or_else(|| ExportError::invalid("defined names must be an array"))?;
            for item in items {
                let map = item.as_object().ok_or_else(|| ExportError::invalid("defined name must be an object"))?;
                let name = map.get("name").and_then(Value::as_str)
                    .filter(|name| !name.trim().is_empty() && name.chars().count() <= 255)
                    .ok_or_else(|| ExportError::invalid("defined name is missing its valid name"))?;
                let reference = map.get("refersTo").or_else(|| map.get("refers_to")).or_else(|| map.get("formulaOrRefString"))
                    .and_then(Value::as_str).filter(|value| !value.trim().is_empty())
                    .ok_or_else(|| ExportError::invalid("defined name is missing its reference or expression"))?;
                let scope = read_scope(workbook, safe_names, map, sheet_scope, local)?;
                let next = DefinedName { name: name.to_owned(), reference: reference.trim_start_matches('=').to_owned(), scope,
                    hidden: map.get("hidden").and_then(Value::as_bool).unwrap_or(false),
                    comment: map.get("comment").and_then(Value::as_str).map(str::to_owned) };
                let key = (scope, name.to_lowercase());
                if let Some(index) = identities.get(&key) {
                    let previous = &names[*index];
                    if previous.reference != next.reference || previous.hidden != next.hidden || previous.comment != next.comment {
                        return Err(ExportError::invalid("defined name has conflicting definitions in one scope"));
                    }
                    continue;
                }
                identities.insert(key, names.len()); names.push(next);
                // Up to two calendar views add private definitions beside the
                // admitted author catalog. The calendar projector separately
                // retains the 8192-author-name limit.
                if names.len() > 24_576 { return Err(ExportError::invalid("defined names exceed the export budget")); }
            }
        }
        Ok(())
    };
    read(workbook.metadata.as_ref(), None)?;
    for (index, sheet) in workbook.sheets.iter().enumerate() { read(sheet.metadata.as_ref(), Some(index))?; }
    if names.is_empty() { return Ok(String::new()); }
    let mut output = String::from("<definedNames>");
    for item in names {
        output.push_str(&format!("<definedName name=\"{}\"", esc(&item.name)));
        if let Some(index) = item.scope { output.push_str(&format!(" localSheetId=\"{index}\"")); }
        if item.hidden { output.push_str(" hidden=\"1\""); }
        if let Some(comment) = item.comment { output.push_str(&format!(" comment=\"{}\"", esc(&comment))); }
        output.push_str(&format!(">{}</definedName>", esc(&item.reference)));
    }
    output.push_str("</definedNames>");
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;
    fn project(value: Value) -> Result<String, ExportError> {
        let workbook: WorkbookInput = serde_json::from_value(value).unwrap();
        let safe_names = workbook.sheets.iter().map(|sheet| sheet.name.clone()).collect::<Vec<_>>();
        write(&workbook, &safe_names)
    }
    #[test]
    fn imported_local_names_follow_ids_after_worksheet_reordering() {
        let xml = project(serde_json::json!({"sheets":[{"id":"destination","name":"Destination"},
            {"id":"imported","name":"Chart data","metadata":{"definedNames":[{"name":"Tax", "refersTo":"=A1", "sheetId":"imported", "localSheetId":0}]}}]})).unwrap();
        assert!(xml.contains("name=\"Tax\" localSheetId=\"1\">A1</definedName>"));
    }
    #[test]
    fn public_string_scopes_and_global_sentinel_preserve_attributes() {
        let xml = project(serde_json::json!({"sheets":[{"id":"s","name":"Source"}], "metadata":{"definedNames":[
            {"name":"Global", "formulaOrRefString":"=2", "localSheetId":"AllDefaultWorkbook"},
            {"name":"Local", "formulaOrRefString":"=Source!A1", "localSheetId":"s", "hidden":true, "comment":"A & B"}]}})).unwrap();
        assert!(xml.contains("name=\"Global\">2</definedName>"));
        assert!(xml.contains("name=\"Local\" localSheetId=\"0\" hidden=\"1\" comment=\"A &amp; B\""));
    }
    #[test]
    fn missing_local_scope_is_rejected_instead_of_becoming_global() {
        assert!(project(serde_json::json!({"sheets":[{"id":"s","name":"Source"}], "metadata":{"localDefinedNames":[
            {"name":"Lost", "refersTo":"A1", "sheetId":"closed"}]}})).is_err());
    }
    #[test]
    fn equal_duplicates_are_deduplicated_and_conflicting_ones_rejected() {
        let make = |reference: &str| serde_json::json!({"sheets":[{"id":"s","name":"Source", "metadata":{"localDefinedNames":[
            {"name":"local", "refersTo":reference}]}}], "metadata":{"localDefinedNames":[{"name":"Local", "refersTo":"A1", "sheetId":"s"}]}});
        assert_eq!(project(make("A1")).unwrap().matches("<definedName ").count(), 1);
        assert!(project(make("B1")).is_err());
    }
}

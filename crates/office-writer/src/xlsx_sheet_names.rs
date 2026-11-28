//! Normalize physical worksheet names before generating any native reference.
//! Authored INDIRECT strings remain aliases; only alias target values change.

use crate::ExportError;
use office_io_common::formula_references::rewrite_sheet_qualifiers;
use serde_json::Value;
use std::collections::BTreeMap;

pub(super) fn prepare(value: &mut Value) -> Result<bool, ExportError> {
    let Some(sheets) = value.get("sheets").and_then(Value::as_array) else { return Ok(false); };
    let names = sheets.iter().map(|sheet| sheet.get("name").and_then(Value::as_str).unwrap_or("")).collect::<Vec<_>>();
    let safe = super::safe_sheet_names(names.iter().copied());
    if names.iter().zip(&safe).all(|(before, after)| *before == after) { return Ok(false); }
    let mut renamed = BTreeMap::new();
    for (before, after) in names.into_iter().zip(&safe) {
        // Invalid duplicate names bind to their first sheet, as they did before
        // normalization. Stable sheet IDs still distinguish all object targets.
        renamed.entry(before.to_lowercase()).or_insert_with(|| after.clone());
    }
    rewrite(value, &renamed, 0)?;
    for (sheet, name) in value.get_mut("sheets").and_then(Value::as_array_mut).into_iter().flatten().zip(safe) {
        if let Some(sheet) = sheet.as_object_mut() { sheet.insert("name".into(), Value::String(name)); }
    }
    Ok(true)
}

fn rewrite(value: &mut Value, names: &BTreeMap<String, String>, depth: usize) -> Result<(), ExportError> {
    if depth > 256 { return Err(ExportError::invalid("XLSX_SHEET_NAME_LIMIT: snapshot nesting exceeds 256")); }
    match value {
        Value::Array(values) => for value in values { rewrite(value, names, depth + 1)?; },
        Value::Object(values) => {
            for (key, value) in values {
                if key == "formulaAliases" {
                    if let Some(aliases) = value.as_object_mut() {
                        for key in ["sheets", "nameSheets"] {
                            for target in aliases.get_mut(key).and_then(Value::as_object_mut).into_iter().flat_map(|map| map.values_mut()) {
                                if let Some(before) = target.as_str() && let Some(after) = names.get(&before.to_lowercase()) {
                                    *target = Value::String(after.clone());
                                }
                            }
                        }
                    }
                    continue;
                }
                if matches!(key.as_str(), "formula" | "refersTo" | "refers_to" | "formulaOrRefString" | "targetFormula" |
                    "projectedFormula" | "calculationFormula" | "formula1" | "formula2") {
                    if let Some(formula) = value.as_str() { *value = Value::String(rewrite_sheet_qualifiers(formula, names)); }
                } else if matches!(key.as_str(), "sheet" | "sheetName" | "calculationSheet") {
                    if let Some(before) = value.as_str() && let Some(after) = names.get(&before.to_lowercase()) { *value = Value::String(after.clone()); }
                } else { rewrite(value, names, depth + 1)?; }
            }
        }
        _ => {}
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn safe_names_rewrite_native_formulas_and_provenance_but_keep_authored_strings_and_alias_keys() {
        let mut data = json!({ "sheets": [{ "id": "s", "name": "Q1/2026", "cells": {
            "A1": { "formula": "='Q1/2026'!B2+INDIRECT(\"'Q1/2026'!C3\")", "metadata": {
                "chartDependencyFormulaSource": { "formula": "='Q1/2026'!B2+INDIRECT(\"'Q1/2026'!C3\")" },
                "chartDependencyDateSource": { "targetFormula": "='Q1/2026'!B2+INDIRECT(\"'Q1/2026'!C3\")" }
            } }
        }, "metadata": { "chartDependencyImport": { "formulaAliases": {
            "sheets": { "Q1/2026": "Q1/2026" }, "nameSheets": { "Tax": "Q1/2026" }
        } } } }], "metadata": { "localDefinedNames": [{ "sheet": "Q1/2026", "refersTo": "='Q1/2026'!B2" }] } });
        assert!(prepare(&mut data).unwrap());
        assert_eq!(data["sheets"][0]["name"], "Q1_2026");
        assert_eq!(data["sheets"][0]["cells"]["A1"]["formula"], "='Q1_2026'!B2+INDIRECT(\"'Q1/2026'!C3\")");
        assert_eq!(data["sheets"][0]["cells"]["A1"]["metadata"]["chartDependencyFormulaSource"]["formula"], data["sheets"][0]["cells"]["A1"]["formula"]);
        assert_eq!(data["sheets"][0]["metadata"]["chartDependencyImport"]["formulaAliases"]["sheets"]["Q1/2026"], "Q1_2026");
        assert_eq!(data["metadata"]["localDefinedNames"][0]["sheet"], "Q1_2026");
        assert!(!prepare(&mut data).unwrap());
    }
}

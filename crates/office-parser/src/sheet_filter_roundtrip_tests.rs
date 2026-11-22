//! Deferred binary fixtures for typed filters and independent hide reasons.

use std::io::{Cursor, Read, Write};

use serde_json::{Value, json};

use crate::parser::parse_office_bytes;

fn export(sheet: Value) -> Vec<u8> {
    office_writer::export_xlsx_bytes(&json!({ "kind": "workbook", "sheets": [sheet] }).to_string()).unwrap().0
}
fn parsed(bytes: &[u8]) -> Value { serde_json::to_value(parse_office_bytes(bytes, "filter-roundtrip.xlsx").unwrap()).unwrap() }
fn rewrite(bytes: &[u8], edit: impl Fn(String) -> String) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap();
        let name = entry.name().to_owned();
        let mut content = Vec::new();
        entry.read_to_end(&mut content).unwrap();
        if name == "xl/worksheets/sheet1.xml" { content = edit(String::from_utf8(content).unwrap()).into_bytes(); }
        writer.start_file(name, options).unwrap();
        writer.write_all(&content).unwrap();
    }
    writer.finish().unwrap().into_inner()
}
fn source() -> Value {
    json!({ "id": "filter", "name": "Filter & Axes", "axisUnit": "cssPixel", "rowCount": 20, "columnCount": 4,
        "rowHeights": { "4": 32.5, "5": 0 }, "hiddenRows": [3],
        "cells": { "0:1": { "row": 0, "col": 1, "value": "Header" }, "1:1": { "row": 1, "col": 1, "value": 1 },
            "2:1": { "row": 2, "col": 1, "value": 2 }, "3:1": { "row": 3, "col": 1, "value": 3 },
            "4:1": { "row": 4, "col": 1, "value": "1" } },
        "autoFilter": { "range": { "startRow": 0, "startCol": 1, "endRow": 9, "endCol": 1 },
            "columns": { "1": { "kind": "values", "selectedItems": [1, "1", null, ""] } }, "filteredRows": [2, 3, 5] } })
}

#[test]
fn typed_members_manual_filter_overlap_and_zero_dimensions_survive_export_and_reopen() {
    let bytes = export(source());
    let reopened = parsed(&bytes);
    let sheet = &reopened["payload"]["data"]["sheets"][0];
    assert_eq!(sheet["autoFilter"]["columns"]["1"]["selectedItems"], json!([1, "1", null, ""]));
    assert_eq!(sheet["autoFilter"]["filteredRows"], json!([2, 3, 5]));
    assert_eq!(sheet["hiddenRows"], json!([3]));
    assert_eq!(sheet["rowHeights"], json!([{ "index": 4, "size": 32.5 }, { "index": 5, "size": 0 }]));
    assert_eq!(sheet["sourceAxisDimensions"]["projectAutoFilterExtension"]["hiddenReasonsExact"], true);
    assert_eq!(sheet["sourceAutoFilter"]["name"], "autoFilter");
}

#[test]
fn explicit_null_conditions_reopen_without_becoming_missing_operands() {
    let mut sheet = source();
    sheet["autoFilter"]["columns"]["1"] = json!({ "kind": "condition", "operator": "equals", "value": null, "caseSensitive": true });
    let bytes = export(sheet);
    let reopened = parsed(&bytes);
    let condition = &reopened["payload"]["data"]["sheets"][0]["autoFilter"]["columns"]["1"];
    assert!(condition.as_object().unwrap().contains_key("value"));
    assert_eq!(condition["value"], Value::Null);
    assert_eq!(condition["caseSensitive"], true);
}

#[test]
fn independent_filter_fingerprint_keeps_typed_criteria_after_external_row_resize() {
    let bytes = export(source());
    let resized = rewrite(&bytes, |xml| xml.replace("ht=\"24.375\"", "ht=\"30\""));
    let reopened = parsed(&resized);
    let sheet = &reopened["payload"]["data"]["sheets"][0];
    assert_eq!(sheet["autoFilter"]["columns"]["1"]["selectedItems"], json!([1, "1", null, ""]));
    assert_eq!(sheet["hiddenRows"], json!([3]));
    assert_eq!(sheet["sourceAxisDimensions"]["projectAutoFilterExtension"]["restored"], true);
    assert!(reopened["diagnostics"]["warnings"].as_array().unwrap().iter().any(|warning| warning.as_str().is_some_and(|value| value.contains("row/column dimensions changed"))));
}

#[test]
fn externally_edited_criteria_use_the_standard_filter_instead_of_the_stale_snapshot() {
    let mut sheet = source();
    sheet["autoFilter"]["columns"]["1"] = json!({ "kind": "values", "selectedItems": [1] });
    // Keep all data in the column numeric for a deterministic standard projection.
    sheet["cells"]["4:1"]["value"] = json!(4);
    let bytes = export(sheet);
    let edited = rewrite(&bytes, |xml| xml.replace("<filter val=\"1\"/>", "<filter val=\"2\"/>"));
    let reopened = parsed(&edited);
    let sheet = &reopened["payload"]["data"]["sheets"][0];
    assert_eq!(sheet["autoFilter"]["columns"]["1"]["selectedItems"], json!([2]));
    assert_eq!(sheet["hiddenRows"], json!([3]));
    assert_eq!(sheet["sourceAxisDimensions"]["projectAutoFilterExtension"]["restored"], false);
    assert!(reopened["diagnostics"]["warnings"].as_array().unwrap().iter().any(|warning| warning.as_str().is_some_and(|value| value.contains("AutoFilter criteria changed"))));
}

#[test]
fn unprojected_standard_criteria_retain_the_complete_resolved_xml_and_visibility() {
    let bytes = export(source());
    let edited = rewrite(&bytes, |xml| {
        let start = xml.find("<autoFilter ").unwrap();
        let end = xml[start..].find("</autoFilter>").unwrap() + start + "</autoFilter>".len();
        format!("{}<autoFilter ref=\"B1:B10\"><filterColumn colId=\"0\"><top10 top=\"1\" percent=\"0\" val=\"3\" filterVal=\"2\"/></filterColumn></autoFilter>{}", &xml[..start], &xml[end..])
    });
    let reopened = parsed(&edited);
    let sheet = &reopened["payload"]["data"]["sheets"][0];
    assert!(sheet.get("autoFilter").is_none());
    assert_eq!(sheet["sourceAutoFilter"]["children"][0]["children"][0]["name"], "top10");
    assert_eq!(sheet["sourceAutoFilter"]["children"][0]["children"][0]["attributes"]["filterVal"], "2");
    assert_eq!(sheet["hiddenRows"], json!([2, 3, 5]));
    assert!(reopened["diagnostics"]["warnings"].as_array().unwrap().iter().any(|warning| warning.as_str().is_some_and(|value| value.contains("top10") && value.contains("retained"))));
}

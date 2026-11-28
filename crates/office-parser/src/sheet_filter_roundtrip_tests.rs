//! Deferred binary fixtures for typed filters and independent hide reasons.

use std::io::{Cursor, Read, Write};

use serde_json::{Value, json};

use crate::parser::parse_office_bytes;

fn export(sheet: Value) -> Vec<u8> {
    office_writer::export_xlsx_bytes(&json!({ "kind": "workbook", "sheets": [sheet] }).to_string()).unwrap().0
}
fn parsed(bytes: &[u8]) -> Value { serde_json::to_value(parse_office_bytes(bytes, "filter-roundtrip.xlsx").unwrap()).unwrap() }
fn rewrite(bytes: &[u8], edit: impl Fn(String) -> String) -> Vec<u8> {
    rewrite_part(bytes, "xl/worksheets/sheet1.xml", edit)
}
fn rewrite_part(bytes: &[u8], part: &str, edit: impl Fn(String) -> String) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap();
        let name = entry.name().unwrap().into_owned();
        let mut content = Vec::new();
        entry.read_to_end(&mut content).unwrap();
        if name == part { content = edit(String::from_utf8(content).unwrap()).into_bytes(); }
        writer.start_file(name, options).unwrap();
        writer.write_all(&content).unwrap();
    }
    writer.finish().unwrap().into_inner()
}

#[test]
fn fill_font_no_fill_and_multicolor_typed_filters_reopen_with_exact_selection_and_hide_reasons() {
    for criterion in [
        json!({ "kind": "color", "cellFillColors": ["#FF0000"] }),
        json!({ "kind": "color", "cellFillColors": [null] }),
        json!({ "kind": "color", "cellTextColors": ["#0000FF"] }),
        json!({ "kind": "color", "cellFillColors": ["#FF0000", null], "cellTextColors": ["#0000FF"] }),
        json!({ "kind": "color", "cellFillColors": [] }),
    ] {
        let mut sheet = source(); sheet["autoFilter"]["columns"]["1"] = criterion.clone();
        let bytes = export(sheet); let reopened = parsed(&bytes);
        let sheet = &reopened["payload"]["data"]["sheets"][0];
        assert_eq!(sheet["autoFilter"]["columns"]["1"], criterion);
        assert_eq!(sheet["hiddenRows"], json!([3]));
        assert_eq!(sheet["autoFilter"]["filteredRows"], json!([2, 3, 5]));
        assert_eq!(sheet["sourceAxisDimensions"]["projectAutoFilterExtension"]["restored"], true);
    }
}

#[test]
fn external_dxf_edit_invalidates_typed_color_snapshot_without_reusing_stale_multiple_colors() {
    let mut sheet = source(); sheet["autoFilter"]["columns"]["1"] = json!({ "kind": "color", "cellFillColors": ["#FF0000", null] });
    let bytes = export(sheet);
    let edited = rewrite_part(&bytes, "xl/styles.xml", |xml| xml.replace("FFFF0000", "FF00FF00"));
    let reopened = parsed(&edited); let sheet = &reopened["payload"]["data"]["sheets"][0];
    assert_eq!(sheet["autoFilter"]["columns"]["1"], json!({ "kind": "color", "cellFillColors": ["#00FF00"] }));
    assert_eq!(sheet["sourceAxisDimensions"]["projectAutoFilterExtension"]["restored"], false);
    assert_eq!(sheet["sourceAutoFilter"]["children"][0]["children"][0]["name"], "colorFilter");
}

#[test]
fn missing_referenced_filter_color_style_retains_complete_native_criteria_and_visibility() {
    let mut sheet = source(); sheet["autoFilter"]["columns"]["1"] = json!({ "kind": "color", "cellTextColors": ["#0000FF"] });
    let bytes = export(sheet);
    let edited = rewrite(&bytes, |xml| xml.replace("<colorFilter dxfId=\"0\"", "<colorFilter dxfId=\"900000\""));
    let reopened = parsed(&edited); let sheet = &reopened["payload"]["data"]["sheets"][0];
    assert!(sheet.get("autoFilter").is_none());
    assert_eq!(sheet["sourceAutoFilter"]["children"][0]["children"][0]["attributes"]["dxfId"], "900000");
    assert_eq!(sheet["hiddenRows"], json!([2, 3, 5]));
}

#[test]
fn table_filter_snapshot_keeps_multiple_colors_while_native_xml_uses_real_dxf_and_current_visibility() {
    let mut sheet = source(); let filter = json!({
        "range": { "startRow": 0, "startCol": 1, "endRow": 9, "endCol": 1 },
        "columns": { "1": { "kind": "color", "cellFillColors": ["#FF0000", null] } }, "filteredRows": [2, 5]
    });
    sheet.as_object_mut().unwrap().remove("autoFilter");
    let workbook = json!({ "kind": "workbook", "sheets": [sheet], "tables": [{
        "id": "table", "name": "ColorTable", "sheetId": "filter", "range": filter["range"], "showHeaderRow": true,
        "columns": [{ "id": "field", "name": "Header" }], "autoFilter": filter
    }] });
    let (bytes, warnings) = office_writer::export_xlsx_bytes(&workbook.to_string()).unwrap();
    assert!(warnings.iter().any(|warning| warning.code == "AUTOFILTER_MULTIPLE_COLORS"));
    let mut archive = zip::ZipArchive::new(Cursor::new(&bytes)).unwrap();
    let mut native = String::new(); archive.by_name("xl/tables/table1.xml").unwrap().read_to_string(&mut native).unwrap();
    assert!(native.contains("colorFilter")); assert!(native.contains("dxfId=\"0\"")); assert!(native.contains("filterState"));
    let reopened = parsed(&bytes); let table = &reopened["payload"]["data"]["tables"][0];
    assert_eq!(table["autoFilter"]["columns"]["1"]["cellFillColors"], json!(["#FF0000", null]));
    assert_eq!(table["autoFilter"]["filteredRows"], json!([2, 5]));
    assert_eq!(reopened["payload"]["data"]["sheets"][0]["hiddenRows"], json!([3]));
    let edited = rewrite_part(&bytes, "xl/styles.xml", |xml| xml.replace("FFFF0000", "FF00FF00"));
    let edited = parsed(&edited);
    assert_eq!(edited["payload"]["data"]["tables"][0]["autoFilter"]["columns"]["1"]["cellFillColors"], json!(["#00FF00"]));
    let unavailable = rewrite_part(&bytes, "xl/tables/table1.xml", |xml| xml.replace("dxfId=\"0\"", "dxfId=\"900000\""));
    let unavailable = parsed(&unavailable);
    assert!(unavailable["payload"]["data"]["tables"][0].get("autoFilter").is_none());
    assert_eq!(unavailable["payload"]["data"]["sheets"][0]["hiddenRows"], json!([2, 3, 5]));
}

#[test]
fn merged_filter_groups_beyond_the_native_reference_remain_in_the_exact_hidden_cache() {
    let mut sheet = source();
    sheet["mergedRanges"] = json!([{ "startRow": 1, "startCol": 1, "endRow": 12, "endCol": 1 }]);
    sheet["autoFilter"]["filteredRows"] = json!([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]);
    let bytes = export(sheet); let reopened = parsed(&bytes);
    let sheet = &reopened["payload"]["data"]["sheets"][0];
    assert_eq!(sheet["autoFilter"]["range"]["endRow"], 9);
    assert_eq!(sheet["autoFilter"]["filteredRows"], json!([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]));
    assert_eq!(sheet["hiddenRows"], json!([3]));
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
fn compound_conditions_reopen_with_exact_join_operands_and_independent_hide_reasons() {
    for join in ["and", "or"] {
        let criterion = json!({ "kind": "compound", "join": join, "conditions": [
            { "kind": "condition", "operator": "greaterThan", "value": 2, "caseSensitive": false },
            { "kind": "condition", "operator": "doesNotEqual", "value": 3, "caseSensitive": false }
        ] });
        let mut sheet = source(); sheet["autoFilter"]["columns"]["1"] = criterion.clone();
        let reopened = parsed(&export(sheet));
        let sheet = &reopened["payload"]["data"]["sheets"][0];
        assert_eq!(sheet["autoFilter"]["columns"]["1"], criterion);
        assert_eq!(sheet["hiddenRows"], json!([3]));
        assert_eq!(sheet["autoFilter"]["filteredRows"], json!([2, 3, 5]));
        let native = &sheet["sourceAutoFilter"]["children"][0]["children"][0];
        assert_eq!(native["name"], "customFilters");
        assert_eq!(native["attributes"]["and"], if join == "and" { "1" } else { "0" });
        assert_eq!(native["children"].as_array().unwrap().len(), 2);
    }
}

#[test]
fn externally_changed_compound_join_invalidates_the_saved_predicate_and_projects_both_native_terms() {
    let mut sheet = source(); sheet["cells"]["4:1"]["value"] = json!(4);
    sheet["autoFilter"]["columns"]["1"] = json!({ "kind": "compound", "join": "and", "conditions": [
        { "kind": "condition", "operator": "greaterThan", "value": 2 },
        { "kind": "condition", "operator": "doesNotEqual", "value": 3 }
    ] });
    let edited = rewrite(&export(sheet), |xml| xml.replace("<customFilters and=\"1\">", "<customFilters and=\"0\">"));
    let reopened = parsed(&edited); let sheet = &reopened["payload"]["data"]["sheets"][0];
    assert_eq!(sheet["autoFilter"]["columns"]["1"], json!({ "kind": "compound", "join": "or", "conditions": [
        { "kind": "condition", "operator": "greaterThan", "value": 2.0 },
        { "kind": "condition", "operator": "doesNotEqual", "value": 3.0 }
    ] }));
    assert_eq!(sheet["sourceAxisDimensions"]["projectAutoFilterExtension"]["restored"], false);
}

#[test]
fn native_two_predicate_limit_reports_a_snapshot_without_losing_the_complete_compound_expression() {
    let criterion = json!({ "kind": "compound", "join": "or", "conditions": [
        { "kind": "condition", "operator": "between", "value": 1, "secondValue": 2, "caseSensitive": false },
        { "kind": "condition", "operator": "between", "value": 4, "secondValue": 8, "caseSensitive": false }
    ] });
    let mut sheet = source(); sheet["autoFilter"]["columns"]["1"] = criterion.clone();
    let model = json!({ "kind": "workbook", "sheets": [sheet] });
    let (bytes, warnings) = office_writer::export_xlsx_bytes(&model.to_string()).unwrap();
    assert!(warnings.iter().any(|warning| warning.code == "AUTOFILTER_COMPOUND_SNAPSHOT"));
    let reopened = parsed(&bytes);
    assert_eq!(reopened["payload"]["data"]["sheets"][0]["autoFilter"]["columns"]["1"], criterion);
}

#[test]
fn compound_writer_rejects_empty_nested_unknown_join_and_missing_interval_operands() {
    for criterion in [
        json!({ "kind": "compound", "join": "and", "conditions": [] }),
        json!({ "kind": "compound", "join": "xor", "conditions": [{ "kind": "condition", "operator": "isBlank" }] }),
        json!({ "kind": "compound", "join": "and", "conditions": [{ "kind": "compound", "join": "or", "conditions": [] }] }),
        json!({ "kind": "compound", "join": "or", "conditions": [{ "kind": "condition", "operator": "between", "value": 1 }] }),
    ] {
        let mut sheet = source(); sheet["autoFilter"]["columns"]["1"] = criterion;
        assert!(office_writer::export_xlsx_bytes(&json!({ "kind": "workbook", "sheets": [sheet] }).to_string()).is_err());
    }
}

#[test]
fn fixed_text_wildcards_retain_star_question_and_literal_tilde_in_native_custom_filters() {
    let mut sheet = source();
    sheet["autoFilter"]["columns"]["1"] = json!({ "kind": "condition", "operator": "equals", "value": "A*B?~", "caseSensitive": false });
    let bytes = export(sheet); let reopened = parsed(&bytes);
    let sheet = &reopened["payload"]["data"]["sheets"][0];
    assert_eq!(sheet["sourceAutoFilter"]["children"][0]["children"][0]["children"][0]["attributes"]["val"], "A*B?~~");
    let edited = rewrite(&bytes, |xml| xml.replace("val=\"A*B?~~\"", "val=\"C*D?~~\""));
    let reopened = parsed(&edited);
    assert_eq!(reopened["payload"]["data"]["sheets"][0]["autoFilter"]["columns"]["1"]["value"], "C*D?~");
}

#[test]
fn native_escaped_literal_wildcards_keep_source_and_visibility_without_becoming_a_match_all_predicate() {
    let mut sheet = source();
    sheet["autoFilter"]["columns"]["1"] = json!({ "kind": "condition", "operator": "equals", "value": "A*B", "caseSensitive": false });
    let edited = rewrite(&export(sheet), |xml| xml.replace("val=\"A*B\"", "val=\"A~*B\""));
    let reopened = parsed(&edited); let sheet = &reopened["payload"]["data"]["sheets"][0];
    assert!(sheet.get("autoFilter").is_none());
    assert_eq!(sheet["sourceAutoFilter"]["children"][0]["children"][0]["children"][0]["attributes"]["val"], "A~*B");
    assert_eq!(sheet["hiddenRows"], json!([2, 3, 5]));
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

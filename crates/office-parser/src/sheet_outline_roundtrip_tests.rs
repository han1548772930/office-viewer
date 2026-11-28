//! Real ZIP/SpreadsheetML contracts for the deferred unified verification pass.
//! These drafts have not been executed during implementation.

use std::io::{Cursor, Read, Write};

use office_io_common::conditional_format_xml::{Element, MAIN_NAMESPACE, STRICT_MAIN_NAMESPACE, parse};
use office_io_common::sheet_outline_source::{SHEET_OUTLINE_EXTENSION_URI, SHEET_OUTLINE_NAMESPACE};
use serde_json::{Value, json};

use crate::parser::parse_office_bytes;

fn source() -> Value {
    json!({ "id": "outline", "name": "Outline & Data", "axisUnit": "cssPixel", "rowCount": 30, "columnCount": 12,
        "defaultRowHeight": 26, "defaultColumnWidth": 112, "rowHeights": { "7": 0, "12": 32.5 },
        "hiddenRows": [4, 15], "hiddenColumns": [3, 10],
        "rowHiddenSpans": [{ "start": 17, "end": 19, "value": true }],
        "columnHiddenSpans": [{ "start": 7, "end": 8, "value": true }],
        "cells": { "0:0": { "row": 0, "col": 0, "value": "Header" }, "1:0": { "row": 1, "col": 0, "value": 42 } },
        "outline": { "summaryAbove": false, "summaryLeft": true, "futureState": { "v": [1, 2] }, "groups": [
            { "id": "outer", "axis": "row", "start": 1, "end": 8, "collapsed": true, "futureGroup": { "label": "A & B" } },
            { "id": "inner", "axis": "row", "start": 3, "end": 5, "collapsed": false },
            { "id": "columns", "axis": "column", "start": 2, "end": 4, "collapsed": true } ] } })
}
fn export(sheet: Value) -> Vec<u8> { office_writer::export_xlsx_bytes(&json!({ "kind": "workbook", "sheets": [sheet] }).to_string()).unwrap().0 }
fn reopened(bytes: &[u8]) -> Value {
    let envelope = serde_json::to_value(parse_office_bytes(bytes, "outline-roundtrip.xlsx").unwrap()).unwrap();
    envelope["payload"]["data"]["sheets"][0].clone()
}
fn part(bytes: &[u8], path: &str) -> String {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut xml = String::new(); archive.by_name(path).unwrap().read_to_string(&mut xml).unwrap(); xml
}
fn rewrite(bytes: &[u8], edit: impl Fn(&str, Vec<u8>) -> (String, Vec<u8>)) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap(); let name = entry.name().unwrap().into_owned();
        let mut content = Vec::new(); entry.read_to_end(&mut content).unwrap();
        let (name, content) = edit(&name, content); writer.start_file(name, options).unwrap(); writer.write_all(&content).unwrap();
    }
    writer.finish().unwrap().into_inner()
}
fn edit_sheet(bytes: &[u8], edit: impl Fn(&mut Element)) -> Vec<u8> {
    rewrite(bytes, |name, content| {
        if name != "xl/worksheets/sheet1.xml" { return (name.to_owned(), content); }
        let mut sheet = parse(std::str::from_utf8(&content).unwrap()).unwrap(); edit(&mut sheet);
        (name.to_owned(), sheet.render().into_bytes())
    })
}
fn row_mut(sheet: &mut Element, number: u32) -> &mut Element {
    sheet.child_mut("sheetData").unwrap().children.iter_mut().find(|row| row.attrs.get("r") == Some(&number.to_string())).unwrap()
}
fn reexport_payload(payload: &Value) -> Vec<u8> {
    let mut sheet = source();
    for key in ["outline", "hiddenRows", "hiddenColumns", "rowHiddenSpans", "columnHiddenSpans", "rowCount", "columnCount"] { sheet[key] = payload[key].clone(); }
    sheet["metadata"] = json!({ "outlineOfficeSource": payload["sourceOutline"] });
    export(sheet)
}

#[test]
fn native_levels_markers_and_exact_ids_unknown_fields_hidden_overlaps_reopen() {
    let input = source(); let bytes = export(input.clone()); let xml = part(&bytes, "xl/worksheets/sheet1.xml");
    let tree = parse(&xml).unwrap();
    assert_eq!(tree.children.iter().filter(|node| node.is_main("sheetPr")).count(), 1);
    let outline_pr = tree.child("sheetPr").unwrap().child("outlinePr").unwrap();
    assert_eq!(outline_pr.attrs.get("summaryBelow").map(String::as_str), Some("1"));
    assert_eq!(outline_pr.attrs.get("summaryRight").map(String::as_str), Some("0"));
    assert_eq!(tree.child("sheetFormatPr").unwrap().attrs.get("outlineLevelRow").map(String::as_str), Some("2"));
    assert!(xml.contains(SHEET_OUTLINE_NAMESPACE));
    let parsed = reopened(&bytes);
    for key in ["outline", "hiddenRows", "hiddenColumns", "rowHiddenSpans", "columnHiddenSpans", "rowCount", "columnCount"] { assert_eq!(parsed[key], input[key], "{key}"); }
    assert_eq!(parsed["sourceOutline"]["projectOutlineExtension"]["hiddenReasonsExact"], true);
    assert_eq!(parsed["sourceOutline"]["manualVisibilityNormalized"], true);
    assert_eq!(reopened(&reexport_payload(&parsed))["outline"], input["outline"]);
}

#[test]
fn native_empty_detail_rows_and_column_runs_are_emitted_beyond_the_used_cells() {
    let bytes = export(source()); let tree = parse(&part(&bytes, "xl/worksheets/sheet1.xml")).unwrap();
    let rows = &tree.child("sheetData").unwrap().children;
    let row = rows.iter().find(|row| row.attrs.get("r").map(String::as_str) == Some("6")).unwrap();
    assert_eq!(row.attrs.get("outlineLevel").map(String::as_str), Some("2")); assert!(row.children.is_empty());
    let marker = rows.iter().find(|row| row.attrs.get("r").map(String::as_str) == Some("10")).unwrap();
    assert_eq!(marker.attrs.get("collapsed").map(String::as_str), Some("1"));
    assert!(tree.child("cols").unwrap().children.iter().any(|column| column.attrs.get("outlineLevel").map(String::as_str) == Some("1")));
}

#[test]
fn external_cell_and_geometry_edits_keep_the_independent_outline_snapshot() {
    let bytes = export(source());
    let content = edit_sheet(&bytes, |sheet| row_mut(sheet, 2).children[0].child_mut("v").unwrap().set_text("99"));
    let parsed = reopened(&content); assert_eq!(parsed["outline"], source()["outline"]);
    let resized = edit_sheet(&bytes, |sheet| row_mut(sheet, 13).set_attr("ht", "30"));
    let parsed = reopened(&resized); assert_eq!(parsed["outline"], source()["outline"]);
    assert_eq!(parsed["hiddenRows"], source()["hiddenRows"]);
    assert_eq!(parsed["sourceOutline"]["projectOutlineExtension"]["visibilityProjectionMatched"], true);
    assert_eq!(parsed["sourceAxisDimensions"]["projectAxisExtension"]["restored"], false);
}

#[test]
fn external_native_level_edit_replaces_stale_group_ids_without_losing_manual_state() {
    let bytes = export(source());
    let edited = edit_sheet(&bytes, |sheet| {
        for row in &mut sheet.child_mut("sheetData").unwrap().children {
            if row.attrs.get("outlineLevel").map(String::as_str) == Some("2") { row.set_attr("outlineLevel", "1"); }
        }
    });
    let parsed = reopened(&edited);
    assert_eq!(parsed["sourceOutline"]["projectOutlineExtension"]["restored"], false);
    assert!(parsed["outline"]["groups"].as_array().unwrap().iter().all(|group| group["id"] != "inner" && group["id"] != "outer"));
    assert_eq!(parsed["hiddenRows"], source()["hiddenRows"]);
    assert!(parsed["sourceOutline"]["projectSnapshot"]["state"]["groups"].as_array().unwrap().iter().any(|group| group["id"] == "inner"));
}

#[test]
fn external_hidden_edit_keeps_model_but_does_not_restore_stale_manual_hides() {
    let bytes = export(source()); let edited = edit_sheet(&bytes, |sheet| row_mut(sheet, 16).set_attr("hidden", "0"));
    let parsed = reopened(&edited);
    assert_eq!(parsed["outline"], source()["outline"]);
    assert_eq!(parsed["sourceOutline"]["projectOutlineExtension"]["restored"], true);
    assert_eq!(parsed["sourceOutline"]["projectOutlineExtension"]["visibilityProjectionMatched"], false);
    assert_eq!(parsed["sourceOutline"]["projectOutlineExtension"]["hiddenReasonsExact"], false);
    assert!(!parsed["hiddenRows"].as_array().unwrap().contains(&json!(15)));
}

#[test]
fn native_only_file_inferrs_groups_and_marks_unrecoverable_hidden_overlap() {
    let bytes = export(source()); let native = edit_sheet(&bytes, |sheet| sheet.replace_children(|node| node.is_main("extLst"), Vec::new()));
    let parsed = reopened(&native);
    assert_eq!(parsed["sourceOutline"]["projectOutlineExtension"]["restored"], false);
    assert_eq!(parsed["sourceOutline"]["projectOutlineExtension"]["hiddenReasonsExact"], false);
    assert!(parsed["outline"]["groups"].as_array().unwrap().iter().any(|group| group["axis"] == "row" && group["start"] == 1 && group["end"] == 8 && group["collapsed"] == true));
    assert!(!parsed["hiddenRows"].as_array().unwrap().contains(&json!(4)));
    assert!(parsed["sourceOutline"]["native"]["diagnostics"].as_array().unwrap().iter().any(|value| value.as_str().unwrap().starts_with("NATIVE_HIDDEN_REASONS_AMBIGUOUS")));
}

#[test]
fn unknown_native_outline_xml_and_axis_attributes_survive_second_binary_export() {
    let bytes = export(source()); let edited = edit_sheet(&bytes, |sheet| {
        let outline = sheet.child_mut("sheetPr").unwrap().child_mut("outlinePr").unwrap();
        outline.set_attr_ns("urn:future:outline", "f:policy", "f:team"); outline.set_attr("xmlns:f", "urn:future:outline");
        let mut node = Element::prefixed("urn:future:outline", "f", "extension"); node.set_attr("key", "A & B"); outline.push(node);
        let row = row_mut(sheet, 2); row.set_attr_ns("urn:future:outline", "f:rowFlag", "future");
    });
    let first = reopened(&edited); let bytes = reexport_payload(&first); let second = reopened(&bytes);
    assert_eq!(second["outline"], first["outline"]);
    let tree = parse(&part(&bytes, "xl/worksheets/sheet1.xml")).unwrap();
    let outline = tree.child("sheetPr").unwrap().child("outlinePr").unwrap();
    assert_eq!(outline.attribute_ns("urn:future:outline", "policy"), Some("f:team"));
    assert!(outline.children.iter().any(|node| node.namespace == "urn:future:outline" && node.attrs.get("key").map(String::as_str) == Some("A & B")));
    let row = tree.child("sheetData").unwrap().children.iter().find(|row| row.attrs.get("r").map(String::as_str) == Some("2")).unwrap();
    assert_eq!(row.attribute_ns("urn:future:outline", "rowFlag"), Some("future"));
}

#[test]
fn eight_equal_layers_and_whole_axis_boundary_groups_keep_full_model_and_extents() {
    let mut sheet = source();
    sheet["outline"]["summaryAbove"] = json!(true); sheet["outline"]["summaryLeft"] = json!(false);
    sheet["outline"]["groups"] = json!((0..8).map(|index| json!({ "id": format!("layer-{index}"), "axis": "row", "start": 0, "end": 29, "collapsed": true })).collect::<Vec<_>>());
    let (bytes, warnings) = office_writer::export_xlsx_bytes(&json!({ "kind": "workbook", "sheets": [sheet.clone()] }).to_string()).unwrap();
    assert!(warnings.iter().any(|warning| warning.message.contains("OUTLINE_DEPTH_SNAPSHOT")));
    assert!(warnings.iter().any(|warning| warning.message.contains("OUTLINE_BOUNDARY_SNAPSHOT")));
    let parsed = reopened(&bytes); assert_eq!(parsed["outline"], sheet["outline"]); assert_eq!(parsed["rowCount"], 30);
    let tree = parse(&part(&bytes, "xl/worksheets/sheet1.xml")).unwrap();
    assert_eq!(tree.child("sheetData").unwrap().children.len(), 30);
    assert!(tree.child("sheetData").unwrap().children.iter().all(|row| row.attrs.get("outlineLevel").map(String::as_str) == Some("7")));
}

#[test]
fn outline_and_filter_share_one_sheet_pr_and_retain_independent_hidden_causes() {
    let mut sheet = source();
    sheet["autoFilter"] = json!({ "range": { "startRow": 0, "endRow": 20, "startCol": 0, "endCol": 0 },
        "columns": { "0": { "kind": "values", "selectedItems": [42] } }, "filteredRows": [4, 15, 20] });
    let bytes = export(sheet.clone()); let tree = parse(&part(&bytes, "xl/worksheets/sheet1.xml")).unwrap();
    let properties = tree.children.iter().filter(|node| node.is_main("sheetPr")).collect::<Vec<_>>();
    assert_eq!(properties.len(), 1); assert_eq!(properties[0].attrs.get("filterMode").map(String::as_str), Some("1"));
    assert!(properties[0].child("outlinePr").is_some());
    let parsed = reopened(&bytes); assert_eq!(parsed["hiddenRows"], source()["hiddenRows"]);
    assert_eq!(parsed["autoFilter"]["filteredRows"], sheet["autoFilter"]["filteredRows"]);
}

#[test]
fn duplicate_and_foreign_project_extensions_fail_without_returning_incomplete_state() {
    let bytes = export(source()); let duplicate = edit_sheet(&bytes, |sheet| {
        let extensions = sheet.child_mut("extLst").unwrap();
        let candidate = extensions.children.iter().find(|node| node.attrs.get("uri").map(String::as_str) == Some(SHEET_OUTLINE_EXTENSION_URI)).unwrap().clone();
        extensions.push(candidate);
    });
    assert!(parse_office_bytes(&duplicate, "duplicate-outline.xlsx").is_err());
    let foreign = edit_sheet(&bytes, |sheet| {
        let extension = sheet.child_mut("extLst").unwrap().children.iter_mut().find(|node| node.attrs.get("uri").map(String::as_str) == Some(SHEET_OUTLINE_EXTENSION_URI)).unwrap();
        extension.children[0].namespace = "urn:foreign:outline".into(); extension.children[0].set_attr("xmlns:oo", "urn:foreign:outline");
    });
    assert!(parse_office_bytes(&foreign, "foreign-outline.xlsx").is_err());
}

#[test]
fn relocated_worksheet_parts_and_strict_namespace_are_resolved_through_opc_relationships() {
    let bytes = export(source());
    let relocated = rewrite(&bytes, |path, data| {
        let name = if path == "xl/worksheets/sheet1.xml" { "xl/alternate/outline.xml" } else { path };
        let text = std::str::from_utf8(&data).unwrap();
        let text = if path == "xl/_rels/workbook.xml.rels" { text.replace("worksheets/sheet1.xml", "alternate/outline.xml") }
            else if path == "[Content_Types].xml" { text.replace("/xl/worksheets/sheet1.xml", "/xl/alternate/outline.xml") } else { text.to_owned() };
        (name.to_owned(), text.into_bytes())
    });
    assert_eq!(reopened(&relocated)["outline"], source()["outline"]);
    let strict = rewrite(&bytes, |path, data| {
        let text = std::str::from_utf8(&data).unwrap().replace(MAIN_NAMESPACE, STRICT_MAIN_NAMESPACE)
            .replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships");
        (path.to_owned(), text.into_bytes())
    });
    // Test the namespace-aware outline sidecar independently from rxls' Strict
    // cell parser; full Strict Office file parsing is a separate unverified gate.
    let mut base = parse_office_bytes(&bytes, "base.xlsx").unwrap();
    if let crate::model::DocumentModel::Spreadsheet(ref mut workbook) = base.payload {
        let mut outline = crate::sheet_outline_sidecar::SheetOutlineSidecar::read(&strict);
        workbook.sheets[0].outline = None; outline.restore(&mut workbook.sheets[0]);
        assert_eq!(serde_json::to_value(&workbook.sheets[0].outline).unwrap(), source()["outline"]);
    } else { panic!("fixture is not a workbook"); }
}

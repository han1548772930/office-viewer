//! Native binary XLSX fixtures for the deferred common verification pass.

use std::io::{Cursor, Read, Write};

use serde_json::{Value, json};

use crate::parser::parse_office_bytes;

fn export(sheet: Value) -> Vec<u8> {
    let (bytes, warnings) = office_writer::export_xlsx_bytes(&json!({ "kind": "workbook", "sheets": [sheet] }).to_string()).unwrap();
    assert!(warnings.is_empty(), "{warnings:?}");
    bytes
}
fn parsed(bytes: &[u8]) -> Value { serde_json::to_value(parse_office_bytes(bytes, "axis-roundtrip.xlsx").unwrap()).unwrap() }
fn part(bytes: &[u8], path: &str) -> String {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut xml = String::new();
    archive.by_name(path).unwrap().read_to_string(&mut xml).unwrap();
    xml
}
fn rewrite(bytes: &[u8], edit: impl Fn(&str, String) -> (String, String)) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap();
        let name = entry.name().unwrap().into_owned();
        let mut xml = String::new();
        entry.read_to_string(&mut xml).unwrap();
        let (name, xml) = edit(&name, xml);
        writer.start_file(name, options).unwrap();
        writer.write_all(xml.as_bytes()).unwrap();
    }
    writer.finish().unwrap().into_inner()
}

fn edit_axis_state(xml: String, edit: impl Fn(&mut Value)) -> String {
    let tag = xml.find("<ov:axisDimensions ").unwrap();
    let start = tag + xml[tag..].find('>').unwrap() + 1;
    let end = start + xml[start..].find("</ov:axisDimensions>").unwrap();
    let decoded = quick_xml::escape::unescape(&xml[start..end]).unwrap();
    let mut state: Value = serde_json::from_str(&decoded).unwrap();
    edit(&mut state);
    let state = state.to_string();
    format!("{}{}{}", &xml[..start], quick_xml::escape::escape(&state), &xml[end..])
}

fn hidden_span_source() -> Value {
    json!({ "id": "hidden-spans", "name": "Hidden axes", "axisUnit": "cssPixel", "rowCount": 40, "columnCount": 12,
        "defaultRowHeight": 26, "defaultColumnWidth": 112, "hiddenRows": [10], "hiddenColumns": [8],
        "rowHiddenSpans": [{ "start": 1, "end": 3, "value": true }],
        "columnHiddenSpans": [{ "start": 2, "end": 4, "value": true }],
        "cells": { "0:0": { "row": 0, "col": 0, "value": 1 } } })
}

#[test]
fn exact_pixels_modes_zero_sizes_hidden_reasons_and_sparklines_survive_binary_roundtrip() {
    let source = json!({
        "id": "axes", "name": "Axis & Data", "axisUnit": "cssPixel", "rowCount": 400, "columnCount": 40,
        "defaultRowHeight": 26.125, "defaultColumnWidth": 112.875, "defaultRowAutoHeight": true,
        "rowHeights": { "7": 31.123456789 }, "columnWidths": { "9": 14.125 },
        "rowHeightSpans": [{ "start": 11, "end": 15, "value": 0 }],
        "columnWidthSpans": [{ "start": 2, "end": 4, "value": 0 }],
        "rowAutoHeightSpans": [{ "start": 4, "end": 4, "value": true }, { "start": 11, "end": 15, "value": false }],
        "rowAutoHeights": { "4": 48.625 }, "hiddenRows": [10], "hiddenColumns": [6],
        "cells": { "0:1": { "row": 0, "col": 1, "value": 2 }, "0:2": { "row": 0, "col": 2, "value": 4 },
            "0:5": { "row": 0, "col": 5, "value": null, "sparkline": { "type": "line", "sourceRange": "B1:C1", "color": "#4472C4", "emptyCells": "gap", "axis": "individual" } } }
    });
    let bytes = export(source);
    let xml = part(&bytes, "xl/worksheets/sheet1.xml");
    assert_eq!(xml.matches("<extLst>").count(), 1);
    assert!(xml.contains("<x14:sparklineGroups"));
    assert!(xml.contains("<ov:axisDimensions"));
    assert!(xml.contains("ht=\"36.4688\" customHeight=\"0\""));
    let reopened = parsed(&bytes);
    let sheet = &reopened["payload"]["data"]["sheets"][0];
    assert_eq!(sheet["axisUnit"], "cssPixel");
    assert_eq!(sheet["rowCount"], 400);
    assert_eq!(sheet["columnCount"], 40);
    assert_eq!(sheet["defaultRowHeight"], 26.125);
    assert_eq!(sheet["defaultColumnWidth"], 112.875);
    assert_eq!(sheet["rowHeights"], json!([{ "index": 7, "size": 31.123456789 }]));
    assert_eq!(sheet["columnWidths"], json!([{ "index": 9, "size": 14.125 }]));
    assert_eq!(sheet["rowAutoHeights"], json!([{ "index": 4, "size": 48.625 }]));
    assert_eq!(sheet["rowHeightSpans"], json!([{ "start": 11, "end": 15, "value": 0 }]));
    assert_eq!(sheet["hiddenRows"], json!([10]));
    assert_eq!(sheet["hiddenColumns"], json!([6]));
    assert!(sheet["cells"].as_array().unwrap().iter().any(|cell| cell["sparkline"]["sourceRange"] == "B1:C1"));
    assert_eq!(sheet["sourceAxisDimensions"]["projectAxisExtension"]["projectionMatched"], true);
}

#[test]
fn whole_axis_manual_height_and_exceptions_are_compact_but_original_defaults_reopen() {
    let bytes = export(json!({ "id": "full", "name": "Full axis", "axisUnit": "cssPixel", "rowCount": 1_048_576, "columnCount": 26,
        "defaultRowHeight": 20.125, "defaultRowAutoHeight": true,
        "rowHeightSpans": [{ "start": 0, "end": 1_048_575, "value": 32.125 }],
        "rowAutoHeightSpans": [{ "start": 0, "end": 1_048_575, "value": false }], "rowHeights": { "9": 40.375 } }));
    let xml = part(&bytes, "xl/worksheets/sheet1.xml");
    assert_eq!(xml.matches("<row ").count(), 1);
    assert!(xml.len() < 8_000);
    assert!(xml.contains("defaultRowHeight=\"24.0938\" customHeight=\"1\""));
    let reopened = parsed(&bytes);
    let sheet = &reopened["payload"]["data"]["sheets"][0];
    assert_eq!(sheet["defaultRowHeight"], 20.125);
    assert_eq!(sheet["defaultRowAutoHeight"], true);
    assert_eq!(sheet["rowHeightSpans"][0]["end"], 1_048_575);
    assert_eq!(sheet["rowHeightSpans"][0]["value"], 32.125);
    assert_eq!(sheet["rowHeights"][0]["size"], 40.375);
}

#[test]
fn stale_projection_is_not_applied_after_external_resize_and_content_edits_keep_it() {
    let bytes = export(json!({ "name": "Edited", "axisUnit": "cssPixel", "rowCount": 300,
        "rowHeights": { "3": 32.5 }, "cells": { "0:0": { "row": 0, "col": 0, "value": 1 } } }));
    let content_only = rewrite(&bytes, |path, xml| (path.to_owned(), if path == "xl/worksheets/sheet1.xml" { xml.replace("<v>1</v>", "<v>9</v>") } else { xml }));
    let sheet = &parsed(&content_only)["payload"]["data"]["sheets"][0];
    assert_eq!(sheet["rowHeights"][0]["size"], 32.5);
    assert_eq!(sheet["rowCount"], 300);
    let resized = rewrite(&bytes, |path, xml| (path.to_owned(), if path == "xl/worksheets/sheet1.xml" { xml.replace("ht=\"24.375\"", "ht=\"30\"") } else { xml }));
    let reopened = parsed(&resized);
    assert_eq!(reopened["payload"]["data"]["sheets"][0]["rowHeights"][0]["size"], 40.0);
    assert!(reopened["diagnostics"]["warnings"].as_array().unwrap().iter().any(|warning| warning.as_str().is_some_and(|warning| warning.contains("changed outside"))));
}

#[test]
fn relocated_worksheet_parts_and_explicit_visible_rows_under_zero_defaults_roundtrip() {
    let bytes = export(json!({ "name": "Moved", "axisUnit": "cssPixel", "rowCount": 50, "columnCount": 20,
        "defaultRowHeight": 0, "defaultRowAutoHeight": false, "rowHeights": { "2": 30.5 }, "hiddenRows": [3] }));
    let xml = part(&bytes, "xl/worksheets/sheet1.xml");
    assert!(xml.contains("zeroHeight=\"1\""));
    assert!(xml.contains("r=\"3\" ht=\"22.875\" customHeight=\"1\" hidden=\"0\""));
    let moved = rewrite(&bytes, |path, xml| {
        let renamed = if path == "xl/worksheets/sheet1.xml" { "custom/axes.xml" } else { path };
        let xml = match path {
            "xl/_rels/workbook.xml.rels" => xml.replace("Target=\"worksheets/sheet1.xml\"", "Target=\"../custom/axes.xml\""),
            "[Content_Types].xml" => xml.replace("/xl/worksheets/sheet1.xml", "/custom/axes.xml"),
            _ => xml,
        };
        (renamed.to_owned(), xml)
    });
    let sheet = &parsed(&moved)["payload"]["data"]["sheets"][0];
    assert_eq!(sheet["defaultRowHeight"], 0);
    assert_eq!(sheet["hiddenRows"], json!([3]));
    assert_eq!(sheet["rowHeights"], json!([{ "index": 2, "size": 30.5 }]));
}

#[test]
fn hidden_spans_emit_standard_rows_columns_and_reopen_as_compact_manual_reasons() {
    let bytes = export(hidden_span_source());
    let xml = part(&bytes, "xl/worksheets/sheet1.xml");
    for row in [2, 3, 4, 11] {
        assert!(xml.contains(&format!("<row r=\"{row}\" hidden=\"1\">")));
    }
    assert!(xml.contains("<col min=\"3\" max=\"5\" width=\"16\" customWidth=\"1\" hidden=\"1\""));
    let reopened = parsed(&bytes);
    let sheet = &reopened["payload"]["data"]["sheets"][0];
    assert_eq!(sheet["rowHiddenSpans"], json!([{ "start": 1, "end": 3, "value": true }]));
    assert_eq!(sheet["columnHiddenSpans"], json!([{ "start": 2, "end": 4, "value": true }]));
    assert_eq!(sheet["hiddenRows"], json!([10]));
    assert_eq!(sheet["hiddenColumns"], json!([8]));
    assert_eq!(sheet["sourceAxisDimensions"]["projectAxisExtension"]["projectionMatched"], true);

    let standard_only = rewrite(&bytes, |path, xml| {
        let xml = if path == "xl/worksheets/sheet1.xml" {
            let marker = format!("<ext uri=\"{}\">", office_io_common::SHEET_AXIS_EXTENSION_URI);
            let start = xml.find(&marker).unwrap();
            let end = start + xml[start..].find("</ext>").unwrap() + "</ext>".len();
            format!("{}{}", &xml[..start], &xml[end..])
        } else { xml };
        (path.to_owned(), xml)
    });
    let reopened = parsed(&standard_only);
    let sheet = &reopened["payload"]["data"]["sheets"][0];
    assert_eq!(sheet["hiddenRows"], json!([1, 2, 3, 10]));
    assert_eq!(sheet["hiddenColumns"], json!([2, 3, 4, 8]));
    assert_eq!(sheet["rowHiddenSpans"], json!([]));
    assert_eq!(sheet["columnHiddenSpans"], json!([]));
}

#[test]
fn older_axis_extensions_without_hidden_span_fields_still_restore_legacy_hidden_points() {
    let bytes = export(json!({ "name": "Legacy hides", "axisUnit": "cssPixel", "rowCount": 40, "columnCount": 12,
        "hiddenRows": [1, 3], "hiddenColumns": [2, 4] }));
    let legacy = rewrite(&bytes, |path, xml| (path.to_owned(), if path == "xl/worksheets/sheet1.xml" {
        edit_axis_state(xml, |state| {
            state.as_object_mut().unwrap().remove("rowHiddenSpans");
            state.as_object_mut().unwrap().remove("columnHiddenSpans");
        })
    } else { xml }));
    let reopened = parsed(&legacy);
    let sheet = &reopened["payload"]["data"]["sheets"][0];
    assert_eq!(sheet["hiddenRows"], json!([1, 3]));
    assert_eq!(sheet["hiddenColumns"], json!([2, 4]));
    assert_eq!(sheet["rowHiddenSpans"], json!([]));
    assert_eq!(sheet["columnHiddenSpans"], json!([]));
    assert_eq!(sheet["sourceAxisDimensions"]["projectAxisExtension"]["projectionMatched"], true);
}

#[test]
fn external_unhide_takes_precedence_over_stale_project_hidden_spans() {
    let bytes = export(hidden_span_source());
    let edited = rewrite(&bytes, |path, xml| (path.to_owned(), if path == "xl/worksheets/sheet1.xml" {
        xml.replace("<row r=\"2\" hidden=\"1\">", "<row r=\"2\" hidden=\"0\">")
            .replace("<col min=\"3\" max=\"5\" width=\"16\" customWidth=\"1\" hidden=\"1\"",
                "<col min=\"3\" max=\"5\" width=\"16\" customWidth=\"1\" hidden=\"0\"")
    } else { xml }));
    assert_ne!(part(&bytes, "xl/worksheets/sheet1.xml"), part(&edited, "xl/worksheets/sheet1.xml"));
    let reopened = parsed(&edited);
    let sheet = &reopened["payload"]["data"]["sheets"][0];
    assert_eq!(sheet["hiddenRows"], json!([2, 3, 10]));
    assert_eq!(sheet["hiddenColumns"], json!([8]));
    assert_eq!(sheet["rowHiddenSpans"], json!([]));
    assert_eq!(sheet["columnHiddenSpans"], json!([]));
    assert!(reopened["diagnostics"]["warnings"].as_array().unwrap().iter().any(|warning|
        warning.as_str().is_some_and(|text| text.contains("changed outside"))));
}

#[test]
fn manual_hidden_spans_and_filter_overlaps_remain_independent_even_after_external_resize() {
    let mut source = hidden_span_source();
    source["hiddenRows"] = json!([6]);
    source["rowHiddenSpans"] = json!([{ "start": 3, "end": 4, "value": true }]);
    source["rowHeights"] = json!({ "9": 32.5 });
    source["autoFilter"] = json!({ "range": { "startRow": 0, "endRow": 7, "startCol": 0, "endCol": 0 },
        "columns": { "0": { "kind": "values", "selectedItems": [1] } }, "filteredRows": [2, 3, 4, 5, 6, 7] });
    for row in 1..=7 {
        source["cells"][format!("{row}:0")] = json!({ "row": row, "col": 0, "value": row });
    }
    let bytes = export(source);
    let reopened = parsed(&bytes);
    let sheet = &reopened["payload"]["data"]["sheets"][0];
    assert_eq!(sheet["hiddenRows"], json!([6]));
    assert_eq!(sheet["rowHiddenSpans"], json!([{ "start": 3, "end": 4, "value": true }]));
    assert_eq!(sheet["autoFilter"]["filteredRows"], json!([2, 3, 4, 5, 6, 7]));
    let resized = rewrite(&bytes, |path, xml| (path.to_owned(), if path == "xl/worksheets/sheet1.xml" {
        xml.replace("ht=\"24.375\"", "ht=\"30\"")
    } else { xml }));
    let reopened = parsed(&resized);
    let sheet = &reopened["payload"]["data"]["sheets"][0];
    assert_eq!(sheet["hiddenRows"], json!([3, 4, 6]));
    assert_eq!(sheet["rowHiddenSpans"], json!([]));
    assert_eq!(sheet["autoFilter"]["filteredRows"], json!([2, 3, 4, 5, 6, 7]));
    assert_eq!(sheet["sourceAxisDimensions"]["projectAutoFilterExtension"]["restored"], true);
    assert_eq!(sheet["sourceAxisDimensions"]["projectAutoFilterExtension"]["hiddenReasonsExact"], false);
}

#[test]
fn hidden_records_outside_declared_axis_counts_are_refused_and_invalid_private_records_do_not_replace_standard_hides() {
    for source in [
        json!({ "rowCount": 40, "hiddenRows": [40] }),
        json!({ "columnCount": 12, "hiddenColumns": [12] }),
        json!({ "rowCount": 40, "rowHiddenSpans": [{ "start": 39, "end": 40, "value": true }] }),
        json!({ "columnCount": 12, "columnHiddenSpans": [{ "start": 11, "end": 12, "value": true }] }),
    ] {
        let input = json!({ "kind": "workbook", "sheets": [source] }).to_string();
        assert!(office_writer::export_xlsx_bytes(&input).is_err());
    }
    let bytes = export(hidden_span_source());
    for hidden in [json!([40]), json!([1_048_576])] {
        let invalid = rewrite(&bytes, |path, xml| (path.to_owned(), if path == "xl/worksheets/sheet1.xml" {
            edit_axis_state(xml, |state| state["hiddenRows"] = hidden.clone())
        } else { xml }));
        let reopened = parsed(&invalid);
        let sheet = &reopened["payload"]["data"]["sheets"][0];
        assert_eq!(sheet["hiddenRows"], json!([1, 2, 3, 10]));
        assert_eq!(sheet["rowHiddenSpans"], json!([]));
        assert!(reopened["diagnostics"]["warnings"].as_array().unwrap().iter().any(|warning|
            warning.as_str().is_some_and(|text| text.contains("hidden indices") && text.contains("limit"))));
    }
}

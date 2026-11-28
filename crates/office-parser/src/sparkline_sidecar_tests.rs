//! Deferred ZIP, native metadata and parser/writer contracts; this batch did not execute them.
use super::*;
use std::io::{Read, Write};
use serde_json::{Value, json};

const SHEET: &str = include_str!("../tests/fixtures/sparklines/native-settings-sheet.xml");
const THEME: &str = include_str!("../tests/fixtures/sparklines/theme.xml");
const PALETTE: &str = include_str!("../tests/fixtures/sparklines/indexed-colors.xml");

fn package(sheet: &str, strict: bool) -> Vec<u8> {
    let main = if strict { xml::STRICT_MAIN_NAMESPACE } else { xml::MAIN_NAMESPACE };
    let rel = if strict { "http://purl.oclc.org/ooxml/officeDocument/relationships" } else { "http://schemas.openxmlformats.org/officeDocument/2006/relationships" };
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    let root = format!("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"root\" Type=\"{rel}/officeDocument\" Target=\"/data/books/main.xml\"/></Relationships>");
    let book = format!("<workbook xmlns=\"{main}\" xmlns:r=\"{rel}\"><sheets><sheet name=\"Trend\" sheetId=\"1\" r:id=\"sheet\"/></sheets></workbook>");
    let rels = format!("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"sheet\" Type=\"{rel}/worksheet\" Target=\"../sheets/trend.xml\"/><Relationship Id=\"theme\" Type=\"{rel}/theme\" Target=\"../resources/theme.xml\"/><Relationship Id=\"styles\" Type=\"{rel}/styles\" Target=\"../resources/styles.xml\"/></Relationships>");
    let sheet = sheet.replace(xml::MAIN_NAMESPACE, main); let palette = PALETTE.replace(xml::MAIN_NAMESPACE, main);
    for (name, source) in [("_rels/.rels", root.as_str()), ("data/books/main.xml", book.as_str()),
        ("data/books/_rels/main.xml.rels", rels.as_str()), ("data/sheets/trend.xml", sheet.as_str()),
        ("data/resources/theme.xml", THEME), ("data/resources/styles.xml", palette.as_str())] {
        writer.start_file(name, options).unwrap(); writer.write_all(source.as_bytes()).unwrap();
    }
    writer.finish().unwrap().into_inner()
}
fn simple_spec(source: &str) -> SparklinePayload {
    serde_json::from_value(json!({ "type": "line", "sourceRange": source, "color": "#246FCE", "emptyCells": "connect", "axis": "individual",
        "groupId": "author-group", "themeType": "custom", "points": {
            "highPoint": { "visible": true, "color": "#CC2244" }, "firstPoint": { "visible": true, "color": "#102030" },
            "lowPoint": { "visible": false, "color": "#556677" }, "lastPoint": { "visible": false, "color": "#334455" },
            "negativePoint": { "visible": false, "color": "rgba(0,0,255,0.5)" }, "markersPoint": { "visible": true, "color": "#AABBCC" } },
        "axisOptions": { "visible": true, "reverse": true, "color": "#010203" },
        "extremumMin": { "type": "group", "value": -99.25 }, "extremumMax": { "type": "custom", "value": 9.123456789123456 },
        "nonNumericCells": "zero", "containHiddenCells": true, "lineWidth": 2.25, "showGradient": true, "showRadius": true })).unwrap()
}
fn export(entries: &[(&str, SparklinePayload)]) -> Vec<u8> {
    let cells: serde_json::Map<String, Value> = entries.iter().map(|(target, spec)| ((*target).into(), json!({ "value": null, "sparkline": spec }))).collect();
    office_writer::export_xlsx_bytes(&json!({ "kind": "workbook", "sheets": [{ "id": "trend", "name": "Trend", "cells": cells }] }).to_string()).unwrap().0
}
fn rewrite(bytes: &[u8], edit: impl Fn(&str, Vec<u8>) -> Vec<u8>) -> Vec<u8> {
    let mut reader = ZipArchive::new(Cursor::new(bytes)).unwrap(); let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for index in 0..reader.len() {
        let mut entry = reader.by_index(index).unwrap(); let name = entry.name().unwrap().into_owned(); let mut data = Vec::new(); entry.read_to_end(&mut data).unwrap();
        writer.start_file(&name, options).unwrap(); writer.write_all(&edit(&name, data)).unwrap();
    }
    writer.finish().unwrap().into_inner()
}
fn edit_sheet(bytes: &[u8], edit: impl Fn(&mut Element)) -> Vec<u8> {
    rewrite(bytes, |name, data| {
        if name != "xl/worksheets/sheet1.xml" { return data; }
        let mut sheet = xml::parse(std::str::from_utf8(&data).unwrap()).unwrap(); edit(&mut sheet); sheet.render().into_bytes()
    })
}
fn native_group_mut(sheet: &mut Element) -> &mut Element {
    sheet.child_mut("extLst").unwrap().children.iter_mut().find(|ext| ext.attribute_ns("", "uri") == Some(io::NATIVE_EXTENSION_URI))
        .unwrap().child_mut("sparklineGroups").unwrap().children.first_mut().unwrap()
}

#[test]
fn relocated_opc_and_both_spreadsheet_namespaces_restore_all_native_members() {
    for strict in [false, true] {
        let mut sidecar = SparklineSidecar::read(&package(SHEET, strict)); assert_eq!(sidecar.failure, None);
        let first = sidecar.take("Trend", "G1").unwrap(); let second = sidecar.take("Trend", "$G$2").unwrap();
        assert_eq!(first.group_id, second.group_id); assert_eq!(first.source_range, "'O''Brien 数据!2026'!$B$1:E$1");
        assert_eq!(first.color, "#FF8080"); assert_eq!(first.axis_options.as_ref().unwrap().color.as_deref(), Some("#102030"));
        assert_eq!(first.line_width, Some(2.0 / 3.0)); assert_eq!(first.extremum_min.as_ref().unwrap().value, Some(-99.25));
        assert_eq!(first.extremum_max.as_ref().unwrap().value, Some(9.123456789123456));
        let remaining = sidecar.take_sheet("Trend"); assert_eq!(remaining.len(), 2); assert_eq!(remaining["K1"].empty_cells, "zero"); assert_eq!(remaining["K2"].sparkline_type, "winloss");
        assert!(sidecar.warnings.iter().any(|warning| warning.starts_with("SPARKLINE_DATE_AXIS_RETAINED")));
    }
}

#[test]
fn writer_to_real_parser_reopens_group_id_all_settings_and_blank_targets() {
    let first = simple_spec("$B$1:E$1"); let second = simple_spec("B2:$E2"); let bytes = export(&[("G1", first.clone()), ("G2", second.clone())]);
    let payload = serde_json::to_value(crate::parser::parse_office_bytes(&bytes, "sparkline-roundtrip.xlsx").unwrap()).unwrap();
    let cells = payload["payload"]["data"]["sheets"][0]["cells"].as_array().unwrap();
    for (row, expected) in [(0, first), (1, second)] {
        let raw = cells.iter().find(|cell| cell["row"] == row && cell["col"] == 6).unwrap();
        let restored: SparklinePayload = serde_json::from_value(raw["sparkline"].clone()).unwrap();
        let mut expected = expected; expected.office_source = restored.office_source.clone(); assert_eq!(restored, expected);
    }
}

#[test]
fn local_pie_and_native_group_can_reopen_and_be_exported_again_without_dropping_fields() {
    let mut pie = simple_spec("B1:E1"); pie.sparkline_type = "pie".into(); let line = simple_spec("B2:E2");
    let bytes = export(&[("G1", pie.clone()), ("G2", line.clone())]); let mut first = SparklineSidecar::read(&bytes);
    assert_eq!(first.failure, None); assert_eq!(first.take("Trend", "G1"), Some(pie.clone()));
    let restored = first.take("Trend", "G2").unwrap(); let second_bytes = export(&[("G1", pie.clone()), ("G2", restored.clone())]);
    let mut second = SparklineSidecar::read(&second_bytes); assert_eq!(second.failure, None);
    assert_eq!(second.take("Trend", "G1"), Some(pie)); assert_eq!(second.take("Trend", "G2").unwrap().non_numeric_cells, restored.non_numeric_cells);
}

#[test]
fn native_only_color_alpha_and_false_marker_flags_are_not_replaced_by_legacy_colors() {
    let mut sidecar = SparklineSidecar::read(&package(SHEET, false)); let first = sidecar.take("Trend", "G1").unwrap();
    assert_eq!(first.negative_color, None); assert_eq!(first.low_point_color, None);
    assert_eq!(first.points.as_ref().unwrap()["negativePoint"].visible, Some(false));
    let (_, alpha) = office_io_common::chart_color::rgba(first.points.as_ref().unwrap()["negativePoint"].color.as_deref().unwrap()).unwrap();
    assert!((alpha - 128.0 / 255.0).abs() < 1e-12);
    assert!(first.office_source.as_ref().unwrap().group_xml.contains("rgb=\"800000FF\""));
}

#[test]
fn externally_changed_native_settings_discard_only_the_stale_local_configuration() {
    let bytes = export(&[("G1", simple_spec("B1:E1")), ("G2", simple_spec("B2:E2"))]);
    let changed = edit_sheet(&bytes, |sheet| native_group_mut(sheet).set_attr("manualMax", "12"));
    let mut sidecar = SparklineSidecar::read(&changed); assert_eq!(sidecar.failure, None);
    let restored = sidecar.take("Trend", "G1").unwrap(); assert_ne!(restored.group_id.as_deref(), Some("author-group"));
    assert_eq!(restored.extremum_max.as_ref().unwrap().value, Some(12.0)); assert_eq!(restored.non_numeric_cells.as_deref(), Some("gap"));
    assert!(sidecar.warnings.iter().any(|warning| warning.starts_with("SPARKLINE_LOCAL_STALE")));
}

#[test]
fn externally_deleted_native_group_is_not_resurrected_from_its_old_local_record() {
    let bytes = export(&[("G1", simple_spec("B1:E1"))]);
    let deleted = edit_sheet(&bytes, |sheet| sheet.child_mut("extLst").unwrap().replace_children(|ext| ext.attribute_ns("", "uri") == Some(io::NATIVE_EXTENSION_URI), Vec::new()));
    let mut sidecar = SparklineSidecar::read(&deleted); assert_eq!(sidecar.failure, None); assert!(sidecar.take_sheet("Trend").is_empty());
    assert!(sidecar.warnings.iter().any(|warning| warning.starts_with("SPARKLINE_LOCAL_STALE")));
}

#[test]
fn externally_changed_source_and_group_membership_invalidate_the_old_group_projection() {
    let bytes = export(&[("G1", simple_spec("B1:E1")), ("G2", simple_spec("B2:E2"))]);
    let changed = edit_sheet(&bytes, |sheet| {
        let group = native_group_mut(sheet); let member = &mut group.child_mut("sparklines").unwrap().children[0];
        member.child_mut("f").unwrap().set_text("Trend!C1:E1"); member.child_mut("sqref").unwrap().set_text("H1");
    });
    let mut sidecar = SparklineSidecar::read(&changed); assert_eq!(sidecar.failure, None); assert!(sidecar.take("Trend", "G1").is_none());
    assert_eq!(sidecar.take("Trend", "H1").unwrap().source_range, "'Trend'!C1:E1");
    assert_ne!(sidecar.take("Trend", "G2").unwrap().group_id.as_deref(), Some("author-group"));
}

#[test]
fn external_theme_color_edits_win_over_the_stale_local_config() {
    let mut imported = SparklineSidecar::read(&package(SHEET, false)); let first = imported.take("Trend", "G1").unwrap();
    let bytes = export(&[("G1", first)]);
    let changed = rewrite(&bytes, |name, data| if name.ends_with("theme1.xml") {
        std::str::from_utf8(&data).unwrap().replace("val=\"FF0000\"", "val=\"00FF00\"").into_bytes()
    } else { data });
    let mut sidecar = SparklineSidecar::read(&changed); assert_eq!(sidecar.failure, None);
    assert_eq!(sidecar.take("Trend", "G1").unwrap().color, "#80FF80");
    assert!(sidecar.warnings.iter().any(|warning| warning.starts_with("SPARKLINE_LOCAL_STALE")));
}

#[test]
fn unrelated_external_theme_edits_refresh_source_resources_without_staling_author_settings() {
    let mut imported = SparklineSidecar::read(&package(SHEET, false)); let mut first = imported.take("Trend", "G1").unwrap();
    first.non_numeric_cells = Some("zero".into()); first.show_gradient = Some(true);
    let bytes = export(&[("G1", first)]);
    let changed = rewrite(&bytes, |name, data| if name.ends_with("theme1.xml") {
        std::str::from_utf8(&data).unwrap().replace("val=\"00FF00\"", "val=\"102030\"").into_bytes()
    } else { data });
    let mut sidecar = SparklineSidecar::read(&changed); assert_eq!(sidecar.failure, None);
    let restored = sidecar.take("Trend", "G1").unwrap(); assert_eq!(restored.color, "#FF8080"); assert_eq!(restored.non_numeric_cells.as_deref(), Some("zero"));
    assert_eq!(restored.show_gradient, Some(true)); assert!(restored.office_source.as_ref().unwrap().theme_xml.as_ref().unwrap().contains("val=\"102030\""));
    assert!(!sidecar.warnings.iter().any(|warning| warning.starts_with("SPARKLINE_LOCAL_STALE")));
}

#[test]
fn duplicate_native_targets_and_duplicate_local_records_fail_atomically() {
    let duplicate = SHEET.replace(">G2<", ">G1<"); let mut sidecar = SparklineSidecar::read(&package(&duplicate, false));
    assert!(sidecar.failure.is_some()); assert!(sidecar.take_sheet("Trend").is_empty());
    let bytes = export(&[("G1", simple_spec("B1:E1"))]); let changed = edit_sheet(&bytes, |sheet| {
        let local = sheet.child_mut("extLst").unwrap().children.iter_mut().find(|ext| ext.attribute_ns("", "uri") == Some(io::MODEL_EXTENSION_URI))
            .unwrap().child_mut("sparklines").unwrap(); local.push(local.children[0].clone());
    });
    let mut sidecar = SparklineSidecar::read(&changed); assert!(sidecar.failure.is_some()); assert!(sidecar.take_sheet("Trend").is_empty());
    assert!(crate::parser::parse_office_bytes(&changed, "duplicate-sparkline.xlsx").is_err());
}

#[test]
fn foreign_namespace_and_wrong_extension_uri_do_not_impersonate_native_sparklines() {
    for changed in [SHEET.replace(xml::X14_NAMESPACE, "urn:foreign:sparkline"), SHEET.replace(io::NATIVE_EXTENSION_URI, "{foreign}")] {
        let mut sidecar = SparklineSidecar::read(&package(&changed, false)); assert_eq!(sidecar.failure, None); assert!(sidecar.take_sheet("Trend").is_empty());
    }
}

#[test]
fn retained_native_metadata_and_date_axis_formula_survive_the_second_export() {
    let changed = SHEET.replace("dateAxis=\"1\"", "dateAxis=\"1\" futurePolicy=\"keep &amp; inspect\"");
    let mut sidecar = SparklineSidecar::read(&package(&changed, false)); assert_eq!(sidecar.failure, None);
    let first = sidecar.take("Trend", "G1").unwrap(); let bytes = export(&[("G1", first)]); let mut second = SparklineSidecar::read(&bytes);
    assert_eq!(second.failure, None); let source = second.take("Trend", "G1").unwrap().office_source.unwrap();
    let group = xml::parse(&source.group_xml).unwrap(); assert_eq!(group.attribute_ns("", "futurePolicy"), Some("keep & inspect"));
    assert_eq!(group.child("f").unwrap().text, "'O''Brien 数据!2026'!$A$1:$D$1");
}
